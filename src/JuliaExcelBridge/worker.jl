# JuliaExcelBridge worker
# JSON protocol over stdin/stdout, one request per line, one response per line.

import Pkg

try
    using DataFrames
catch e
    @warn "DataFrames not available" exception=(e, catch_backtrace())
end

# Ensure JSON is available, then load it at top level.
try
    Base.find_package("JSON")
catch
end

if Base.find_package("JSON") === nothing
    Pkg.add("JSON")
end

using JSON

const OBJECT_STORE = Dict{String, Any}()
const PLOTS_LOADED = Ref(false)

function ensure_plots_loaded()
    if PLOTS_LOADED[]
        return
    end

    # Prevent GR from opening a GUI window
    ENV["GKSwstype"] = "100"

    try
        Core.eval(Main, :(using Plots))
    catch e
        error("Plots.jl is not installed. Please run in Julia: using Pkg; Pkg.add(\"Plots\")")
    end

    PLOTS_LOADED[] = true
end

function make_ok(id, result)
    Dict("id" => id, "ok" => true, "result" => result)
end

function make_err(id, msg)
    Dict("id" => id, "ok" => false, "error" => string(msg))
end

is_scalar_json_value(x) = x === nothing || x isa Number || x isa String || x isa Bool

function is_json_matrix_like(x)
    x isa Vector || return false
    isempty(x) && return false
    all(row -> row isa Vector, x) || return false
    all(row -> !isempty(row), x) || return false
    ncols = length(x[1])
    all(row -> length(row) == ncols, x) || return false
    all(row -> all(is_scalar_json_value, row), x)
end

function common_scalar_type(values)
    if all(v -> v === nothing || v isa Bool, values)
        return :bool
    elseif all(v -> v === nothing || v isa Number, values)
        return :number
    elseif all(v -> v === nothing || v isa String, values)
        return :string
    else
        return :any
    end
end

function vector_from_json(values)
    t = common_scalar_type(values)
    if t == :bool
        return [v === nothing ? false : Bool(v) for v in values]
    elseif t == :number
        return [v === nothing ? NaN : Float64(v) for v in values]
    elseif t == :string
        return [v === nothing ? "" : String(v) for v in values]
    else
        return Any[v for v in values]
    end
end

function matrix_from_json(x)
    nr = length(x)
    nc = length(x[1])
    flat = collect(Iterators.flatten(x))
    t = common_scalar_type(flat)

    if t == :number
        data = Array{Float64}(undef, nr, nc)
        for r in 1:nr, c in 1:nc
            data[r, c] = x[r][c] === nothing ? NaN : Float64(x[r][c])
        end
        return data
    elseif t == :bool
        data = Array{Bool}(undef, nr, nc)
        for r in 1:nr, c in 1:nc
            data[r, c] = x[r][c] === nothing ? false : Bool(x[r][c])
        end
        return data
    elseif t == :string
        data = Matrix{String}(undef, nr, nc)
        for r in 1:nr, c in 1:nc
            data[r, c] = x[r][c] === nothing ? "" : String(x[r][c])
        end
        return data
    else
        data = Matrix{Any}(undef, nr, nc)
        for r in 1:nr, c in 1:nc
            data[r, c] = x[r][c]
        end
        return data
    end
end

function convert_json_value(x)
    if x === nothing
        return nothing
    elseif is_scalar_json_value(x)
        return x
    elseif is_json_matrix_like(x)
        return matrix_from_json(x)
    elseif x isa Vector
        converted = map(convert_json_value, x)
        if all(v -> v === nothing || is_scalar_json_value(v), converted)
            return vector_from_json(converted)
        else
            return converted
        end
    elseif x isa AbstractDict
        if get(x, "__jexcel_arg_type", nothing) == "jobj"
            nm = strip(String(get(x, "name", "")))
            isempty(nm) && error("JObj argument had a blank object name.")
            return resolve_object(nm)
        end
        return Dict(string(k) => convert_json_value(v) for (k, v) in pairs(x))
    else
        return x
    end
end

normalize_call_args(args) = args isa Vector ? map(convert_json_value, args) : [convert_json_value(args)]

function resolve_object(name::AbstractString)
    nm = String(name)
    if occursin(".", nm)
        parts = split(nm, ".")
        obj = if isdefined(Main, Symbol(parts[1]))
            getfield(Main, Symbol(parts[1]))
        elseif haskey(OBJECT_STORE, parts[1])
            OBJECT_STORE[parts[1]]
        elseif isdefined(Base, Symbol(parts[1]))
            getfield(Base, Symbol(parts[1]))
        else
            error("Object not found: " * parts[1])
        end

        for p in parts[2:end]
            obj = getproperty(obj, Symbol(p))
        end
        return obj
    end

    if haskey(OBJECT_STORE, nm)
        return OBJECT_STORE[nm]
    elseif isdefined(Main, Symbol(nm))
        return getfield(Main, Symbol(nm))
    elseif isdefined(Base, Symbol(nm))
        return getfield(Base, Symbol(nm))
    else
        error("Object not found: " * nm)
    end
end

function resolve_function(fun_name::AbstractString)
    obj = resolve_object(fun_name)
    obj isa Function || error("Function not callable: " * String(fun_name))
    return obj
end

function coerce_for_json(x)
    if x === nothing
        return nothing
    elseif x isa Number || x isa String || x isa Bool
        return x
    elseif x isa AbstractMatrix
        rows = Vector{Any}(undef, size(x, 1))
        for r in axes(x, 1)
            rows[r] = [coerce_for_json(x[r, c]) for c in axes(x, 2)]
        end
        return rows
    elseif x isa AbstractVector
        return [coerce_for_json(v) for v in x]
    elseif x isa Tuple
        return [coerce_for_json(v) for v in x]
    elseif x isa NamedTuple
        return Dict(String(k) => coerce_for_json(v) for (k, v) in pairs(x))
    elseif x isa AbstractDict
        return Dict(string(k) => coerce_for_json(v) for (k, v) in pairs(x))
    else
        return string(x)
    end
end

function safe_length(x)
    try
        return length(x)
    catch
        return 1
    end
end

function format_dim(x)
    if x isa AbstractArray
        return join(size(x), " x ")
    elseif is_dataframe_like(x)
        return string(nrow_df(x), " x ", ncol_df(x))
    else
        return "1 x 1"
    end
end

function object_summary_row(name, x)
    [name, string(typeof(x)), string(typeof(x)), string(safe_length(x)), format_dim(x)]
end

function object_describe_table(name, x)
    [
        ["Field", "Value"],
        ["Name", name],
        ["Class", string(typeof(x))],
        ["Type", string(typeof(x))],
        ["Length", string(safe_length(x))],
        ["Dimensions", format_dim(x)]
    ]
end

function render_plot_to_file(code, file, width=800, height=600)
    if file === nothing || isempty(file)
        error("Plot file path is blank.")
    end

    mkpath(dirname(file))
    ensure_plots_loaded()

    # Evaluate plotting code in Main and require it to return a plot object.
    plt = Core.eval(Main, Meta.parse(code))

    plt === nothing && error("Plot code did not return a plot object.")

    # Apply requested size in pixels using Plots.jl size=(w,h).
    try
        Core.eval(Main, :(Plots.plot!($plt, size=($(Int(width)), $(Int(height))))))
    catch
        # If plot! sizing fails for some plot object, continue and save anyway.
    end

    Core.eval(Main, :(Plots.savefig($plt, $file)))
    return normpath(file)
end


# --- Fast transfer helpers ---
using Dates
using UUIDs

const LAST_TRANSFER = Dict{String, Any}(
    "method" => "none",
    "source" => "startup",
    "name" => "",
    "class" => "",
    "rows" => 0,
    "cols" => 0,
    "elapsed_ms" => 0.0
)

function _set_last_transfer(method::AbstractString, source::AbstractString, name::AbstractString, obj, rows::Integer, cols::Integer, elapsed_ms)
    LAST_TRANSFER["method"] = String(method)
    LAST_TRANSFER["source"] = String(source)
    LAST_TRANSFER["name"] = String(name)
    LAST_TRANSFER["class"] = string(typeof(obj))
    LAST_TRANSFER["rows"] = Int(rows)
    LAST_TRANSFER["cols"] = Int(cols)
    LAST_TRANSFER["elapsed_ms"] = round(Float64(elapsed_ms); digits=3)
end

function last_transfer_table()
    rows = Any[["Field", "Value"]]
    for k in ["method", "source", "name", "class", "rows", "cols", "elapsed_ms"]
        push!(rows, [k, LAST_TRANSFER[k]])
    end
    return rows
end

function _transfer_dir()
    d = joinpath(tempdir(), "JuliaExcelBridgeTransfer")
    mkpath(d)
    return d
end

function _resolve_named_object(name::AbstractString)
    nm = strip(String(name))

    if haskey(OBJECT_STORE, nm)
        return OBJECT_STORE[nm]
    end

    sym = Symbol(nm)

    if isdefined(Main, sym)
        return getfield(Main, sym)
    end

    # fallback: allow names like Main.df or df
    try
        return Core.eval(Main, Meta.parse(nm))
    catch
        error("Object not found: " * nm)
    end
end

function _array_shape_for_excel(x)
    if x isa AbstractMatrix
        return size(x, 1), size(x, 2)
    elseif x isa AbstractVector
        return length(x), 1
    elseif x isa Number
        return 1, 1
    else
        return 0, 0
    end
end

function is_numeric_exportable(x)
    x isa Number || x isa AbstractArray{<:Number}
end

function export_numeric_object(x; source="JGetNumeric", name="")
    t0 = time()
    rows, cols = _array_shape_for_excel(x)
    rows > 0 && cols > 0 || error("Object is not a numeric scalar/vector/matrix.")

    file = joinpath(_transfer_dir(), "jexcel_get_numeric_" * string(uuid4()) * ".bin")
    open(file, "w") do io
        if x isa Number
            write(io, Float64(x))
        elseif x isa AbstractVector
            for r in eachindex(x)
                v = x[r]
                write(io, ismissing(v) ? NaN : Float64(v))
            end
        else
            for r in 1:rows, c in 1:cols
                v = x[r, c]
                write(io, ismissing(v) ? NaN : Float64(v))
            end
        end
    end
    elapsed = (time() - t0) * 1000.0
    _set_last_transfer("numeric", source, String(name), x, rows, cols, elapsed)
    return Dict("__jexcel_transfer_type" => "numeric", "file" => file, "rows" => rows, "cols" => cols, "class" => string(typeof(x)))
end

function set_numeric_from_file(name::AbstractString, file::AbstractString, rows::Integer, cols::Integer)
    t0 = time()
    values = Vector{Float64}(undef, Int(rows) * Int(cols))
    open(file, "r") do io
        read!(io, values)
    end
    arr = Matrix{Float64}(undef, Int(rows), Int(cols))
    k = 1
    for r in 1:Int(rows), c in 1:Int(cols)
        arr[r, c] = values[k]
        k += 1
    end
    OBJECT_STORE[String(name)] = arr
    Core.eval(Main, :($(Symbol(name)) = OBJECT_STORE[$(String(name))]))
    elapsed = (time() - t0) * 1000.0
    _set_last_transfer("numeric to Julia", "JSet", String(name), arr, rows, cols, elapsed)
    return true
end

function _try_import_dataframes()
    try
        Base.require(Main, :DataFrames)
        return getfield(Main, :DataFrames)
    catch
        try
            Core.eval(Main, :(using DataFrames))
            return getfield(Main, :DataFrames)
        catch e
            error("DataFrames.jl is required for table transfer. In Julia run: using Pkg; Pkg.add(\"DataFrames\")")
        end
    end
end

function is_dataframe_like(x)
    occursin("DataFrame", string(typeof(x)))
end

function nrow_df(x)
    try
        return Int(Base.invokelatest(getfield(_try_import_dataframes(), :nrow), x))
    catch
        return Int(Base.invokelatest(size, x, 1))
    end
end

function ncol_df(x)
    try
        return Int(Base.invokelatest(getfield(_try_import_dataframes(), :ncol), x))
    catch
        return Int(Base.invokelatest(size, x, 2))
    end
end

function table_column_names(x)
    return [String(nm) for nm in Base.invokelatest(propertynames, x)]
end



function export_table_object(x; source="JGetTable", name="")
    t0 = time()
    is_dataframe_like(x) || error("Object is not a DataFrame.")

    rows = nrow_df(x)
    colnames = table_column_names(x)
    cols = length(colnames)
    columns = Any[]

    for cname in colnames
        col = Base.invokelatest(getindex, x, !, Symbol(cname))
        nonmissing = [v for v in col if !ismissing(v)]

        # Fast numeric path only when there are no missing values.
        # If numeric column has missing values, send through mixed/character path
        # so missing can return to Excel as a blank instead of 0.
        if all(v -> v isa Number, nonmissing) && !any(ismissing, col)
            vals = Vector{Float64}(undef, rows)

            for i in 1:rows
                vals[i] = Float64(col[i])
            end

            file = joinpath(_transfer_dir(), "jexcel_get_table_" * string(uuid4()) * ".bin")

            open(file, "w") do io
                for v in vals
                    write(io, v)
                end
            end

            push!(columns, Dict(
                "name" => cname,
                "type" => "numeric",
                "file" => file,
                "na" => "NaN"
            ))
        else
            vals = Any[]

            for i in 1:rows
                push!(vals, ismissing(col[i]) ? "" : coerce_for_json(col[i]))
            end

            push!(columns, Dict(
                "name" => cname,
                "type" => "character",
                "values" => vals
            ))
        end
    end

    elapsed = (time() - t0) * 1000.0

    _set_last_transfer("typed table", source, String(name), x, rows, cols, elapsed)

    return Dict(
        "__jexcel_transfer_type" => "table",
        "rows" => rows,
        "cols" => cols,
        "include_headers" => true,
        "columns" => columns,
        "class" => string(typeof(x))
    )
end


function set_table_from_payload(name::AbstractString, rows::Integer, cols::Integer, columns)
    t0 = time()
    DF = _try_import_dataframes()
    pairs = Pair{Symbol, Any}[]
    for col in columns
        cname = String(get(col, "name", "V" * string(length(pairs) + 1)))
        typ = String(get(col, "type", "character"))
        if typ == "numeric"
            file = String(get(col, "file", ""))
            vals = Vector{Float64}(undef, Int(rows))
            open(file, "r") do io
                read!(io, vals)
            end
            push!(pairs, Symbol(cname) => vals)
        else
            vals = get(col, "values", Any[])
            push!(pairs, Symbol(cname) => Any[v === nothing ? missing : v for v in vals])
        end
    end
    df = DF.DataFrame(pairs...)
    OBJECT_STORE[String(name)] = df
    Core.eval(Main, :($(Symbol(name)) = OBJECT_STORE[$(String(name))]))
    elapsed = (time() - t0) * 1000.0
    _set_last_transfer("typed table to Julia", "JSetTable", String(name), df, rows, cols, elapsed)
    return true
end

function coerce_for_transport(x; source="general", name="")
    if is_numeric_exportable(x)
        return export_numeric_object(x; source=source, name=name)
    elseif is_dataframe_like(x)
        return export_table_object(x; source=source, name=name)
    else
        return coerce_for_json(x)
    end
end

function _contains_assignment_expr(ex)
    if ex isa Expr
        if ex.head == :(=) || ex.head == :(:=) || ex.head == :(.=)
            return true
        end
        return any(_contains_assignment_expr, ex.args)
    end
    return false
end

function _assigned_symbols!(out::Vector{Symbol}, ex)
    if ex isa Expr
        if ex.head == :(=) && length(ex.args) >= 2
            lhs = ex.args[1]
            if lhs isa Symbol
                push!(out, lhs)
            end
        end
        for a in ex.args
            _assigned_symbols!(out, a)
        end
    end
    return out
end

function _simple_global_assignment_name(code::AbstractString)
    # Common Excel workflow: JEval("x = randn(10000,20)").
    # Force this case to top-level global assignment so the object is visible
    # to later JGetNumeric/JCall calls in the persistent worker.
    m = match(r"^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=(?!=)", code)
    return m === nothing ? nothing : String(m.captures[1])
end

function eval_code_for_excel(code)
    expr = Meta.parse("begin\n" * code * "\nend")

    has_assignment = _contains_assignment_expr(expr)
    assigned = Symbol[]
    _assigned_symbols!(assigned, expr)

    value = Core.eval(Main, expr)

    if has_assignment
        for sym in assigned
            if isdefined(Main, sym)
                OBJECT_STORE[String(sym)] = getfield(Main, sym)
            end
        end
        return "OK"
    end

    return coerce_for_transport(value; source="JEval", name="")
end

    
function handle_request(req::AbstractDict)
    id = get(req, "id", nothing)
    cmd = get(req, "cmd", "")

    try
        if cmd == "ping"
            return make_ok(id, "OK | JuliaExcelBridge")

        elseif cmd == "source"
            file = get(req, "file", nothing)
            file isa AbstractString || error("file is missing.")
            include(file)
            return make_ok(id, true)

        elseif cmd == "eval"
            code = get(req, "code", nothing)
            code isa AbstractString || error("code is missing.")
            value = eval_code_for_excel(code)
            return make_ok(id, value)

        elseif cmd == "plot"
            code = get(req, "code", nothing)
            file = get(req, "file", nothing)
            width = get(req, "width", 800)
            height = get(req, "height", 600)

            code isa AbstractString || error("code is missing.")
            file isa AbstractString || error("file is missing.")

            return make_ok(id, render_plot_to_file(code, file, Int(width), Int(height)))

        elseif cmd == "call"
            fun_name = get(req, "fun", nothing)
            fun_name isa AbstractString || error("fun is missing.")
            f = resolve_function(fun_name)
            args = normalize_call_args(get(req, "args", Any[]))
            value = f(args...)
            return make_ok(id, coerce_for_transport(value; source="JCall", name=fun_name))

        elseif cmd == "set"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")
            value = convert_json_value(get(req, "value", nothing))
            OBJECT_STORE[name] = value
            Core.eval(Main, :($(Symbol(name)) = OBJECT_STORE[$name]))
            return make_ok(id, true)

        elseif cmd == "get"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")

            if haskey(OBJECT_STORE, name)
                return make_ok(id, coerce_for_transport(OBJECT_STORE[name]; source="JGet", name=name))
            elseif isdefined(Main, Symbol(name))
                value = getfield(Main, Symbol(name))
                return make_ok(id, coerce_for_transport(value; source="JGet", name=name))
            else
                error("Object '" * name * "' was not found.")
            end

        elseif cmd == "set_numeric"
            name = get(req, "name", nothing)
            file = get(req, "file", nothing)
            rows = Int(get(req, "rows", 0))
            cols = Int(get(req, "cols", 0))
            name isa AbstractString || error("name is missing.")
            file isa AbstractString || error("file is missing.")
            return make_ok(id, set_numeric_from_file(name, file, rows, cols))

        elseif cmd == "get_numeric"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")
            return make_ok(id, export_numeric_object(_resolve_named_object(name); source="JGetNumeric", name=name))

        elseif cmd == "set_table"
            name = get(req, "name", nothing)
            rows = Int(get(req, "rows", 0))
            cols = Int(get(req, "cols", 0))
            columns = get(req, "columns", Any[])
            name isa AbstractString || error("name is missing.")
            return make_ok(id, set_table_from_payload(name, rows, cols, columns))

        elseif cmd == "get_table"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")
            return make_ok(id, export_table_object(_resolve_named_object(name); source="JGetTable", name=name))

        elseif cmd == "last_transfer"
            return make_ok(id, last_transfer_table())

        elseif cmd == "exists"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")
            return make_ok(id, haskey(OBJECT_STORE, name) || isdefined(Main, Symbol(name)))

        elseif cmd == "remove"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")
            if haskey(OBJECT_STORE, name)
                delete!(OBJECT_STORE, name)
            end
            if isdefined(Main, Symbol(name))
                Core.eval(Main, :($(Symbol(name)) = nothing))
            end
            return make_ok(id, true)

        elseif cmd == "objects"
            rows = Any[["Name", "Class", "Type", "Length", "Dimensions"]]

            for nm in sort(collect(keys(OBJECT_STORE)))
                push!(rows, object_summary_row(nm, OBJECT_STORE[nm]))
            end

            return make_ok(id, rows)

        elseif cmd == "describe"
            name = get(req, "name", nothing)
            name isa AbstractString || error("name is missing.")

            if haskey(OBJECT_STORE, name)
                value = OBJECT_STORE[name]
                return make_ok(id, object_describe_table(name, value))
            elseif isdefined(Main, Symbol(name))
                value = getfield(Main, Symbol(name))
                return make_ok(id, object_describe_table(name, value))
            else
                error("Object '" * name * "' was not found.")
            end

        else
            return make_err(id, "Unknown command: " * string(cmd))
        end

    catch e
        return make_err(id, sprint(showerror, e))
    end
end

function main()
    startup_file = length(ARGS) >= 1 ? ARGS[1] : nothing

    try
        if startup_file !== nothing && isfile(startup_file)
            println(stderr, "Including startup file: " * startup_file)
            flush(stderr)
            include(startup_file)
            println(stderr, "Startup file loaded successfully.")
            flush(stderr)
        else
            println(stderr, "No startup file provided.")
            flush(stderr)
        end
    catch e
        println(stderr, "FATAL startup error: " * sprint(showerror, e, catch_backtrace()))
        flush(stderr)
        println(JSON.json(make_err(nothing, "Startup error: " * sprint(showerror, e))))
        flush(stdout)
        return
    end

    println(stderr, "Worker entering request loop.")
    flush(stderr)

    while !eof(stdin)
        line = try
            readline(stdin)
        catch e
            println(stderr, "Readline failed: " * sprint(showerror, e, catch_backtrace()))
            flush(stderr)
            break
        end

        line = strip(line)
        isempty(line) && continue

        req = try
            JSON.parse(line)
        catch e
            println(JSON.json(make_err(nothing, "Invalid JSON: " * sprint(showerror, e))))
            flush(stdout)
            continue
        end

        resp = try
            handle_request(req)
        catch e
            make_err(nothing, sprint(showerror, e, catch_backtrace()))
        end

        println(JSON.json(resp))
        flush(stdout)
    end
end

main()