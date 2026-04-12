# JuliaExcelBridge worker
# JSON protocol over stdin/stdout, one request per line, one response per line.

import Pkg

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
    else
        return x
    end
end

normalize_call_args(args) = args isa Vector ? map(convert_json_value, args) : [convert_json_value(args)]

function resolve_function(fun_name::AbstractString)
    if occursin(".", fun_name)
        parts = split(fun_name, ".")
        obj = getfield(Main, Symbol(parts[1]))
        for p in parts[2:end]
            obj = getfield(obj, Symbol(p))
        end
        return obj
    end

    if isdefined(Main, Symbol(fun_name))
        return getfield(Main, Symbol(fun_name))
    elseif isdefined(Base, Symbol(fun_name))
        return getfield(Base, Symbol(fun_name))
    else
        error("Function not found: " * fun_name)
    end
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
    elseif x isa Dict
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

function eval_code_for_excel(code)
    expr = Meta.parse("begin\n" * code * "\nend")
    value = Core.eval(Main, expr)
    return coerce_for_json(value)
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
            return make_ok(id, coerce_for_json(value))

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
                return make_ok(id, coerce_for_json(OBJECT_STORE[name]))
            elseif isdefined(Main, Symbol(name))
                value = getfield(Main, Symbol(name))
                return make_ok(id, coerce_for_json(value))
            else
                error("Object '" * name * "' was not found.")
            end

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