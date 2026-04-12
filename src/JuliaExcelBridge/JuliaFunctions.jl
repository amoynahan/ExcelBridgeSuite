using LinearAlgebra

function CholDecomp(mat)
    m = Matrix{Float64}(mat)
    Matrix(cholesky(Symmetric(m)).U)
end

function CholDecompLower(mat)
    m = Matrix{Float64}(mat)
    Matrix(cholesky(Symmetric(m)).L)
end

function ShowObjectInfo(x)
    dims = isa(x, AbstractArray) ? join(size(x), "x") : "NULL"
    string("type = ", typeof(x), "; dim = ", dims, "; length = ", length(x))
end

function MatrixMultiply(a, b)
    Matrix{Float64}(a) * Matrix{Float64}(b)
end

function IdentityMatrix(n)
    Matrix{Float64}(I, Int(n), Int(n))
end

function RowSums(mat)
    vec(sum(Matrix{Float64}(mat), dims = 2))
end

function ReloadFunctionsJulia()
    include("JuliaFunctions.jl")
    "JuliaFunctions.jl reloaded"
end


hello_julia_bridge() = "Hello from JuliaExcelBridge"
