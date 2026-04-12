using LinearAlgebra

hello_world() = "Hello from Julia"

add_numbers(x, y) = x + y

make_seq_table(from, to; step = 1) = [collect(from:step:to) (collect(from:step:to) .^ 2)]

include("JuliaFunctions.jl")
