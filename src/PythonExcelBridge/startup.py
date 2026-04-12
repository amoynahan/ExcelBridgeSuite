import math


def hello_world():
    return "Hello from Python"


def add_numbers(x, y):
    return x + y


def make_seq_table(start, stop, step=1):
    xs = list(range(int(start), int(stop) + 1, int(step)))
    return [[x, x * x] for x in xs]


from PythonFunctions import *
