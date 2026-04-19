import math


def hello_world():
    return "Hello from Python"


def add_numbers(x, y):
    return x + y


def make_seq_table(start, stop, step=1):
    xs = list(range(int(start), int(stop) + 1, int(step)))
    return [[x, x * x] for x in xs]


try:
    import pandas as pd  # noqa: F401
except Exception:
    pass

from PythonFunctions import *
