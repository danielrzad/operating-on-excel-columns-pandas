from datetime import date
import numpy as np

today = date.today()
print(today)
print(type(today))
dt64 = np.datetime64(today)
print(dt64)
print(type(dt64))