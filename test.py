from greenbox import *
import xlwings as xw
import sys

n_samples = 500
print('{} samples'.format(n_samples))
greenbox = Greenbox(xw)
bluebox = Bluebox(greenbox)

bluebox.sample(n_samples)
bluebox.to_excel()
bluebox.plot()
