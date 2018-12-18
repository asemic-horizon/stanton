from greenbox import *
import xlwings as xw
import sys

n_samples = 10000
print('{} samples'.format(n_samples))
greenbox = Greenbox(xw)
bluebox = Bluebox(greenbox)

bluebox.sample(n_samples)

from sklearn.tree import DecisionTreeRegressor, export_graphviz
tree = DecisionTreeRegressor()
tree.fit(X=bluebox.input_samples, y = bluebox.outcomes['_y'])

export_graphviz(decision_tree = tree, out_file="test.dot", feature_names = bluebox.input_names,
                label = None, impurity = False, node_ids = False, rounded = True, filled = True)
