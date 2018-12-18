# Greenbox: a module for sensitivity analysis and model emulation

Greenbox is a Python module for Monte Carlo three-point-estimate sensitivity analysis and model emulation in Excel. It relies on [xlwings](https://www.xlwings.org/) to operate Excel from within Python. To install, clone this repository (or just download `greenbox.py` and `requirements.txt` and do

    pip install -r requirements.txt

## Sensitivity analysis

As a demonstration, we'll study the "range of expression" of a simple neural network:

![](https://github.com/asemic-horizon/stanton/blob/master/net1.png)

Here, `_x1` and `_x2` are fixed inputs; `_a1,_a2` are weights leading into the intermediate unit `_a`, `_b1,_b2` are weights leading into the intermediate unit `_b` and finally `_y` is linked to `_a` and `_b` by the weights `_wa, _wb`. A more realistic use case, of course, are those financial spreadsheets that grow by accretion of consensus and that can't really be developed by alternate methodologies (such as Jupyter notebooks). Our goal is to specify probability distributions for the inputs such as

![](https://github.com/asemic-horizon/stanton/blob/master/input%20_a1.png)


and arrive at a result such as this:

![](https://github.com/asemic-horizon/stanton/blob/master/output%20_y.png)

as well as a spreadsheet containing histogram data for plotting directly in Excel, the full set of sample inputs and resulting outputs and some summary statistics.

To use Greenbox, first give names to the cells you want to change (inputs) or watch (outputs)

![](https://github.com/asemic-horizon/stanton/blob/master/cell_names.png)

Then define two named ranges named `greenbox` and `bluebox`. The first will have five columns as follows:

![](https://github.com/asemic-horizon/stanton/blob/master/greenbox.png)

The first column is the name of the cell that will be stressed; the next three columns are minimum, central and maximum estimates and the last is a [concentration parameter](https://en.wikipedia.org/wiki/Beta_distribution#Mode_and_concentration). The larger this parameter, the more the distribution is concentrated around its central value. $\kappa = 2$ gives the uniform distribution; $\kappa=1$ gives the Jeffrey's distribution where the extremes are more likely than the center.

The `bluebox` is simply a list of outputs.

![](https://github.com/asemic-horizon/stanton/blob/master/bluebox.png)

Python will actually only read the first column; the second says `=INDIRECT()` to the cell name so we can watch the numbers shuffle in one place, but that's optional.

To run simulations, we import the `Greenbox` and `Bluebox` classes from the `greenbox` module and use them in script mode as follows:

    from greenbox import *
    import xlwings as xw

    n_samples = 500
    print('{} samples'.format(n_samples))

    # we pass an object that supports the .Range() method. The simples is `xw` itself,
    # but this requires the target spreadsheet to remain in Excel focus all the time
    greenbox = Greenbox(xw)

    # notionally a greenbox could correspond to a number of blueboxes
    # we can even pass a separate .Range-supporting object using the  `excel_ref` parameter at init.
    bluebox = Bluebox(greenbox)

    # this does the heavy lifting.
    bluebox.sample(n_samples)

    # this saves histogram plots in png format for the greenbox (input) and bluebox (output) variables.
    bluebox.plot()
    # on an environment such as Jupyter you can pass save=False.

    # this makes an excel spreadsheet
    bluebox.to_excel(filename = 'sensitivity.xlsx)

(*Note that realistically you should be sampling ~1K for this problem size and  ~10K for any kind of complex model with >5 random inputs*).

    from greenbox import *
    import xlwings as xw

    n_samples = 500

    greenbox = Greenbox(xw)
    bluebox = Bluebox(greenbox)

    bluebox.sample(n_samples)

    bluebox.plot(inputs = False, outputs = True, save = False)

## Model emulation

The idea of model emulation is to approximately replicate a spreadsheet that's too complex to understand with a statistical or machine learning model that's too opaque to understand. (*Sad trombone*) Here we will approximate a spreadsheet by a decision tree that is sort of visually inspectable.


    from greenbox import *
    import xlwings as xw
    n_samples = 10000
    greenbox = Greenbox(xw)
    bluebox = Bluebox(greenbox)
    bluebox.sample(n_samples)

    from sklearn.tree import DecisionTreeRegressor
    tree = DecisionTreeRegressor()
    tree.fit(X=bluebox.input_samples, y = bluebox.outcomes['_y'])
    export_graphviz(decision_tree = tree, out_file="test.dot", feature_names = bluebox.input_names,
                label = None, impurity = False, node_ids = False, rounded = True, filled = True)

The main idea in this code snippet is that input samples and outcomes are saved in the `bluebox` object as the `input_samples` and `outcomes` attributes, each of which is a pandas DataFrame. Here are the results.

![](https://github.com/asemic-horizon/stanton/blob/master/tree.png)
