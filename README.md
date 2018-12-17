# Greenbox

Greenbox is a Python module for Monte Carlo three-point-estimate sensitivity analysis. 

As a demonstration, we'll study the range of a simple neural network:

![](https://github.com/asemic-horizon/stanton/blob/master/net1.png)

Here, `_x1` and `_x2` are fixed inputs; `_a1,_a2` are weights leading into the intermediate unit `_a`, `_b1,_b2` are weights leading into the intermediate unit `_b` and finally `_y` is linked to `_a` and `_b` by the weights `_wa, _wb`. A more realistic use case, of course, are those financial spreadsheets that grow by accretion of consensus and that can't really be developed by alternate methodologies (such as Jupyter notebooks). 

We will get to a result such as this:

![](https://github.com/asemic-horizon/stanton/blob/master/output%20_y.png)

as well as a spreadsheet containing histogram data for plotting directly in Excel, the full set of sample inputs and resulting outputs and some summary statistics.

To use Greenbox, first give names to the cells you want to change (inputs) or watch (outputs)

![](https://github.com/asemic-horizon/stanton/blob/master/cell_names.png)

Then define two named ranges named `greenbox` and `bluebox`. The first will have five columns as follows:

![](https://github.com/asemic-horizon/stanton/blob/master/greenbox.png)

The first column is the name of the cell that will be stressed; the next three columns are minimum, central and maximum estimates and the last is a [concentration parameter](https://en.wikipedia.org/wiki/Beta_distribution#Mode_and_concentration). The larger this parameter, the more the distribution is concentrated around its central value. $\kappa = 2$ gives the uniform distribution; $\kappa=1$ gives the Jeffrey's distribution where the extremes are more likely than the center.

The `bluebox` is simply a list of outputs.

![](https://github.com/asemic-horizon/stanton/blob/master/greenbox.png)

Python will actually only read the first column; the second says `=INDIRECT()` to the cell name so we can watch the numbers shuffle in one place, but that's optional.

To run simulations, we use a snippet of code such as the following:

    from greenbox import *
    import xlwings as xw
    import sys
    
    n_samples = 500
    print('{} samples'.format(n_samples))
    greenbox = Greenbox(xw)
    bluebox = Bluebox(greenbox)
    
    bluebox.sample(n_samples)
    bluebox.plot()
    bluebox.to_excel()

 
(*Note that realistically you should be sampling ~1K for this problem size and  ~10K for any kind of complex sheet with >5 random inputs*).
