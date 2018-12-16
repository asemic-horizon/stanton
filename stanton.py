import xlwings as xw
import numpy as np
from numpy.random import normal, gamma, beta
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from time import time
import sys

class Sampler():
    """Three-point estimate sampler using the reparameterized beta function.

    Initialization arguments and object properties:
    left -- the lowermost bound of the three-point estimate
    mode -- the most likely value of the three-point estimate
    right -- the uppermost bound of the three-point estimate
    kappa -- concentration parameter. As kappa goes to infinity the distribution becomes concentrated arounds the mode.

    Returns an object with the method "self.sample that takes the keyword size
    as the size (usually a number) of the output with i.i.d samples."""

    def __init__(self, left, mode, right, kappa):
        self.left = left
        self.mode = mode
        self.right = right
        self.kappa = kappa
        self.sample = self.make_sampler(left, mode, right, kappa)

    def make_sampler(self, left,mode,right,kappa):
        loc = (mode -left)/(right - left)
        al = loc * (kappa - 2) + 1
        be = (1-loc)*(kappa -2) + 1
        std_sampler = lambda size: beta(a = al, b = be, size = size)
        def sample(size):
            """Three-point estimate sampler: takes keyword "size" """
            return left + (right - left) * std_sampler(size)
        return sample

class Greenbox():
    def __init__(self, excel_ref, range_name='greenbox'):
        """
        The greenbox is where inputs or variables to be stressed go.

        In your Excel spreadsheet, the greenbox will be a named range with the following
        five columns in exact order and without column titles:

        Column 1 - Variable name: this must be the name of a cell in your spreadsheet;
        Column 2 - Left or lowermost bound of a three-point estimate;
        Column 3 - Mode or most likely value of a three-point estimate;
        Column 4 - Right or uppermost bound of a three-point estimate;
        Column 5 - Kappa or concentration parameter. As kappa goes to infinity
                   the distribution becomes concentrated arounds the mode.

        The greenbox can have as many rows as needed, although succintness is advised
        because getting a decent number (~10K) of samples becomes very slow.

        In Python the greenbox will be represented by an object with a "sample" method.
        It is initialized with the following keywords:

        excel_ref - an xlwings objects that accesses range names
        range_name (optional, default "greenbox") a range name."""
        self.excel_ref = excel_ref
        self.range_name = range_name
        self.greenbox = self.greenbox()
        self.input_names = list(self.greenbox.keys())
        self.original_state = self.save_state()

    def greenbox(self):
        excel_range = self.excel_ref.Range(self.range_name).value
        samplers = dict()
        for spec in excel_range:
            varname, left, mode, right, kappa = tuple(spec)
            if not varname is None:
                print(spec)
                samplers[varname] = Sampler(left, mode, right, kappa)
        return samplers
    def save_state(self):
        """Saves the original state of the greenbox variables so changes can be
        undone after sensitivity analysis"""
        original = dict()
        for input in self.input_names:
            original[input] = self.excel_ref.Range(input).value
        return original
    def sample(self,size):
        """Greenbox sampler. Takes the keyword "size" """
        raw_samples = pd.DataFrame(columns = self.input_names)
        for var in self.input_names:
            raw_samples[var] = self.greenbox[var].sample(size)
        return raw_samples

class Bluebox():
        def __init__(self, greenbox, excel_ref = None, range_name='bluebox'):
            """
            The bluebox is where variables whose response to stress is measured.

            In your Excel spreadsheet, the bluebox will be a named range with at least
            one column containing the name of a cell to be watched. The bluebox can have
            as many rows as needed, with a lower performance penalty as compared to
            the greenbox.

            In Python, the Bluebox is initialized by the following keywords:

            greenbox - a previously initialized greenbox
            excel_ref (optional) - an xlwings object that accesses range names
            range_name (optional, default "bluebox") - a range name.

            The Bluebox has a method "sample" that also samples the greenbox.
            """
            self.greenbox = greenbox
            self.input_names = self.greenbox.input_names
            if excel_ref is None:
                self.excel_ref = self.greenbox.excel_ref
            else:
                self.excel_ref = excel_ref
            self.range_name = range_name
            self.bluebox = self.bluebox()
        def bluebox(self):
            excel_range = self.excel_ref.Range(self.range_name).value
            names = [s[0] for s in excel_range if s[0] is not None]
            return names
        def sample(self, size, verbose = True):
            t0 = time(); t1 = time()
            self.input_samples = self.greenbox.sample(size)
            self.outcomes = pd.DataFrame(columns = self.bluebox)
            for i in self.input_samples.index:
                if verbose and i % 50 == 0:
                    tt0 = time()-t0; tt1 = time()-t1
                    est = ((size-i)/i)*tt0 if i>0 else np.inf
                    print("{} samples in {:3.2f}mins, last batch in {:3.2f}s, est. {:3.2f} mins remaining".format(i,tt0/60,tt1,est/60))
                    t1 = time()
                row = self.input_samples.iloc[i]
                # sometimes trying to set an excel value fails for mysterious reasons
                # but losing a sample isn't a death blow
                try:
                    for var in self.input_names:
                        self.excel_ref.Range(var).value = row[var]
                    res = dict()
                    for outcome in self.bluebox:
                        res[outcome] = self.excel_ref.Range(outcome).value

                    self.outcomes = self.outcomes.append(res, ignore_index = True)
                except:
                    print('Malfunctioning inputs:\n')
                    print(row)


        def pd_hist(self,v):
            values, _from = np.histogram(v, bins='sqrt')
            _from = _from.reshape(-1,1)
            values = np.insert(values,0,0).reshape(-1,1)
            _to = np.insert(_from[1:],len(_from)-1,np.inf).reshape(-1,1)
            table = np.concatenate([_from, _to, values],axis=1)
            return pd.DataFrame(table,columns=['from','to','freq'])
        def to_excel(self, filename = 'sensitivity.xlsx'):
            writer = pd.ExcelWriter(filename)
            self.input_samples.to_excel(writer,'Inputs')
            self.outcomes.to_excel(writer,'Outputs')
            for var in vars: pd_hist(u[var]).to_excel(writer,'in '+var)
            for outcome in outcomes: pd_hist(v[outcome]).to_excel(writer,'out '+outcome)
        def plot(self):
                for var in self.input_names:
                    fig, ax = plt.subplots()

                    sns.distplot(self.input_samples[var], kde = True)
                    ax.set_title('Input: '+var)
                    fig.savefig('input '+var+'.png')
                for outcome in self.outcomes:
                    fig, ax = plt.subplots()
                    sns.distplot(self.outcomes[outcome], kde = True)
                    ax.set_title('Output: '+outcome)
                    fig.savefig('output '+outcome+'.png')
