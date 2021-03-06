# utf-8

import os
import traceback
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import pandas as pd

def get_data(filepath, sheet_idx):
  print ("Opening file {0}\n".format(filepath))
  xl = xlrd.open_workbook(filepath).sheet_by_index(sheet_idx)

  cell_0_0 = xl.cell(0, 0).value
  cell_0_1 = xl.cell(0, 1).value
  cell_0_2 = xl.cell(0, 2).value
  
  print ("Title of graph : {0}, x-label : {1}, y-label : {2}".format(cell_0_0, cell_0_1, cell_0_2))
  
  df = pd.read_excel(filepath, sheet_name=sheet_idx, skiprows = [0], index_col=0)
  return df, str(cell_0_0), str(cell_0_1), str(cell_0_2)
  
  
# get_data  
  
def plot_bar(filepath, worksheet_index, error_bar=False, verbosity=False):
  df, title, xlabel, ylabel = get_data (os.path.normpath(filepath), worksheet_index)

  df = df.T
  print ("\nData:\n {0}".format(df))
  
  ncol = len(df.columns)
  print ("\nNumber of columns = {0}".format(ncol))
  
  if error_bar:
    nrow = int(len(df.index)/2)
    print("\nNumber of rows (excluding error values) = {0}".format(nrow))
  else:
    nrow = len(df.index)
    print("\nNumber of rows = {0}".format(nrow))
  
  ind = np.arange(nrow)
  margin = 0.1
  width = (1.-2.*margin)/ncol

  hatch_patterns = ['/','+', '.', 'x', '\\', '-', '*', 'o', 'O']
    
  for i in range(ncol):
    xdata = ind + margin + (i*width)
    
    if verbosity:
      print("x = {0}\nheight = \n{1}".format(xdata, df[df.columns[i]]))
    
    if error_bar:
      plt.bar(x=xdata, height=df[df.columns[i]][:nrow], 
              width=width-0.08, 
              yerr=df[df.columns[i]][nrow:], capsize=6,
              facecolor='w', edgecolor='k',
              hatch=3*hatch_patterns[i%len(hatch_patterns)]
              )
    else:
      plt.bar(x=xdata, height=df[df.columns[i]], 
              width=width-0.08,
              facecolor='w', edgecolor='k',
              hatch=3*hatch_patterns[i%len(hatch_patterns)]
              )
    
  plt.title(title)
  plt.xticks(ind+0.38, list(df.index))
  plt.xlabel(xlabel)
  plt.ylabel(ylabel)

  plt.legend(df.columns)

  plt.style.use('seaborn-white')
  
  printAndSave (plt, title, filepath, worksheet_index)
# plot_bar  

def plot_line(filepath, worksheet_index, error_bar=False, verbosity=False):
  df, title, xlabel, ylabel = get_data (os.path.normpath(filepath), worksheet_index)

  print ("\nData:\n {0}".format(df))
  
  ncol = len(df.columns)
  print ("\nNumber of columns = {0}".format(ncol))
  
  if error_bar:
    nrow = int(len(df.index)/2)
    print("\nNumber of rows (excluding error values) = {0}".format(nrow))
  else:
    nrow = len(df.index)
    print("\nNumber of rows = {0}".format(nrow))
  
  ax = df.plot(style=['o-','*-','^-','s-'])
  
  ax.set_ylabel(ylabel)
  ax.set_title(title)
  ax.set_xticks(df.index)
  
  plt.style.use('seaborn-white')
  
  printAndSave (plt, title, filepath, worksheet_index)


def printAndSave (plt, title, filepath, worksheet_index):
  title_ = title.replace(" ", "_")  
  filename = os.path.splitext(os.path.basename(filepath))[0]
  dirname = os.path.dirname(filepath)
  fullname = dirname+'\\'+\
              "_"+title_+\
              "_"+str(filename)+\
              "_"+str(worksheet_index)+".png"

  print("Saving as {0}".format(fullname))

  plt.savefig(fullname, bbox='tight')

  plt.show()
# printAndSave
  
# Call function
if __name__ == "__main__":
  import sys
  import argparse
  
  parser = argparse.ArgumentParser()
  
  parser.add_argument("filepath", help="Enter the data file path")

  parser.add_argument("-e", "--errorbar", dest="error_bar", action="store_true",\
                      default=False, help="Show error bars")
                      
  parser.add_argument("-v", "--verbose", dest="verbosity", action="store_true",\
                      default=False, help="Prints more information")           
                      
  parser.add_argument("-s", "--sheet", dest="sheet_num", default=1, \
                      type=int, help="Worksheet number to use (defaults to 1)")

  parser.add_argument("-t", "--type", dest="chart_type", default='bar', \
                      type=str, help="Chart type : bar, line, pie (default is bar)")
                      
  args = parser.parse_args()

  if not os.path.isfile(args.filepath):
    print("Input file not found !!!")
  else:
    try:
      if args.chart_type == 'bar':
        print ("\nPlotting bar graph ...\n")
        plot_bar(args.filepath, args.sheet_num-1, args.error_bar, args.verbosity)
      elif args.chart_type == 'line':
        print ("\nPlotting line graph ...\n")
        plot_line(args.filepath, args.sheet_num-1, args.error_bar, args.verbosity)
      else:
        print ("Chart type {0} not supported\n".format(args.chart_type))
    except:
        print("EXCEPTION")
        traceback.print_exc()
# Done
