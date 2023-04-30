# -*- coding: utf-8 -*-
"""
Created on Thu Apr  6 23:22:54 2023

@author: dkane
"""
import pandas as pd
import yaml
import os
import py7zr
import glob
import warnings

fp = r"C:\Users\dkane\OneDrive - Presto Engineering\Documents\AMF\56G PD Quals\Lot #2\test data\HTOL\56GPDL2_TO39_HTOL_T0_Test_112122_154047.csv"

# from yaml import load
try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader

if __name__ == '__main__':
    
    # with open(r"C:/Users/dkane/OneDrive - Presto Engineering/Documents/python_scripts/AMF/config/56GPDL3_config.yml", 'r') as stream:
    #     config = yaml.load(stream, Loader)
    #     # print(config)
    #     for key, value in config.items():
    #         print (key + " : " + str(value))
    # zip_fp = r"C:/Users/dkane/OneDrive - Presto Engineering/Documents/AMF/56G PD Quals/Lot #2/test data/56GPDL2_plots.7z"
    # archive = py7zr.SevenZipFile(zip_fp, mode='w')
    # fp = r"C:/Users/dkane/OneDrive - Presto Engineering/Documents/AMF/56G PD Quals/Lot #2/test data/56GPDL2_SM_DH_1000HR_screened_and_plotted.xlsx"
    # archive.write(fp, os.path.basename(fp))
    # archive.close()
    
    import pandas as pd
    
    # # Create a sample DataFrame
    # df = pd.DataFrame({'A': [1, 2, 3], 'B': ['foo', 'bar', 'baz']})
    
    # # Filter rows where column 'B' matches 'bar'
    # result = df[df['B'] == 'bar']
    
    # # Print the resulting row(s)
    # print(len(result))
    # print(type(result))
    
    
    print(float('Nan'))