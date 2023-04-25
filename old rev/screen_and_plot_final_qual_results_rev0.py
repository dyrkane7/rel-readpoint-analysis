# -*- coding: utf-8 -*-
"""
Created on Thu Apr  6 21:02:20 2023

@author: dkane
"""
import pandas as pd
import plotly.graph_objects as go
import numpy as np
from datetime import datetime as dt
import os 
import glob
import re

# This script plots parametric test results versus readpoint for all samples
#
# Expected directory structure:
#      <base_path>\<stress>\<file_name>.csv
#
# Expected filename format:
#      <device-name>_<package-type>_<stress>_<readpoint>_Test_<date-of-test>_<time-of-test>.csv
#      ex. "56GPDL2_TO39_HTOL_168HR_Test_010523_083023.csv"
#
# User must specify the following:
#      -Base test data directory
#      -Device name
#      -Stress
#      -Package type
        
# ~~~~~~~~~~ User Config ~~~~~~~~~~ #
base_path = r"C:\Users\dkane\OneDrive - Presto Engineering\Documents\AMF\56G PD Quals\Lot #2\test data" + "\\"

dev = "56GPDL2"
# dev = "56GPD"
# dev = "70GPD"
# dev = "SM3"

pkg_list = [
    "SM", # SM fiber pigtail
    "TO39",
    # "CDIP28"
    ]

stress_list = [
    "HTOL",
    "THB",
    "DH",
    "TMCL",
    # "HTS"
    ]

sn_to_skip = [] # ex. ["75", "11"]
# ~~~~~~~~ End User Config ~~~~~~~~~ #


    

# seperate excel file with plots for each combination of stress and package type
# base path is directory path containing all 
class device_data:  
    def __init__(self, base_path, dev, pkg_list, stress_list):
        self.stress_options = ["HTOL", "THB", "TMCL", "DH", "HTS"]
        
        self.check_dir_structure()
        self.stress_list = self.get_stress_list()
        self.pkg_list = pkg_list
        self.dev = dev
        self.base_path = base_path
        
        # self.rps = self.get_rps()
        # self.src_fps = self.get_src_fps()
        # self.dst_fps = self.get_dst_fps()
        # self.print_src_fps()
        # self.print_dst_fps()
        # self.src_data = self.get_src_data()
    
    def check_dir_structure(self):
        dev = ""
        stress_list = []
        for stress in self.stress_options:
            if os.path.isdir(base_path + "\\" + stress):
                stress_list.append(stress)
                match_files = glob.glob("\\".join([base_path, stress, "*_Test_*.csv"]))
                rp_list = []
                for fp in match_files:
                    fn = os.path.basename(fp)
                    splits = fn.split('_')
                    assert splits[3] not in rp_list, f"found duplicate readpoint result ({splits[3]}) for {stress}"
                    rp_list.append(splits[3])
                    assert len(splits) == 7, f"expected 7 fields serperated by '_' but found {len(splits)} for file {fn}"
                    if not dev:
                        dev = splits[0]
                    assert dev == splits[0], f"expected device {dev} but found {splits[0]} for file {fn}"
                    assert stress == splits[2], f"expected {stress} but found {splits[2]} in file {fn}"
        assert stress_list, f"found no valid stress directory in base dir {base_path}. Options: {self.stress_options}"
        
    def get_stress_list(self):
        stress_list = []
        for stress in self.stress_options:
            if os.path.isdir(base_path+ "\\" + stress):
                stress_list.append(stress)
    
    
    def get_rp_from_fp(self, fp):
        fn = fp.split("\\")[-1]
        rp = fn.split("_")[3]
        return rp
        
    def get_rp_hours_from_fp(self, fp):
        rp = self.get_rp_from_fp(fp)
        hours = int(re.sub("[^0-9]", "", rp)) # remove non-numeric chars to get raw hours (ex "168HR"->168, "T0"->0)
        return hours
 
    def get_src_data(self):
        src_data = {pkg:{stress:{} for stress in self.stress_list} for pkg in self.pkg_list}
        for pkg in self.pkg_list:
            for stress in self.stress_list:
                for rp in self.rps[pkg][stress]:
                    src_data[pkg][stress][rp] = pd.read_csv(self.src_fps[pkg][stress][rp])
        return src_data
    
    # def format_dst_data(self):
        
    
    def sort_sn_order(self, src_data):
        for pkg in self.pkg_list:
            for stress in self.stress_list:
                ref_sn_list = list(src_data[pkg][stress]['T0'].loc[:,"DUT_SN"])
                ref_fp = os.path.basename(self.src_fps[pkg][stress]['T0'])
                for rp in self.rps[pkg][stress][1:]:
                    sn_list = list(src_data[pkg][stress][rp].loc[:,"DUT_SN"])
                    fp = os.path.basename(self.src_fps[pkg][stress][rp])
                    assert set(ref_sn_list) == set(sn_list), f"2 files have different set of SN: {ref_fp} and {fp}"
                    assert len(set(ref_sn_list)) == len(ref_sn_list), f"file ({ref_fp}) contains duplicate SN(s)"
                    assert len(set(sn_list)) == len(sn_list), f"file ({fp}) contains duplicate SN(s)"
                    src_data[pkg][stress][rp].sort_values(
                        by=["DUT_SN"], inplace = False, key=lambda col: col.map({sn:order for order, sn in enumerate(ref_sn_list)}))
    
    # returns dict with format {<pkg> : {<stress> : [sorted readpoint list]}, ... , ... }
    def get_rps(self):
        rps = {pkg:{stress:{} for stress in self.stress_list} for pkg in self.pkg_list}
        for pkg in self.pkg_list:
            for stress in self.stress_list:
                stress_base_path = self.base_path + stress + "\\"
                base_filename = self.dev + "_" + pkg + "_" + stress + "_"
                base_filepath = stress_base_path + base_filename
                match_files = glob.glob(base_filepath + "*_Test_*.csv")
                filepaths = sorted(match_files, key=self.get_rp_hours_from_fp)
                rps[pkg][stress] = [fp.split('_')[-4] for fp in filepaths]
        return rps
        
    def get_src_fps(self):
        src_fps = {pkg:{stress:{} for stress in self.stress_list} for pkg in self.pkg_list}
        for pkg in self.pkg_list:
            for stress in self.stress_list:
                for rp in self.rps[pkg][stress]:
                    base_fp = self.base_path + stress + "\\" + "_".join([self.dev,pkg,stress,rp,"Test"])
                    match_files = glob.glob(base_fp + "_*.csv")
                    assert len(match_files) == 1, f"Found {len(match_files)} files with base fp: {base_fp}, expected 1"
                    src_fps[pkg][stress][rp] = match_files[0]
        return src_fps
    
    def get_dst_fps(self):
        dst_fps = {pkg:{} for pkg in self.pkg_list}
        for pkg in self.pkg_list:
            for stress in self.stress_list:
                rp = self.rps[pkg][stress][0]
                dst_fps[pkg][stress] = self.src_fps[pkg][stress][rp].rsplit("_", 4)[0] + "_final_with_plots.xlsx"
        return dst_fps
    
    def print_src_fps(self):
        print("Input .csv filenames:")
        for pkg in self.src_fps:
            for stress in self.src_fps[pkg]:
                for rp in self.src_fps[pkg][stress]:
                    print("\t" + os.path.basename(self.src_fps[pkg][stress][rp]))
                print()
                
    def print_dst_fps(self):
        print("Output .xlsx filenames:")
        for pkg in self.dst_fps:
            for stress in self.dst_fps[pkg]:
                print("\t" + os.path.basename(self.dst_fps[pkg][stress]))
            print()

dev_data = device_data(base_path, dev, pkg_list, stress_list)   
# print(dev_data.dst_fps) 
# print(dev_data.src_fps) 

