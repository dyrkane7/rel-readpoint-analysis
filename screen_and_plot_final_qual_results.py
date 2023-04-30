# -*- coding: utf-8 -*-
"""
Created on Thu Apr  6 21:02:20 2023

@author: dkane
"""
import pandas as pd
import xlsxwriter
import plotly.graph_objects as go
import numpy as np
from datetime import datetime as dt
import os 
import glob
import re
import yaml
import py7zr
import sys
import math
import warnings

import timeit
import time

try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader

# Expected source data directory structure:
#      <base_path>\<stress>\<file_name>.csv
# Files in subdirectory for each stress should share the same device and stress type
#
# Directories may contain multiple csv file for the same (dev,pkg,stress) and readpoint,
# except there must be a single csv file for t0
#
# If multiple results exist for a particular combination of serial number and readpoint
# this script will select the param result with minimum %change
#
# Expected filename format:
#      <device-name>_<package-type>_<stress>_<readpoint>_Test_<date-of-test>_<time-of-test>.csv 
#      or
#      <device-name>_<package-type>_<stress>_<readpoint>_Retest_<date-of-test>_<time-of-test>.csv
#      ex. "56GPDL2_TO39_HTOL_168HR_Test_010523_083023.csv"
#
# User must specify the following:
#   .yml device config filename
#
# This script assumes the .yml file is located in a subdirectory named 'config' in the same directory as this script
        
# ~~~~~~~~~~ User Config ~~~~~~~~~~ #

# dev_config_fp = r"C:/Users/dkane/OneDrive - Presto Engineering/Documents/python_scripts/rel-readpoint-analysis/config/56GPDL2_config.yml"
# dev_config_fp = r"C:/Users/dkane/OneDrive - Presto Engineering/Documents/python_scripts/rel-readpoint-analysis/config/SM3_config.yml"

# dev_config_fn = "SM3_config.yml"
dev_config_fn = "56GPDL2_config.yml"

# ~~~~~~~~ End User Config ~~~~~~~~~ #

# seperate excel file with plots for each combination of stress and package type
# base path is directory path containing all 
class DeviceData:  
    def __init__(self, dev_config_fn):
        self.stress_options = ["HTOL", "THB", "TMCL", "DH", "HTS"]
        
        self.config = self.get_dev_config_from_yml_file(dev_config_fn)

        self.rp_dict = self.get_rp_dict()
        self.print_rp_dict()
        
        self.src_fps = self.get_src_fps()
        self.dst_fps = self.get_dst_fps()
        self.print_src_fps()
        self.print_dst_fps()
        self.src_data = self.get_src_data()

        # self.get_optimal_result_series(("56GPDL2", "TO39", "HTOL"), "1000HR", "Dark Current (A) @ -3V")

        print("Formatting dst dataframes...")
        start_time = time.time()
        self.dst_data = self.format_dst_data()
        print(f"format_dst_data() execution time: {time.time() - start_time:.2f} seconds")
        print("Done")
        self.verify_src_params_match_config()
        self.axis_info = self.get_axis_info()
    
    def get_dev_config_from_yml_file(self, dev_config_fn):
        parent_dir = os.path.dirname(os.path.abspath(__file__))
        dev_config_dir = os.path.join(parent_dir, "config")
        dev_config_fp = os.path.join(dev_config_dir, dev_config_fn)
        print("dev_config_fp:", dev_config_fp)
        with open(dev_config_fp, 'r') as file_obj:
            return yaml.load(file_obj, Loader)
        
    # rp_dict = {tup:{rp:[rp_list], }, }
    def get_rp_dict(self):
        rp_dict = {}
        stress_list = []
        for stress in self.stress_options:
            if os.path.isdir(self.config["base_path"] + "\\" + stress):
                stress_list.append(stress)
                pattern = "/".join([self.config["base_path"], stress, "**",  "*_*test*_*.csv"])
                match_files = glob.glob(pattern, recursive=True)
                if match_files:
                    ref_dev, _, ref_stress, _ = os.path.basename(match_files[0]).split('_')[:4]
                    for fp in match_files:
                        fn = os.path.basename(fp)
                        splits = fn.split('_')
                        assert len(splits) == 7, f"expected 7 fields serperated by '_' but found {len(splits)} for file {fn}"
                        dev, pkg, stress, rp = splits[:4]
                        tup = (dev, pkg, stress)
                        assert (dev, stress) == (ref_dev, ref_stress),  f"expected {ref_dev} {ref_stress} file, found {dev} {stress} file"
                        if tup not in rp_dict:
                            rp_dict[tup] = []
                        if rp not in rp_dict[tup]:
                            rp_dict[tup].append(rp)
        assert stress_list, f"found no valid stress directory in base dir {self.config['base_path']}. Options: {self.stress_options}"
        assert rp_dict, f"found no valid rp result files in base dir {self.config['base_path']}"
        for tup in rp_dict:
            rp_dict[tup] = sorted(rp_dict[tup], key=lambda rp: int(re.sub("[^0-9]", "", rp)))
        return rp_dict

    def print_rp_dict(self): 
        # get max test tuple string length
        l_max = max([len(str(tup)) for tup in self.rp_dict])
        dev_set, pkg_set, stress_set = set(), set(), set()
        for tup in self.rp_dict:
            dev_set.add(tup[0])
            pkg_set.add(tup[1])
            stress_set.add(tup[2])
        for dev in dev_set:
            for pkg in pkg_set:
                for stress in stress_set:
                    tup = (dev, pkg, stress)
                    print(f"{str(tup):<{l_max+1}}{self.rp_dict[tup]}")
        print()
    
    def get_rp_from_fp(self, fp):
        fn = fp.split("\\")[-1]
        rp = fn.split("_")[3]
        return rp
        
    def get_rp_hours_from_fp(self, fp):
        rp = self.get_rp_from_fp(fp)
        hours = int(re.sub("[^0-9]", "", rp)) # remove non-numeric chars to get raw hours (ex "168HR"->168, "T0"->0)
        return hours
 
    def get_src_data(self):
        src_data = {tup : {rp : [] for rp in self.rp_dict[tup]} for tup in self.rp_dict}
        for tup in self.src_fps:
            for rp in self.src_fps[tup]:
                for fp in self.src_fps[tup][rp]:
                    src_data[tup][rp].append(pd.read_csv(fp))
        return src_data
    
    def get_src_fps(self):
        src_fps = {tup : {} for tup in self.rp_dict}
        for tup in self.rp_dict:
            dev, pkg, stress = tup
            for rp in self.rp_dict[tup]:
                pattern = self.config["base_path"] + "/" + stress + "/**/" + "_".join([dev, pkg, stress, rp, "*test*_*.csv"])
                match_files = glob.glob(pattern, recursive = True)
                assert len(match_files), f"Found 0 files with base fp: {pattern}, expected 1 or more"
                src_fps[tup][rp] = match_files
                if rp == 'T0':
                    assert len(match_files) == 1, f"Found multiple T0 files with base fp: {pattern}, expected 1"
        return src_fps
    
    def verify_src_params_match_config(self):
        for tup in self.src_data:
            pkg = tup[1]
            config_params = list(self.config['params'][pkg].keys())
            for rp in self.src_data[tup]: # verify all files column names match
                for df in self.src_data[tup][rp]:
                    reverse_cols = list(reversed(df.columns))
                    src_params = []
                    for col in reverse_cols:
                        if "% change of " not in col:
                            break
                        src_params.append(col[len("% change of "):])
                    assert reverse_cols, "reverse_cols is empty"
                    n_params = len(src_params)
                    for j, param in enumerate(src_params):
                        assert reverse_cols[n_params + j] == param, "unexpected mismatch, {reverse_cols[count + j]} != {param}"
                    assert len(src_params) == len(config_params), f"found length mismatch between src and config params for {tup}, {rp}"
                    assert set(src_params) == set(config_params), f"found mismatch between src and config params for {tup}, {rp}"

    # returns tuple (optimal_raw_series, optimal_delta_series)
    # For each device, find the result with minimum %change and build a series
    def get_optimal_result_series(self, tup, rp, param):
        opt_raw_list, opt_delta_list = [], []
        t0_sn_list = self.src_data[tup]["T0"][0].loc[:,"DUT_SN"]
        param_delta = "% change of " + param
        if rp == 'T0':
            opt_raw_series = pd.Series(data = self.src_data[tup]["T0"][0].loc[:,param])
            opt_delta_series = [0] * len(t0_sn_list)
        else:
            for sn in t0_sn_list:
                raw_list, delta_list = [], []
                for df in self.src_data[tup][rp]:
                    rows = df[df["DUT_SN"] == sn]
                    raw_list += list(rows.loc[:, param])
                    delta_list += list(rows.loc[:, param_delta])
                # if len(delta_list) > 1:
                #     print("raw_list:", raw_list)
                #     print("delta_list:", delta_list)
                assert delta_list, f"found no row for sn {sn} in {tup}, {rp} csv files"
                opt_delta = min(delta_list, key=lambda fl_str:abs(float(fl_str)))
                i_opt_delta = delta_list.index(opt_delta)
                opt_raw = raw_list[i_opt_delta]
                
                opt_delta_list.append(opt_delta)
                opt_raw_list.append(opt_raw)
            opt_raw_series = pd.Series(data = opt_raw_list)
            opt_delta_series = pd.Series(data = opt_delta_list)
            
        return (opt_raw_series, opt_delta_series)
                         
    '''
    format dst data from unaltered src data
    '''
    def format_dst_data(self):
        dst_data = {tup : {} for tup in self.rp_dict}
        # src_data = self.sort_sn_order(self.src_data)
        # self.check_lot_order(src_data)
        for tup in self.rp_dict:
            pkg = tup[1]
            for param, param_dict in self.config["params"][pkg].items():
                dst_data[tup][param] = {}
                dst_data[tup][param]["DUT_SN"] = self.src_data[tup]["T0"][0].loc[:,"DUT_SN"]
                dst_data[tup][param]["Lot/Wafer#"] = self.src_data[tup]["T0"][0].loc[:,"Lot/Wafer#"]
                screened_list = ["No" if sn not in param_dict["sn_to_skip"] else "Yes" for sn in dst_data[tup][param]["DUT_SN"]]
                dst_data[tup][param]["Screened Out?"] = pd.Series(data = screened_list)
                for rp in self.rp_dict[tup]:
                    opt_raw_results = self.get_optimal_result_series(tup, rp, param)[0]
                    dst_data[tup][param][" " + rp] = opt_raw_results
                    
                    # if param_dict['axis_type'] == 'log':
                    #     self.dst_data[tup][param][" " + rp] = opt_raw_results.abs()
                    # elif param_dict['axis_type'] == 'linear':
                    #     self.dst_data[tup][param][" " + rp] = opt_raw_results
                    # else:
                    #     raise KeyError(f"unsupported axis type: {param_dict['axis_type']}")
                    
                    # if axis type is log, use magnitude of data. Negative values cause errors
                    # for lot_id in self.get_lot_id_set(param, tup):
                    #     min_val, max_val = self.get_min_max_param_result(tup, param, lot_id)
                    #     if self.is_log_axis(min_val, max_val):
                    #         self.dst_data[tup][param][" " + rp].abs(inplace=True)
                            
                for rp in self.rp_dict[tup]:
                    opt_delta_results = self.get_optimal_result_series(tup, rp, param)[1]
                    dst_data[tup][param][rp + " "] = opt_delta_results
                dst_data[tup][param] = pd.DataFrame(data = dst_data[tup][param])
                # print(dst_data[tup][param])
        return dst_data
    
    # def is_log_axis(self, min_val, max_val):
    #     if min_val > 0 and max_val > 0 and max_val/min_val > 1000:
    #         return True
    #     if min_val < 0 and max_val < 0 and min_val/max_val > 1000:
    #         return True
    #     if (max_val > 0 and min_val < 0):
            
    #     if (max_val == 0 or min_val == 0):
    #         print("Warning: ")
    #         return False
    #     return False
        
    # def get_log_base(self, min_val, max_val):
    #     assert self.is_log_axis(min_val, max_val), "Expected log axis but is_log_axis() returned False"
    #     if min_val > 0 and max_val > 0 and max_val/min_val > 1000:
    #         return 1 + (max_val / min_val) ** (1/3)
    #     if min_val < 0 and max_val < 0 and min_val/max_val > 1000:
    #         return 1 + (min_val / max_val) ** (1/3)
    #     if (max_val >= 0 and min_val <= 0):
            
    #         return True
    #     return False
    
    def get_min_max_param_result(self, tup, param, lot_id, use_magnitude=False):
        minimums, maximums = [], []
        pkg = tup[1]
        sn_to_skip = self.config['params'][pkg][param]["sn_to_skip"]
        for rp in self.rp_dict[tup]:
            if use_magnitude:
                param_data = list(self.dst_data[tup][param].loc[:," " + rp].abs())
            else:
                param_data = list(self.dst_data[tup][param].loc[:," " + rp])
            lot_list = list(self.dst_data[tup][param].loc[:,"Lot/Wafer#"])
            screened_param_data = []
            for lot, val in zip(lot_list, param_data):
                if val not in sn_to_skip and lot == lot_id:
                    screened_param_data.append(val)
            minimums.append(min(screened_param_data)) 
            maximums.append(max(screened_param_data))
        return min(minimums), max(maximums)
    
    # lot index is 0-indexed
    def get_chart_cell(self, hor_pos, vert_pos, param_dict, data):
        chart_ycells = 16
        n_readpoints = int((data.shape[1] - 3) / 2)
        num_shown_below = 0
        col_num = 4 + 2*n_readpoints + hor_pos*8
        # if hor_pos > 0:# compensate for wide cell with highlighting rules
        #     col_num -= 2 # compensate for wide cell with highlighting rules
        if vert_pos == 0:
            row_num = 3
        elif vert_pos > 0:
            for i, sn in enumerate(data.loc[:,"DUT_SN"]):
                # print("here 1")
                if sn not in param_dict["sn_to_skip"]:
                    # print("here 2")
                    num_shown_below += 1
                    if num_shown_below >= chart_ycells*vert_pos:
                        row_num = i + 3
                        break
        if num_shown_below < chart_ycells*vert_pos:
            row_num = len(data.loc[:,"DUT_SN"]) + (chart_ycells - num_shown_below) + 3
        # print("final row:", row_num, "final col:", col_num)
        # print("num_shown_below:", num_shown_below)
        # print("chart_ycells*vert_pos:", chart_ycells*vert_pos)
        
        return xlsxwriter.utility.xl_col_to_name(col_num) + str(row_num)

    def get_lot_id_set(self, param, tup):
        sn_list = self.dst_data[tup][param].loc[:,"DUT_SN"]
        lot_list = self.dst_data[tup][param].loc[:,"Lot/Wafer#"]
        pkg = tup[1]
        sn_to_skip = self.config['params'][pkg][param]["sn_to_skip"]
        return set([lot_id for (lot_id, sn) in zip(lot_list, sn_list) if sn not in sn_to_skip])
    
    def is_any_lot_id_log_axis(self, tup, param):
        for lot_id in self.axis_info[tup][param]:
            if self.axis_info[tup][param][lot_id]["type"] == "log":
                return True
        return False
                
    def get_axis_info(self):
        assert self.dst_data, "found empty dst_data. Need initialized dst_data for get_axis_info()"
        axis_info = {tup : {} for tup in self.rp_dict}
        for tup in self.rp_dict:
            pkg = tup[1]
            for param, param_dict in self.config['params'][pkg].items():
                axis_info[tup][param] = {}
                lot_ids = self.get_lot_id_set(param, tup)
                vert_scale_threshold = 100
                # set axis type
                for lot_id in lot_ids:
                    axis_info[tup][param][lot_id] = {}
                    min_val, max_val = self.get_min_max_param_result(tup, param, lot_id)
                    if max_val == 0 or min_val == 0:
                        axis_type = 'linear'
                        warn_msg = f"found max_val = {max_val}, min_val = {min_val} \
                            for {tup}, lot: {lot_id}, param: {param}. Plotting with linear axis. \
                            Can't calculate log base if min or max val is 0"
                        warnings.warn(warn_msg)
                    if min_val < 0 and max_val > 0:
                        axis_type = 'linear'
                        warn_msg = f"found max_val > 0 ({max_val}), min_val < 0 = ({min_val}) \
                            for {tup}, lot: {lot_id}, param: {param}. Plotting with linear axis."
                        warnings.warn(warn_msg)
                    elif max_val/min_val > vert_scale_threshold or min_val/max_val > vert_scale_threshold:
                        axis_type = 'log'
                    else:
                        axis_type = "linear"
                    axis_info[tup][param][lot_id]['type'] = axis_type
                    # if min_val > 0 and max_val > 0 and max_val/min_val > vert_scale_threshold:
                        # final_min_val = min_val
                        # final_max_val = max_val
                        # axis_type = 'log'
                        # log_base = 1 + (max_val / min_val) ** (1/3)
                    # elif min_val < 0 and max_val < 0 and min_val/max_val > vert_scale_threshold:
                        # final_min_val = abs(max_val)
                        # final_max_val = abs(min_val)
                        # axis_type = 'log'
                        # log_base = 1 + (min_val / max_val) ** (1/3)
                    # elif (max_val > 0 and min_val < 0):
                    #     axis_type = 'log'
                        # final_min_val, final_max_val = self.get_min_max_param_result(tup, param, lot_id, use_magnitude=True)
                        # log_base = 1 + (final_max_val / final_min_val) ** (1/3)
                        # if max_val >= abs(min_val):
                        #     log_base = 1 + (max_val / abs(min_val)) ** (1/3)
                        # elif max_val < abs(min_val):
                        #     log_base = 1 + (abs(min_val) / max_val) ** (1/3)
                    # elif max_val == 0 or min_val == 0:
                        # if sheet_raw_data_is_mag:
                        #     final_min_val, final_max_val = self.get_min_max_param_result(tup, param, lot_id, use_magnitude=True)
                        # else:
                        #     final_min_val = min_val
                        #     final_max_val = max_val
                        # axis_type = 'linear'
                        # log_base = float('Nan')
                        # warn_msg = f"found max_val = {max_val}, min_val = {min_val} \
                        #     for {tup}, lot: {lot_id}, param: {param}. Plotting with linear axis. \
                        #     Can't calculate log base if min or max val is 0"
                        # warnings.warn(warn_msg)
                    # else:
                    #     if sheet_raw_data_is_mag:
                    #         final_min_val, final_max_val = self.get_min_max_param_result(tup, param, lot_id, use_magnitude=True)
                    #     else:
                    #         final_min_val = min_val
                    #         final_max_val = max_val
                    #     axis_type = "linear"
                    #     log_base = float('Nan')
                    # axis_info[tup][param][lot_id]['type'] = axis_type
                    
                # set final min/max values and log base
                sheet_raw_data_is_mag = False
                for lot_id in lot_ids:
                    if axis_info[tup][param][lot_id]['type'] == 'log':
                        sheet_raw_data_is_mag = True
                for lot_id in lot_ids:
                    if sheet_raw_data_is_mag:
                        final_min_val, final_max_val = self.get_min_max_param_result(tup, param, lot_id, use_magnitude=True)
                    else:
                        final_min_val, final_max_val = self.get_min_max_param_result(tup, param, lot_id)
                    axis_info[tup][param][lot_id]['min'] = final_min_val
                    axis_info[tup][param][lot_id]['max'] = final_max_val
                    
                    if axis_info[tup][param][lot_id]['type'] == 'log':
                        log_base = 1 + (final_max_val / final_min_val) ** (1/3)
                    else:
                        log_base = float('nan')
                    axis_info[tup][param][lot_id]['log_base'] = log_base

                        
        # debug start
        # for tup in axis_info:
        #     for param in axis_info[tup]:
        #         for lot_id in axis_info[tup][param]:
        #             info = axis_info[tup][param][lot_id]
        #             print(f"{tup}, {param}, {lot_id}, {info}")
        # print()
        # debug stop
        return axis_info
        
    
    def generate_xlsx_for_stress(self, tup):
        dev, pkg, stress = tup
        with pd.ExcelWriter(self.dst_fps[tup]) as writer:
            wb = writer.book
            for param, param_dict in self.config['params'][pkg].items():
                # write dataframe to excel worksheet
                self.dst_data[tup][param].to_excel(
                    writer, sheet_name = param, startrow = 2, index = False, header = False)
                ws = writer.sheets[param]
                header_format = wb.add_format({"align" : "center"})
                for i, col in enumerate(self.dst_data[tup][param].columns): # write header
                    ws.write(1, i, col, header_format)
                    
                n_readpoints = len(self.rp_dict[tup])
                ws.write(1, 4+2*n_readpoints, f"*Highlighted cells failed test limits (hilim: {param_dict['hilim']}%, lolim: {param_dict['lolim']}%)")
                
                # True or False
                sheet_raw_data_is_mag = self.is_any_lot_id_log_axis(tup, param)
                if sheet_raw_data_is_mag:
                    # if log axis type, write header as magnitude of param: |<param>|
                    ws.merge_range(0, 3, 0, 2 + n_readpoints, '|' + param + '|', header_format)
                    
                    # use magnitude of data for log plot
                    start_col = 3
                    end_col = 2 + n_readpoints
                    start_row = 2
                    end_row = self.dst_data[tup][param].shape[0] + 1
                    for row in range(start_row, end_row+1):
                        for col in range(start_col, end_col+1):
                            cell_value = self.dst_data[tup][param].iloc[row-2,col]
                            formula = f'=ABS({cell_value})'
                            ws.write_formula(row, col, formula)
                else:
                    ws.merge_range(0, 3, 0, 2 + n_readpoints, param, header_format)
                ws.merge_range(0, 3+n_readpoints, 0, 2+2*n_readpoints, "% change of " + param + "*", header_format)
                
                ws.set_column(0,0,10)
                ws.set_column(1,1,12)
                ws.set_column(2,2,14)
                ws.set_column(4 + 2*n_readpoints, 4 + 2*n_readpoints, 10)
                ws.set_column(3, 2 + n_readpoints, 14) 
                ws.set_column(3 + n_readpoints, 2 + 2*n_readpoints, 14)
                ws.freeze_panes(2,0)
                ws.set_zoom(60)
                
                # highlight fail results red
                n_rows = self.dst_data[tup][param].shape[0]
                fail_result_format = wb.add_format()
                fail_result_format.set_bg_color('FFABAB') # highlight fail result red
                cond_format = {
                    "type": "cell",
                    "format": fail_result_format
                }   
                if param_dict["hilim"] not in ["inf", "-inf", "nan"]:
                    cond_format["criteria"] = "greater than"
                    cond_format["value"] = param_dict["hilim"]
                    ws.conditional_format(2, 3+n_readpoints, n_rows+1, 2+2*n_readpoints, cond_format)
                if param_dict["lolim"] not in ["inf", "-inf", "nan"]:
                    cond_format["criteria"] = "less than"
                    cond_format["value"] = param_dict["lolim"]
                    ws.conditional_format(2, 3+n_readpoints, n_rows+1, 2+2*n_readpoints, cond_format)
                # Hide rows for screened out SN
                for i,sn in enumerate(self.dst_data[tup][param].loc[:,"DUT_SN"]):
                    if sn in param_dict['sn_to_skip']:
                        ws.set_row(i+2,None,None,{"hidden":True})
                     
                # add param plots to excel worksheet
                sn_list = self.dst_data[tup][param].loc[:,"DUT_SN"]
                lot_list = self.dst_data[tup][param].loc[:,"Lot/Wafer#"]
                # lot_ids = set([lot_id for (lot_id, sn) in zip(lot_list, sn_list) if sn not in param_dict["sn_to_skip"]])
                lot_ids = self.get_lot_id_set(param, tup)
                # print("lot_ids:", lot_ids)
                for j, plot_type in enumerate(["raw_data", "%_change"]):
                    for k,lot_id in enumerate(lot_ids):
                        chart = wb.add_chart({'type': 'line'})
                        for i, (sn, lot) in enumerate(zip(sn_list, lot_list)):
                            if sn not in param_dict["sn_to_skip"] and lot == lot_id:
                                chart.add_series({
                                    'categories': [ws.get_name(), 1, 3+(j*n_readpoints), 1, 2+((j+1)*n_readpoints)],
                                    'values':     [ws.get_name(), i+2, 3+(j*n_readpoints), i+2, 2+((j+1)*n_readpoints)],
                                    'name': str(sn),
                                    'marker': {'type': 'circle'}
                                })
                        # min_val, max_val = self.get_min_max_param_result(tup, param, lot_id)
                        # low_margin, high_margin = 0.1 * abs(min_val), 0.1 * abs(max_val)
                        #debug start
                        # print("chart type:", plot_type, ", axis type:", param_dict["axis_type"], ", min:", min_val, ", max:", max_val)
                        # assert min_val > 0 and max_val > 0, "found non-negative min/max"
                        #debug stop
                        
                        
                        if plot_type == "raw_data":
                            axis_info = self.axis_info[tup][param][lot_id]
                            min_val = axis_info['min']
                            max_val = axis_info['max']
                            margin = (max_val - min_val) * 0.1
                            unit = param.split("(")[1].split(")")[0]
                            if sheet_raw_data_is_mag:
                                chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - |{param}| vs Readpoint"})
                                if axis_info['type'] == 'log':
                                    # if min_val > 0 and max_val > 0:
                                    y_axis_format_opt = {
                                        'name': unit, 'min':min_val - 0.5*abs(min_val), 'max':max_val + 0.5*abs(max_val), 'log_base': axis_info['log_base']}
                                    # if min_val < 0 and max_val < 0:
                                    #     y_axis_format_opt = {
                                    #         'name': unit, 'min':min_val, 'max':max_val, 'log_base': axis_info['log_base']}
                                    # if min_val < 0 and max_val > 0:
                                    #     # min_magnitude, max_magnitude = self.get_min_max_param_result(tup, param, lot_id, use_magnitude=True)
                                    #     # print("min magnitude:", min_magnitude)
                                    #     y_axis_format_opt = {
                                    #         'name': unit, 'min':min_val, 'max':max_val, 'log_base': axis_info['log_base']}
                                elif axis_info['type'] == 'linear':
                                    y_axis_format_opt = {'name': unit, 'min':min_val - margin, 'max':max_val + margin}
                                    # y_axis_format_opt = {'name': unit}
                                else:
                                    raise KeyError(f"unsupported axis type: {axis_info['type']}")
                            else: # all axis are linear if sheet raw data is not magnitude
                                chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - {param} vs Readpoint"})
                                assert axis_info['type'] == 'linear', f"Expected linear axis, found {axis_info['type']}"
                                y_axis_format_opt = {'name': unit, 'min':min_val - margin, 'max':max_val + margin}
                                # y_axis_format_opt = {'name': unit}
                                
                            # if axis_info['type'] == "log":
                            #     chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - |{param}| vs Readpoint"})
                            #     y_axis_format_opt = {
                            #         'name': unit, 'min':abs(min_val), 'max':abs(max_val), 'log_base': axis_info['log_base']}
                            #         # 'name': unit, 'min':abs(min_val), 'max':abs(max_val), 'log_base': 3}
                            # elif param_dict["axis_type"] == "linear":
                            #     chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - {param} vs Readpoint"})
                            #     y_axis_format_opt = {'name': unit, 'min':min_val, 'max':max_val}
                            # else:
                            #     raise KeyError(f"unsupported axis_type: {param_dict['axis_type']}")
                                
                                
                        elif plot_type == "%_change": # always linear axis
                            chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - % change of {param} vs Readpoint"})
                            y_axis_format_opt = {'name': "% change"}
                        else:
                            raise KeyError(f"unsupported plot_type: {plot_type}")
                        x_axes_format_opt = {'name': "Readpoint", 'label_position': 'low'}
                        chart.set_y_axis(y_axis_format_opt)
                        chart.set_x_axis(x_axes_format_opt)
                        chart.set_legend({'none': True})
            
                        chart_cell = self.get_chart_cell(k, j, param_dict, self.dst_data[tup][param])
                        ws.insert_chart(chart_cell, chart)
    
    # def check_lot_order(self, src_data):
    #     for tup in self.rp_dict:
    #         ref_lot_list = list(src_data[tup]['T0'].loc[:,"Lot/Wafer#"])
    #         ref_fp = os.path.basename(self.src_fps[tup]['T0'])
    #         for rp in self.rp_dict[tup]:
    #             lot_list = list(src_data[tup][rp].loc[:,"Lot/Wafer#"])
    #             fp = os.path.basename(self.src_fps[tup][rp])
    #             assert lot_list == ref_lot_list, f"2 fiels have different lot lists: {ref_fp} and {fp}"
    
    # def sort_sn_order(self, src_data):
    #     sorted_src_data = {tup : {} for tup in self.rp_dict}
    #     for tup in self.rp_dict:
    #         ref_sn_list = list(src_data[tup]['T0'].loc[:,"DUT_SN"])
    #         ref_fp = os.path.basename(self.src_fps[tup]['T0'])
    #         for rp in self.rp_dict[tup]:
    #             sn_list = list(src_data[tup][rp].loc[:,"DUT_SN"])
    #             fp = os.path.basename(self.src_fps[tup][rp])
    #             assert set(ref_sn_list) == set(sn_list), f"2 files have different set of SN: {ref_fp} and {fp}"
    #             assert len(set(ref_sn_list)) == len(ref_sn_list), f"file ({ref_fp}) contains duplicate SN(s)"
    #             assert len(set(sn_list)) == len(sn_list), f"file ({fp}) contains duplicate SN(s)"
    #             sorted_src_data[tup][rp] = src_data[tup][rp].sort_values(
    #                 by=["DUT_SN"], inplace = False, key=lambda col: col.map({sn:order for order, sn in enumerate(ref_sn_list)}))
    #     return sorted_src_data
                
    # xlsx files stored in test data directory one level above src data
    def get_dst_fps(self):
        dst_fps = {}
        for tup in self.rp_dict:
            rp = self.rp_dict[tup][-1]
            basename = os.path.basename(self.src_fps[tup][rp][0])
            # print("basename:", basename)
            dirname = os.path.dirname(self.src_fps[tup][rp][0])
            # print("dirname:", dirname)
            dirname = os.path.split(dirname)[0] # go one directory above source filepath
            # print("dirname:", dirname)
            dst_fps[tup] = os.path.normpath(dirname + '\\' + basename.rsplit("_", 3)[0] + "_screened_and_plotted.xlsx")
            # print("dst_fps[tup]:", dst_fps[tup])
        return dst_fps
    
    def print_src_fps(self):
        print("Input .csv filenames:")
        for tup in self.src_fps:
            for rp in self.src_fps[tup].values():
                for fp in rp:
                    print("\t" + os.path.basename(fp))
            print()
                
    def print_dst_fps(self):
        print("Output .xlsx filenames:")
        for tup in self.rp_dict:
            print("\t" + os.path.basename(self.dst_fps[tup]))
            # print("\t" + self.dst_fps[tup])
        print()
        
def open_file_in_excel(fp):
    # add quotes around any directory name with spaces, or system command wont work
    splits = fp.split('\\')
    tmp = ""
    for split in splits:
        if  (' ' in split) == True:
            split = ('"' + split + '"')
        tmp += (split + "\\")
    xlsx_fp = tmp[0:-1]
    os.system(xlsx_fp)

if __name__ == "__main__":
    dev_data = DeviceData(dev_config_fn)
    # tup = ("56GPDL2", "TO39", "HTOL")

    for tup in dev_data.dst_data:
        print(tup)
        
        # remove excel file if it exists
        try:
            os.remove(dev_data.dst_fps[tup])
        except OSError:
            pass
        
        dev_data.generate_xlsx_for_stress(tup)
        # open_file_in_excel(dev_data.dst_fps[tup])
    
    dirname = os.path.dirname(list(dev_data.dst_fps.values())[0])
    fn = os.path.basename(list(dev_data.dst_fps.values())[0])
    zip_fp = dirname + '\\' + fn.split('_')[0] + "_qual_results_with_plots.7z"
    
    # remove 7z file if it exists
    try:
        os.remove(zip_fp)
    except OSError:
        pass
    with py7zr.SevenZipFile(zip_fp, 'w') as archive:
        for fp in dev_data.dst_fps.values():
            archive.write(fp, os.path.basename(fp))


