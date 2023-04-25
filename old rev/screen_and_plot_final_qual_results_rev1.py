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
import time

try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader

# Expected directory structure:
#      <base_path>\<stress>\<file_name>.csv
#
# Expected filename format:
#      <device-name>_<package-type>_<stress>_<readpoint>_Test_<date-of-test>_<time-of-test>.csv
#      ex. "56GPDL2_TO39_HTOL_168HR_Test_010523_083023.csv"
#
# User must specify the following:
#   .yml device config file
        
# ~~~~~~~~~~ User Config ~~~~~~~~~~ #

dev_config_fp = r"C:/Users/dkane/OneDrive - Presto Engineering/Documents/python_scripts/AMF/config/56GPDL3_config.yml"

# ~~~~~~~~ End User Config ~~~~~~~~~ #

#TODO compare lot list for each file

# seperate excel file with plots for each combination of stress and package type
# base path is directory path containing all 
class device_data:  
    def __init__(self, dev_config_fp):
        self.stress_options = ["HTOL", "THB", "TMCL", "DH", "HTS"]
        self.config = yaml.load(open(dev_config_fp, 'r'), Loader)
        
        #debug start
        # print(self.config['params']['TO39']['T_ambient (C)']['hilim'])
        # print(self.config['params']['TO39']['T_ambient (C)']['lolim'])
        # hilim = self.config['params']['TO39']['T_ambient (C)']['hilim']
        # lolim = self.config['params']['TO39']['T_ambient (C)']['lolim']
        # assert hilim > 0, "hilim not >0"
        # assert lolim < 0, "hilim not <0"
        # sys.exit()
        #debug stop
        
        # 2 lines below to be removed
        sn_to_skip = []
        self.sn_to_skip = sn_to_skip
        
        self.rp_dict = self.get_rp_dict()
        self.print_rp_dict()
        
        self.src_fps = self.get_src_fps()
        self.dst_fps = self.get_dst_fps()
        self.print_src_fps()
        self.print_dst_fps()
        self.src_data = self.get_src_data()
        
        # print("Formatting dst dataframes...")
        # start_time = time.time()
        self.dst_data = self.format_dst_data()
        # print(f"format_dst_data() execution time: {time.time() - start_time:.2f} seconds")
        # print("Done")
        
        self.verify_src_params_match_config()
        
    
    def get_rp_dict(self):
        rp_dict = {}
        stress_list = []
        for stress in self.stress_options:
            if os.path.isdir(self.config["base_path"] + "\\" + stress):
                stress_list.append(stress)
                match_files = glob.glob("\\".join([self.config["base_path"], stress, "*_Test_*.csv"]))
                for fp in match_files:
                    fn = os.path.basename(fp)
                    splits = fn.split('_')
                    assert len(splits) == 7, f"expected 7 fields serperated by '_' but found {len(splits)} for file {fn}"
                    dev, pkg, stress, rp = splits[:4]
                    tup = (dev, pkg, stress)
                    if tup not in rp_dict:
                        rp_dict[tup] = []
                    assert rp not in rp_dict[tup], f"found duplicate rp ({rp}) for tup ({tup})"
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
        src_data = {
            tup : {
                rp : pd.read_csv(self.src_fps[tup][rp]) for rp in self.rp_dict[tup]
                } for tup in self.rp_dict
            }
        return src_data
    
    def get_src_fps(self):
        src_fps = {tup : {} for tup in self.rp_dict}
        for tup in self.rp_dict:
            dev, pkg, stress = tup
            for rp in self.rp_dict[tup]:
                base_fp = self.config["base_path"] + "\\" + stress + "\\" + "_".join([dev, pkg, stress, rp, "Test"])
                match_files = glob.glob(base_fp + "_*.csv")
                assert len(match_files) == 1, f"Found {len(match_files)} files with base fp: {base_fp}, expected 1"
                src_fps[tup][rp] = match_files[0]
        return src_fps
    
    def verify_src_params_match_config(self):
        for tup in self.src_data:
            pkg = tup[1]
            config_params = list(self.config['params'][pkg].keys())
            for rp in self.src_data[tup]: # verify all files column names match
                reverse_cols = list(reversed(self.src_data[tup][rp].columns))
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

    def format_dst_data(self):
        dst_data = {tup : {} for tup in self.rp_dict}
        src_data = self.sort_sn_order(self.src_data)
        self.check_lot_order(src_data)
        for tup in self.rp_dict:
            pkg = tup[1]
            for param, param_dict in self.config["params"][pkg].items():
                dst_data[tup][param] = {}
                dst_data[tup][param]["DUT_SN"] = src_data[tup]["T0"].loc[:,"DUT_SN"]
                dst_data[tup][param]["Lot/Wafer#"] = src_data[tup]["T0"].loc[:,"Lot/Wafer#"]
                screened_list = ["No" if sn not in param_dict["sn_to_skip"] else "Yes" for sn in dst_data[tup][param]["DUT_SN"]]
                dst_data[tup][param]["Screened Out?"] = pd.Series(data = screened_list)
                for rp in self.rp_dict[tup]:
                    # if axis type is log, use magnitude of data. Negative values cause errors
                    if param_dict['axis_type'] == 'log':
                        dst_data[tup][param][rp + " "] = src_data[tup][rp].loc[:,param].abs()
                    elif param_dict['axis_type'] == 'linear':
                        dst_data[tup][param][rp + " "] = src_data[tup][rp].loc[:,param]
                    else:
                        raise KeyError(f"unsupported axis type: {param_dict['axis_type']}")
                for rp in self.rp_dict[tup]:
                    dst_data[tup][param][" " + rp] = pd.Series(data = [0] * len(screened_list)) \
                        if rp == "T0" else src_data[tup][rp].loc[:,"% change of " + param]
                dst_data[tup][param] = pd.DataFrame(data = dst_data[tup][param])
                # print(dst_data[tup][param])
        return dst_data
       
    def get_min_max_param_result(self, tup, param, param_dict, lot_id):
        minimums, maximums = [], []
        for rp in self.rp_dict[tup]:
            param_data = list(self.src_data[tup][rp].loc[:,param])
            lot_list = list(self.src_data[tup][rp].loc[:,"Lot/Wafer#"])
            # screened_param_data = [val for val in param_data if val not in param_dict["sn_to_skip"]]
            screened_param_data = []
            for lot, val in zip(lot_list, param_data):
                if val not in param_dict["sn_to_skip"] and lot == lot_id:
                    screened_param_data.append(val)
            minimums.append(min(screened_param_data)) 
            maximums.append(max(screened_param_data))
        return min(minimums), max(maximums)
    
    # lot index is 0-indexed
    def get_chart_cell(self, hor_pos, vert_pos, param_dict, data):
        chart_ycells = 16
        n_readpoints = int((data.shape[1] - 3) / 2)
        num_shown_below = 0
        col_num = 4 + 2*n_readpoints + hor_pos * 9
        if hor_pos > 0:# compensate for wide cell with highlighting rules
            col_num -= 3 # compensate for wide cell with highlighting rules
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
                if param_dict['axis_type'] == 'log':
                    # if log axis type, write header as magnitude of param: |<param>|
                    ws.merge_range(0, 3, 0, 2 + n_readpoints, '|' + param + '|', header_format)
                elif param_dict['axis_type'] == 'linear':
                    ws.merge_range(0, 3, 0, 2 + n_readpoints, param, header_format)
                ws.merge_range(0, 3+n_readpoints, 0, 2+2*n_readpoints, "% change of " + param + "*", header_format)
                # limits_str = f"*Highlighted cells failed test limits "
                # if str(param_dict["lolim"]).lower() not in ['inf', '-inf', 'nan']:
                #     limits_str += ", lohim: {param_dict['lolim']}%"
                # if str(param_dict["hilim"]).lower() not in ['inf', '-inf', 'nan']:
                #     limits_str += ", hilim: {param_dict['hilim']}%"
                ws.write(1, 4+2*n_readpoints, f"*Highlighted cells failed test limits (hilim: {param_dict['hilim']}%, lolim: {param_dict['lolim']}%)")
                
                ws.set_column(0,0,9)
                ws.set_column(1,1,11)
                ws.set_column(2,2,13)
                ws.set_column(4+2*n_readpoints,4+2*n_readpoints,36)
                ws.freeze_panes(2,0)
                ws.set_zoom(60)
                
                # highlight fail results red
                n_rows = self.dst_data[tup][param].shape[0]
                fail_result_format = wb.add_format() 
                fail_result_format.set_bg_color('FFABAB') 
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
                lot_ids = set([lot_id for (lot_id, sn) in zip(lot_list, sn_list) if sn not in param_dict["sn_to_skip"]])
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
                        min_val, max_val = self.get_min_max_param_result(tup, param, param_dict, lot_id)
                        
                        #debug start
                        # print("chart type:", plot_type, ", axis type:", param_dict["axis_type"], ", min:", min_val, ", max:", max_val)
                        # assert min_val > 0 and max_val > 0, "found non-negative min/max"
                        #debug stop
                        if plot_type == "raw_data":
                            unit = param.split("(")[1].split(")")[0]
                            if param_dict["axis_type"] == "log":
                                chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - |{param}| vs Readpoint"})
                                y_axis_format_opt = {
                                    'name': unit, 'min':abs(min_val), 'max':abs(max_val), 'log_base':10}
                            elif param_dict["axis_type"] == "linear":
                                chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - {param} vs Readpoint"})
                                y_axis_format_opt = {'name': unit, 'min':min_val, 'max':max_val}
                            else:
                                raise KeyError(f"unsupported axis_type: {param_dict['axis_type']}")
                        elif plot_type == "%_change": # always linear axis
                            chart.set_title({'name': f"{pkg} - {stress} - {lot_id} - {param} vs Readpoint"})
                            y_axis_format_opt = {'name': "% change"}
                        else:
                            raise KeyError(f"unsupported plot_type: {plot_type}")
                        x_axes_format_opt = {'name': "Readpoint", 'label_position': 'low'}
                        chart.set_y_axis(y_axis_format_opt)
                        chart.set_x_axis(x_axes_format_opt)
                        chart.set_legend({'none': True})
            
                        chart_cell = self.get_chart_cell(k, j, param_dict, self.dst_data[tup][param])
                        ws.insert_chart(chart_cell, chart)
    
    # throws error if 
    def check_lot_order(self, src_data):
        for tup in self.rp_dict:
            ref_lot_list = list(src_data[tup]['T0'].loc[:,"Lot/Wafer#"])
            ref_fp = os.path.basename(self.src_fps[tup]['T0'])
            for rp in self.rp_dict[tup]:
                lot_list = list(src_data[tup][rp].loc[:,"Lot/Wafer#"])
                fp = os.path.basename(self.src_fps[tup][rp])
                assert lot_list == ref_lot_list, f"2 fiels have different lot lists: {ref_fp} and {fp}"
    
    def sort_sn_order(self, src_data):
        sorted_src_data = {tup : {} for tup in self.rp_dict}
        for tup in self.rp_dict:
            ref_sn_list = list(src_data[tup]['T0'].loc[:,"DUT_SN"])
            ref_fp = os.path.basename(self.src_fps[tup]['T0'])
            for rp in self.rp_dict[tup]:
                sn_list = list(src_data[tup][rp].loc[:,"DUT_SN"])
                fp = os.path.basename(self.src_fps[tup][rp])
                assert set(ref_sn_list) == set(sn_list), f"2 files have different set of SN: {ref_fp} and {fp}"
                assert len(set(ref_sn_list)) == len(ref_sn_list), f"file ({ref_fp}) contains duplicate SN(s)"
                assert len(set(sn_list)) == len(sn_list), f"file ({fp}) contains duplicate SN(s)"
                sorted_src_data[tup][rp] = src_data[tup][rp].sort_values(
                    by=["DUT_SN"], inplace = False, key=lambda col: col.map({sn:order for order, sn in enumerate(ref_sn_list)}))
        return sorted_src_data
                
    # xlsx files stored in test data directory one level above src data
    def get_dst_fps(self):
        dst_fps = {}
        for tup in self.rp_dict:
            rp = self.rp_dict[tup][-1]
            basename = os.path.basename(self.src_fps[tup][rp])
            dirname = os.path.dirname(self.src_fps[tup][rp])
            dirname = dirname.rsplit('\\', 1)[0] # remove stress directory
            dst_fps[tup] = dirname + '\\' + basename.rsplit("_", 3)[0] + "_screened_and_plotted.xlsx"
        return dst_fps
    
    def print_src_fps(self):
        print("Input .csv filenames:")
        for tup in self.src_fps:
            for rp in self.src_fps[tup]:
                print("\t" + os.path.basename(self.src_fps[tup][rp]))
            print()
                
    def print_dst_fps(self):
        print("Output .xlsx filenames:")
        for tup in self.rp_dict:
            print("\t" + os.path.basename(self.dst_fps[tup]))
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
    dev_data = device_data(dev_config_fp)
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


