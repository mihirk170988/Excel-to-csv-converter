#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Jan  2 10:11:01 2022

@author: mihirkarnik
"""

import os
import PySimpleGUI as Sg
import pandas as pd
import warnings

# GUI Layout
layout = [[Sg.Text("Please select one of the below two options:")],
          [Sg.T("         "), Sg.Radio('Single File', "RADIO1", default=False, key="-IN1-")],
          [Sg.T("         "), Sg.Radio('All Excel Files in the same folder', "RADIO1", default=False, key="-IN2-")],
          [Sg.T("         "), Sg.Radio('All Excel Files in all the sub-folders', "RADIO1", default=False, key="-IN3-")],
          [Sg.Text("Enter Input filepath:")],
          [Sg.Input(), Sg.FileBrowse()],
          [Sg.Text("Enter Output filepath:")],
          [Sg.Text("(Keep the below path 'Blank' to keep the excel and csv files in the same place)")],
          [Sg.Input()],
          [Sg.Button("OK"), Sg.Button("Cancel")]]

# Create the window
window = Sg.Window("Create_csv_files", layout)


# Function to get the filepath where the csv file is to be pasted
def get_filedata(string1, string2):
    if string2 == "":
        # Getting the filepath for the 'xlsx', 'xlsb', and 'xlsb' files 
        if string1.endswith('.xlsx') or string1.endswith('.xlsb') or string1.endswith('.xlsm'):
            p = string1[:-5]
            return p
        # Getting the filepath for the 'xls' files
        else:
            p = string1[:-4]
            return p
    else:
        # Getting the filepath for the 'xlsx', 'xlsb', and 'xlsb' files
        if string1.endswith('.xlsx') or string1.endswith('.xlsb') or string1.endswith('.xlsm'):
            p = string1[:-5]
            fnm = p.rfind('/')
            fp = p[fnm + 1:]
            return string2 + '/' + fp
        # Getting the filepath for the 'xls' files
        else:
            p = string1[:-4]
            fnm = p.rfind('/')
            fp = p[fnm + 1:]
            return string2 + '/' + fp


# Function to get the path where the Excel file(s) is(are) stored
def get_data(string1):
    my_dir = string1.rfind('/')
    folder_1 = string1[:my_dir]
    files_1 = os.listdir(folder_1)
    return my_dir, folder_1, files_1


while True:
    warnings.filterwarnings("ignore")
    event, values = window.read()
    # End program if user closes window
    if event == Sg.WIN_CLOSED:
        break
    # End program if user clicks 'Cancel'
    if event == 'Cancel':
        break
    # When output path if left 'Blank'
    if event == 'OK' and (values["-IN1-"] == True or values["-IN2-"] == True or values["-IN3-"] == True) and values[0] != '' and values[1] == '':
        # Single file execution
        if values["-IN1-"]:
            f1 = values[0]
            if f1.endswith('.xlsx') or f1.endswith('.xls') or f1.endswith('.xlsb') or f1.endswith('.xlsm'):
                xlsx = pd.ExcelFile(f1)
                y = xlsx.sheet_names
                for i in y:
                    df1 = pd.read_excel(xlsx, i)
                    t1 = get_filedata(f1, "")
                    df1.to_csv(t1 + '_' + i + '.csv', index=None, header=True)
            else:
                Sg.popup("Incorrect Input File.", title='Error Input')
                continue
        # All excel files in the same folder execution
        elif values["-IN2-"]:
            # noinspection PyBroadException
            try:
                mydir, folder, files = get_data(values[0])
                for f in files:
                    if f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsb') or f.endswith('.xlsm'):
                        xlsx = pd.ExcelFile(folder + '/' + f)
                        y = xlsx.sheet_names
                        for i in y:
                            df1 = pd.read_excel(xlsx, i)
                            t1 = get_filedata(f, folder)
                            df1.to_csv(t1 + '_' + i + '.csv', index=None, header=True)
                    else:
                        pass
            except:
                pass
        # All excel files in all the sub-folders execution
        else:
            # noinspection PyBroadException
            try:
                str1 = []
                mydir, folder, files = get_data(values[0])
                for root, dirs, files1 in os.walk(folder, topdown=False):
                    for name in files1:
                        str1.append(os.path.join(root, name))
                for f in str1:
                    if f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsb') or f.endswith('.xlsm'):
                        xlsx = pd.ExcelFile(f)
                        y = xlsx.sheet_names
                        mydir1 = f.rfind('/')
                        folder1 = f[:mydir1]
                        fnm1 = f[mydir1 + 1:]
                        for i in y:
                            f1 = values[0]
                            df1 = pd.read_excel(xlsx, i)
                            t1 = get_filedata(fnm1, folder1)
                            df1.to_csv(t1 + '_' + i + '.csv', index=None, header=True)
                    else:
                        pass
            except:
                pass
        Sg.popup("All csvs created", title='Success!!')
        continue
    # When output path if provided
    if event == 'OK' and (values["-IN1-"] == True or values["-IN2-"] == True or values["-IN3-"] == True) and values[0] != '' and values[1] != '':
        out1 = values[1]
        # Single file execution
        if values["-IN1-"]:
            f1 = values[0]
            if f1.endswith('.xlsx') or f1.endswith('.xls') or f1.endswith('.xlsb') or f1.endswith('.xlsm'):
                xlsx = pd.ExcelFile(f1)
                y = xlsx.sheet_names
                for i in y:
                    df1 = pd.read_excel(xlsx, i)
                    t1 = get_filedata(f1, out1)
                    df1.to_csv(t1 + '_' + i + '.csv', index=None, header=True)
            else:
                Sg.popup("Incorrect Input File.", title='Error Input')
                continue
        # All excel files in the same folder execution
        elif values["-IN2-"]:
            # noinspection PyBroadException
            try:
                mydir, folder, files = get_data(values[0])
                for f in files:
                    if f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsb') or f.endswith('.xlsm'):
                        xlsx = pd.ExcelFile(folder + '/' + f)
                        y = xlsx.sheet_names
                        for i in y:
                            df1 = pd.read_excel(xlsx, i)
                            t1 = get_filedata(f, out1)
                            df1.to_csv(t1 + '_' + i + '.csv', index=None, header=True)
                    else:
                        pass
            except:
                pass
        # All excel files in all the sub-folders execution
        else:
            # noinspection PyBroadException
            try:
                str1 = []
                mydir, folder, files = get_data(values[0])
                for root, dirs, files1 in os.walk(folder, topdown=False):
                    for name in files1:
                        str1.append(os.path.join(root, name))
                for f in str1:
                    if f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.xlsb') or f.endswith('.xlsm'):
                        xlsx = pd.ExcelFile(f)
                        y = xlsx.sheet_names
                        mydir1 = f.rfind('/')
                        fnm1 = f[mydir1 + 1:]
                        for i in y:
                            df1 = pd.read_excel(xlsx, i)
                            t1 = get_filedata(fnm1, out1)
                            df1.to_csv(t1 + '_' + i + '.csv', index=None, header=True)
                    else:
                        pass
            except:
                pass
        Sg.popup("All csvs created", title='Success!!')
        continue
        # Error message when no radio button is selected and Input path is left 'Blank'
    if event == 'OK' and values["-IN1-"] == False and values["-IN2-"] == False and values["-IN3-"] == False and values[0] == '':
        Sg.popup(
            "Please select one from 'Single File' or 'All Excel files' buttons.\nPlease also enter an Input Filepath",
            title='Select an option and need Input filepath!!')
        continue
    # Error message when Input path is left 'Blank'
    if event == 'OK' and values[0] == '':
        Sg.popup("Please enter an Input filepath.", title='Need Input filepath!!')
        continue
    # Error message when no radio button is selected
    if event == 'OK' and values["-IN1-"] == False and values["-IN2-"] == False and values["-IN3-"] == False:
        Sg.popup("Please select one from 'Single File' or 'All Excel files' buttons.", title='Select an option!!')
        continue
window.close()
