# -*- coding: utf-8 -*-
"""
VevoLab Data Extraction Helper
@author: Chris Ward (C) 2021


This program will ask the user for exported data from VevoLab as well as
metadata and analysis settings information

inputs...
report_path : a list of the paths to all VevoLab files that contain data
    (GUI tool helps user to select these files)

metadata_path  : path to the excel file (.xlsx) that contains metadata and
    settings - the file should contain at least 2 sheets ...
    ...[DerivedData,ColumnNames]...
    (GUI tool helps user to select this file)
    *animal data - table with columns for between subjects metadata - MUST
        CONTAIN 'Animal ID' column as this is the key used to link this data 
        with the VevoLab Data
    *timepoint data - table allowing the user to categorize the timepoints in 
        their study with a column for the Timepoint and a column for the data
        (date values need to match the dates found in the VevoLab data in order 
        to link with this table)
    *DerivedData - table indicating if any calculations of derived data such
        as age, post treatment time, or time in study ar needed. If the 
        derived data are calculated, they may be used in the stats/graphs. If
        neccessary data for their calculation are missing an error will be 
        raised when attempting to calculate the values, and the stats/graphs
        may not be created
    *ColumnNames - table indicating preferred name for columns in the output
        data and the name to expect in the VevoLab data. Typos will create
        problems.
    *model - single column table indicating the factors to use for statistical
        analysis and for clustering data on plots
        (default sorting order of clustering variables is ascending, which 
        populates the graph from bottom up)

output_path : path to folder and filename for the excel summary that will be 
    produced - this will also be a prefix for the graph files that are produced
    (GUI tool helps user to select/name this file)
    
    
!Warnings! - it is possible to create errors if columns in the data have names
that include reserved 'patterns' of characters
Currently reserved patterns:
    * '__F{}__'

***CHANGE LOG***
V3.0 -> 4.0
    * fix for data collected from measurement not creating an average if
        multiple instances of the measurement are performed (AutoLV impacted by
        this bug, prior version would only report last entry in the measurement
        table - need to implement an aggregation and average function)
    * update to behavior - if user provides a metadata file requesting an 
        outcome measure that wasn't present in the vevolab report an empty
        column will be included in the report for that measure (instead of an 
        ungracious crash). If the user fails to provide a metadata file, the
        default behavior will be to scan the vevolab report for all contents of 
        the measurement and calculations table (and expected fields) and
        provide an output that includes those outcome measures with their 
        default names, and a template if the user wishes to create modified 
        column names.

V2.0 -> 3.0
    * minor updates to address packaging for distribution
    * v3.0 placed in github - code edits to be tracked in git hub following v4

V 1.0 -> 2.0
    * modified import of measurement data to accomodate AUTOLV data or other 
      measurements that are labeled with similar prefixes but end with an 
      unpredicatable number
                ...py
                if FLAG_measurements == 1:
                    report_dict[rows[0]][
                        '_'.join(columns[0:3])] = columns[4]
                    
                changed to ...
                if FLAG_measurements == 1:
                    if columns[0].isdigit():
                        columns[0]=re.search(
                            '(?P<text>.*?)(?P<digit>\d+$)',
                            columns[0]
                            ).group('text')
                    report_dict[rows[0]][
                        '_'.join(columns[0:3])] = columns[4]
    * modified import of measurement data to accomodate renamed series 
                Removed ...
                if 'Series' not in rows[0]:
                    continue
                            
"""

__version__ = '4.1'


#%% import libraries
import tkinter
import tkinter.filedialog
import pandas
import re
import numpy
import scipy
import pingouin
import itertools
import logging
import traceback
import os
import sys
import datetime

#%% define functions

def guiOpenFileName(kwargs={}):
    """Returns the path to the files selected by the GUI.
*Function calls on tkFileDialog and uses those arguments
  ......
  (declare as a dictionairy)
  {"defaultextension":'',"filetypes":'',"initialdir":'',...
  "initialfile":'',"multiple":'',"message":'',"parent":'',"title":''}
  ......"""
    root = tkinter.Tk()
    outputtext = tkinter.filedialog.askopenfilename(
        **kwargs)
    root.destroy()
    return outputtext


def guiOpenFileNames(kwargs={}):
    """Returns the path to the files selected by the GUI.
*Function calls on tkFileDialog and uses those arguments
  ......
  (declare as a dictionairy)
  {"defaultextension":'',"filetypes":'',"initialdir":'',...
  "initialfile":'',"multiple":'',"message":'',"parent":'',"title":''}
  ......"""
    root = tkinter.Tk()
    outputtextraw = tkinter.filedialog.askopenfilenames(
        **kwargs)
    outputtext = root.tk.splitlist(outputtextraw)
    root.destroy()
    return outputtext


def guiSaveFileName(kwargs={}):
    """Returns the path to the filename and location entered in the GUI.
*Function calls on tkFileDialog and uses those arguments
  ......
  (declare as a dictionairy)
  {"defaultextension":'',"filetypes":'',"initialdir":'',...
  "initialfile":'',"multiple":'',"message":'',"parent":'',"title":''}
  ......"""
    root = tkinter.Tk()
    outputtext = tkinter.filedialog.asksaveasfilename(
        **kwargs)
    root.destroy()
    return outputtext


def log_info_from_dict(local_logger, input_dict, log_prefix=''):
    """
    creates log entries from dict input

    Parameters
    ----------
    local_logger : instance of a logging.logger

    input_dict : dict {k:message,...}
        dict contents will be used to populate log entries

    log_prefix : string, optional
        The default is ''. string is placed ahead of dict values in the
        log entry

    Returns
    -------
    None.

    """

    for k in input_dict:
        local_logger.info('{}{} : {}'.format(log_prefix, k, input_dict[k]))


def scan_for_column_names(report_paths):
    """
    

    Parameters
    ----------
    report_paths : list of strings
        list of filepaths to check for cancidate column names within

    Returns
    -------
    column_names : dict of lists
        dict with 2 entriesDESCRIPTION.
            'MetaData Fields' - fields that are likely metadata containing
            'VevoLab Measurement_Mode_Parameter or Calculation' - fields that
                appear to contain measurements of calculations


    """
    
    column_names = {}
    column_names['MetaData Fields'] = []
    column_names['VevoLab Measurement_Mode_Parameter or Calculation'] = []
    # iterate through files
    for f in report_paths:
        # open files
        with open(f,'r') as opfi:
            report_text = opfi.read().replace('"','')

            # iterate through series
            for b in report_text.split('Series Name,'):
                # iterate through sections
                rows = []
                rows = b.split('\n')
                
                # iterate through rows
                FLAG_calculations = 0
                FLAG_measurements = 0
 
                for r in rows:
                    columns = []
                    columns=r.split(',')
                    
                    # collect calculation values
                    if columns[0] == '':
                        FLAG_calculations = 0
                        FLAG_measurements = 0
                        continue
                
                    # check and set flags for whether the line indicates 
                    # transition between calculation, measurement, or other 
                    # section
                    elif columns[0] == 'Calculation':
                        FLAG_calculations = 1
                        FLAG_measurements = 0
                        continue
                
                    elif columns[0] == 'Measurement':
                        FLAG_measurements = 1
                        FLAG_calculations = 0
                        continue
                
                    if FLAG_calculations == 1:
                      column_names[
                        'VevoLab Measurement_Mode_Parameter or Calculation'
                        ].append(columns[0])
                    if FLAG_measurements == 1:
                        # screen for cases of measurements with number suffix
                        if columns[0][-1].isdigit():
                            # if measurement is number suffixed, grab the 
                            # initial portion
                            columns[0]=re.search(
                                '(?P<text>.*?)(?P<digit>\d+$)',
                                columns[0]
                                ).group('text')
                        column_names[
                          'VevoLab Measurement_Mode_Parameter or Calculation'
                          ].append('_'.join(columns[0:3]))
                    if FLAG_calculations == 0 and FLAG_measurements == 0:    
                        column_names['MetaData Fields'].append(
                            columns[0]
                            )
                        
    # clean up columns names to remove duplicates
    for key in column_names:
        column_names[key] = list(set(column_names[key]))
    
    return column_names


def create_metadata_template(report_paths,template_output_path,logger):
    column_names = scan_for_column_names(report_paths)
    
    template = column_names[
        'VevoLab Measurement_Mode_Parameter or Calculation'
        ]
    template.sort()
    template_df = pandas.DataFrame(
        {
            'VevoLab Measurement_Mode_Parameter or Calculation':template,
            'Output Name':template
            }
        )
    template_writer=pandas.ExcelWriter(
        template_output_path,engine='xlsxwriter'
        )
    logger.info(f'Preparing Column Metadata File:\n{template_output_path}')
    template_df.to_excel(template_writer,'ColumnNames',index=False)
    template_writer.save()
    return template_output_path
    
    
    
#%% main
#%%

def main():
    #%%
    # get input files (report, groups, analysis settings)
    report_path = guiOpenFileNames({'title':'Select VevoLab Report'})
    # get study metadata
    metadata_path = guiOpenFileName({'title':'Select Metadata File'})
    
    
    # set output file (xlsx)
    output_path = guiSaveFileName(
            {
                'title':'Select Name and Location to Save Output',
                'defaultextension':'.xlsx',
                'filetypes':[('excel file','*.xlsx')]
                }
            )

    logger = logging.getLogger('VevoLab Data Extraction Helper')
    logger.setLevel(logging.DEBUG)

    # create file and console handlers to receive logging
    console_handler = logging.StreamHandler(sys.stdout)
    file_handler = logging.FileHandler(
        os.path.join(output_path, 'VLDEH_ERROR.log'), delay = True
        )
    
    console_handler.setLevel(logging.DEBUG)
    file_handler.setLevel(logging.ERROR)

    # create format for log and apply to handlers
    log_format = logging.Formatter(
            '%(asctime)s | %(name)s | %(levelname)s | %(message)s'
            )
    console_handler.setFormatter(log_format)
    file_handler.setFormatter(log_format)

    # add handlers to the logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    # log initial inputs
    logger.info('Beginning Analysis')
    logger.info('{}'.format(datetime.datetime.now().isoformat()))
    
    log_info_from_dict(
        logger,
        {
            'report path':report_path,
            'metadata path':metadata_path,
            'output path':output_path},
        'input paths : '
        )
    try:
        # if no metadata, create a template
        if metadata_path is None or metadata_path == '':
            logger.info('No Metadata File Selected')
            template_output_path = guiSaveFileName(
                {
                    'title':'Select location to save column metadata template',
                    'defaultextension':'.xlsx',
                    'filetypes':[('excel file','*.xlsx')]
                    }
                )
            
            metadata_path = create_metadata_template(
                report_path,template_output_path,logger
                )
            logger.info('Metadata Template will be used')
    except Exception:
        logger.info('unable to resolve missing metadata file')
    
    
    #% grab metadata
    try:
        animal_data = pandas.read_excel(
            metadata_path,sheet_name = 'animal data', dtype={'Animal ID':str}
            )
        skip_animal_data=0
    except Exception:
        logger.info('No Animal Data Found')
        
        skip_animal_data=1
    
    try:    
        timepoint_data = pandas.read_excel(
            metadata_path,sheet_name='timepoint data'
            )
        skip_timepoint_data=0
    except Exception:
        logger.info('No Timepoint Data Found')
        skip_timepoint_data=1
    
    try:
        model_data = pandas.read_excel(metadata_path,sheet_name='model')
        skip_model_data=0
    except Exception:
        logger.info('No Model Information Found')
        skip_model_data=1
        
    try:
        column_names = pandas.read_excel(
            metadata_path,sheet_name='ColumnNames'
            )
        
    except Exception:
        logger.warning('No Column Names Found - default columns will be used')
        column_names = scan_for_column_names(report_path)
        
    try:    
        derived_data_settings = pandas.read_excel(
            metadata_path,sheet_name='DerivedData'
            )
        skip_derived_data=0
    except Exception:
        logger.info('No Settings For Derived Data Found')
        skip_derived_data=1
    
    #%% grab column name settings
    try:
        ColumnStyles = dict(
            zip(
                column_names['VevoLab Measurement_Mode_Parameter or Calculation'],
                column_names['Output Name']
                )
            )
        
        
        #% grab data from the reports
        primary_df = pandas.DataFrame()
        Study_Name=''
        for current_file in report_path:
            print(current_file)
            with open(current_file,'r') as opfi:
                report_text = opfi.read().replace('"','')
            
            # parse the report into blocks (b), and rows (r) and columns (c)
            report_dict = {}
            for b in report_text.split('Series Name,'):
                
                rows = []
                rows = b.split('\n')
                
                for r in b.split('\n'):
                    columns = []
                    columns=r.split(',')
                    if columns[0] == 'Study Name':
                        Study_Name=columns[1]
                
               
                report_dict[rows[0]] = {}
                FLAG_calculations = 0
                FLAG_measurements = 0
                
                
                
                for r in b.split('\n'):
                    columns = []
                    columns = r.split(',')
                    
                    # collect calculation values
                    if columns[0] == '':
                        FLAG_calculations = 0
                        FLAG_measurements = 0
                        continue
                    
                    # check and set flags for whether the line indicates 
                    # transition between calculation, measurement, or other 
                    # section
                    elif columns[0] == 'Calculation':
                        FLAG_calculations = 1
                        FLAG_measurements = 0
                        continue
                    
                    elif columns[0] == 'Measurement':
                        FLAG_measurements = 1
                        FLAG_calculations = 0
                        continue
                    
                    # if row is not a transition indicator, extract data if
                    # row is calculation or measurement data (add to list) -...
                    # take average at the end
                    if FLAG_calculations == 1:
                        report_dict[rows[0]][
                            columns[0]] = columns[3]
                        
                    if FLAG_measurements == 1:
                        # screen for cases of measurements with number suffix
                        if columns[0][-1].isdigit():
                            # if measurement is number suffixed, grab the 
                            # initial portion
                            columns[0]=re.search(
                                '(?P<text>.*?)(?P<digit>\d+$)',
                                columns[0]
                                ).group('text')
                        # place the data
                        if '_'.join(columns[0:3]) in report_dict[rows[0]]:
                            report_dict[rows[0]]['_'.join(columns[0:3])].append(
                                columns[4]
                                )
                        else:
                            report_dict[rows[0]][
                                '_'.join(columns[0:3])] = [columns[4]]
                        
                    if columns[0] == 'Series Date':
                        report_dict[rows[0]][columns[0]] = pandas.to_datetime(
                            columns[1]
                            )
                        
                    if columns[0] == 'Animal ID':
                        report_dict[rows[0]][columns[0]] = columns[1]
                        report_dict[rows[0]]['Study Name'] = Study_Name
                        report_dict[rows[0]]['Series Name']=rows[0]
                    if columns[0] == 'Sex':
                        report_dict[rows[0]][columns[0]] = columns[1]
                        
                    if columns[0] == 'Weight':
                        report_dict[rows[0]][columns[0]] = columns[1]
                        
                         
            # After the end of the data scraping - collapse to a mean() all 
            # entries containing a list of repeated measurements 
            # (affects AutoLV)
            
            for first_key in report_dict:
                for second_key in report_dict[first_key]:
                    if type(report_dict[first_key][second_key]) is list:
                        data_list = []
                        try:
                            data_list = [
                                float(i) for i in 
                                report_dict[first_key][second_key]
                                ]
                            report_dict[first_key][second_key] = \
                            sum(data_list) / len(data_list)
                        
                        except Exception:
                            logging.error(
                                'ERROR: issue summarizing collected data'
                                )
                            logging.error(traceback.format_exc())
                            report_dict[first_key][second_key] = 'ERROR_NA'

                        
            current_df = pandas.DataFrame.from_dict(report_dict,orient = 'index')
            
            current_df = current_df.rename(columns = ColumnStyles)
            output_df_columns = ['Animal ID','Series Date']+list(
                    ColumnStyles.values()
                    )
            
            output_df = current_df[output_df_columns]
            
            if skip_timepoint_data==0:
                output_df = pandas.merge(timepoint_data, output_df, how = 'right', 
                    left_on = 'date', right_on = 'Series Date')
                
            if skip_animal_data==0:
                output_df = pandas.merge(animal_data, output_df, how = 'right',
                    on = 'Animal ID')
    
            primary_df = primary_df.append(output_df)
                
    except Exception:
        logging.error('ERROR: Unable to collect data')
        logging.error(traceback.format_exc())
        
    #%% perform derived data calculations if selected
    # calculate ages
    try:
        if derived_data_settings[
                derived_data_settings['calculation']=='Age(days)'
                ]['Include'].values[0] == 1:
            primary_df['Age(days)']=(
                (primary_df['date']-primary_df['DOB']) / numpy.timedelta64(1,'D')
                ).astype(int)
        
        if derived_data_settings[
                derived_data_settings['calculation']=='Age(wks)'
                ]['Include'].values[0] == 1:
            primary_df['Age(wks)']=(
                (primary_df['date']-primary_df['DOB']) / numpy.timedelta64(7,'D')
                ).astype(int)
        
        if derived_data_settings[
                derived_data_settings['calculation']=='Age(Mo)'
                ]['Include'].values[0] == 1:
            primary_df['Age(Mo)']=(
                (primary_df['date']-primary_df['DOB']) / numpy.timedelta64(28,'D')
                ).astype(int)
    except Exception:
        logging.info('Unable to calculate Age Data')    
        
    # calculate days post treatment
    try:
        if derived_data_settings[
                derived_data_settings['calculation']=='PostTreat(days)'
                ]['Include'].values[0] == 1:
            primary_df['PostTreat(days)']=(
                (primary_df['date']-primary_df['Treatment Date']) / \
                    numpy.timedelta64(1,'D')
                ).astype(int)
        
        if derived_data_settings[
                derived_data_settings['calculation']=='PostTreat(wks)'
                ]['Include'].values[0] == 1:
            primary_df['PostTreat(wks)']=(
                (primary_df['date']-primary_df['Treatment Date']) / \
                    numpy.timedelta64(7,'D')
                ).astype(int)
        
        if derived_data_settings[
                derived_data_settings['calculation']=='PostTreat(Mo)'
                ]['Include'].values[0] == 1:
            primary_df['PostTreat(Mo)']=(
                (primary_df['date']-primary_df['Treatment Date']) / \
                    numpy.timedelta64(28,'D')
                ).astype(int)
    except Exception:
        logging.info('Unable to calculate PostTreatment time')        
        
    # calculate days within study
    try:
        if derived_data_settings[
                derived_data_settings['calculation']=='TimeInStudy(days)'
                ]['Include'].values[0] == 1:
            primary_df['TimeInStudy(days)']=(
                (primary_df['date']-primary_df['Study Start Date']) / \
                    numpy.timedelta64(1,'D')
                ).astype(int)
        
        if derived_data_settings[
                derived_data_settings['calculation']=='TimeInStudy(wks)'
                ]['Include'].values[0] == 1:
            primary_df['TimeInStudy(wks)']=(
                (primary_df['date']-primary_df['Study Start Date']) / \
                    numpy.timedelta64(7,'D')
                ).astype(int)
        
        if derived_data_settings[
                derived_data_settings['calculation']=='TimeInStudy(Mo)'
                ]['Include'].values[0] == 1:
            primary_df['TimeInStudy(Mo)']=(
                (primary_df['date']-primary_df['Study Start Date']) / \
                    numpy.timedelta64(28,'D')
                ).astype(int)
    except Exception:
        logging.info('Unable to calculate Time In Study')    
        
    
    try:
        if skip_model_data==0:
            primary_df = primary_df.sort_values(
                by = list(model_data['factors'].values)+['Animal ID'])
    except Exception:
        logging.info('Unable to calculate Time In Study')    


    #%% prepare summary ouputs
    
    try:    
        if skip_animal_data==0 and \
                skip_timepoint_data==0 and \
                skip_model_data==0 and \
                skip_derived_data==0:
            # repeated measures style output for use with spss
            horiz_split_var=model_data['factors'].values[0]
            horiz_split_values=list(primary_df[horiz_split_var].unique())
            horiz_split_values.sort()
            
            secondary_df=primary_df[primary_df[horiz_split_var]==horiz_split_values[0]]
            secondary_df.columns=[
                '{}_[{}]'.format(c,horiz_split_values[0]) \
                if c not in animal_data.columns \
                else c \
                for c in secondary_df.columns
                ]
            
            for i in range(len(horiz_split_values)-1):
                t=horiz_split_values[i+1]
                
                temp_df=primary_df[primary_df[horiz_split_var]==t][
                    ['Animal ID']+['Series Date']+
                    list(ColumnStyles.values())]
                temp_df.columns=[
                    '{}_[{}]'.format(c,t) if c!='Animal ID' else c 
                    for c in temp_df.columns]
            
                secondary_df=pandas.merge(
                    secondary_df,
                    temp_df,
                    how='outer',
                    on='Animal ID',
                    suffixes=(
                        '_[{}]'.format(horiz_split_values[i]),
                        '_[{}]'.format(t))
                    )
            #% prism style output
            split_var=model_data['factors'].values[-1]
            group_splits=list(secondary_df[split_var].unique())
            group_splits.sort()
            
            col_split=re.compile('(((?P<col>.+)_\[(?P<tp>.*)\])|((?P<alt>.+)))')
            
            tertiery_df=secondary_df[secondary_df[split_var]==group_splits[0]]
            new_cols=[]
            for c in tertiery_df.columns:
                temp_re=re.search(col_split,c)
                if temp_re['col'] is not None:
                    new_cols.append('{}_[{}]_[{}]'.format(
                    re.search(col_split,c)['col'],
                    group_splits[0],
                    re.search(col_split,c)['tp']))
                else:
                    new_cols.append('{}_[{}]'.format(
                        c,group_splits[0]))
            tertiery_df.columns=new_cols
                        
            
            for i in range(len(group_splits)-1):
                g=group_splits[i+1]
                
                temp_df=secondary_df[secondary_df[split_var]==g]
                new_cols=[]
                for c in temp_df.columns:
                    temp_re=re.search(col_split,c)
                    if temp_re['col'] is not None:
                        new_cols.append('{}_[{}]_[{}]'.format(
                        re.search(col_split,c)['col'],
                        g,
                        re.search(col_split,c)['tp']))
                    else:
                        new_cols.append('{}_[{}]'.format(
                            c,g))
                temp_df.columns=new_cols  
                
                tertiery_df=pandas.concat(
                    [tertiery_df,temp_df],
                    axis=1,
                    sort=True # added because of future warning
                    )
            tc=list(tertiery_df.columns)    
            tc.sort()
            tertiery_df=tertiery_df[tc]
            tertiery_df=tertiery_df.fillna('')
            
        #% prepare for excel export
        stats_df=pandas.DataFrame()
        graphs_df=pandas.DataFrame()
        
        writer=pandas.ExcelWriter(output_path,engine='xlsxwriter')
        
        try:
            if derived_data_settings[
                    derived_data_settings['calculation']=='KOMP_STYLE'
                    ]['Include'].values[0] == 1:
                primary_df=primary_df.rename(columns={'Animal ID':'Animal_ID','Series Date':'Study_Date'})
                primary_df['Respiration']=''
                primary_df.loc[:,'Study_Date']=primary_df['Study_Date'].dt.strftime('%d-%b-%y')
                
                primary_df.to_csv(output_path+'.csv', index=False)
        except:
            pass
        
        primary_df.to_excel(writer,'vertical', index=False)
         
        
        if skip_animal_data==0 and \
                skip_timepoint_data==0 and \
                skip_model_data==0 and \
                skip_derived_data==0:
            secondary_df.to_excel(writer,'horizontal', index=False)
            tertiery_df.to_excel(writer,'split', index=False)
            graphs_df.to_excel(writer,'graphs', index=False)
            worksheet=writer.sheets['graphs']
            
            #% run stats
            # get list of independent factors
            ind_vars=list(model_data['factors'].values)
            iv_dict={}
            iv_dict_rev={}
            
            # prepare stats dataframe to be used for easy export
            stats_df=pandas.DataFrame()
            pairwise_df=pandas.DataFrame()
            
            # prepare key for independent factors 
                # - column names #note reserved format style
            for k in range(len(ind_vars)):
                iv_dict[ind_vars[k]]='__F{}__'.format(k)
                iv_dict_rev['__F{}__'.format(k)]=ind_vars[k]
            
            # iterate through outcome measure columns and clean data for ANOVA
            counter=0
            for c in ColumnStyles.values():
                temp_df=primary_df[[c]+list(model_data['factors'].values)]
                temp_df.loc[:,c]=pandas.to_numeric(temp_df[c], errors='coerce')
                temp_df=temp_df.dropna()
                
                temp_df['om']=temp_df[c]
                temp_df['gp']=''
                
                for k in iv_dict_rev:
                    temp_df[k]=temp_df[iv_dict_rev[k]]
                    temp_df['gp']+=temp_df[k].astype(str)
                
                homosced=pingouin.homoscedasticity(
                        temp_df,dv='om',group='gp')['pval'].values[0]
                try:
                    normal=pingouin.normality(
                            temp_df,dv='om',group='gp')['pval'].values[0]
                except:
                    normal='unable to test'
                table=temp_df.anova('om',between=list(iv_dict_rev.keys()),ss_type=3)
                table['levene pval']=str(homosced)
                table['shapiro pval']=str(normal)
                
                
                # replace independent factor placeholders with original names
                for k in iv_dict_rev:
                    table['Source']=table['Source'].replace(
                            k,iv_dict_rev[k],regex=True)
                table['outcome_measure']=c
                stats_df=stats_df.append(table)
                
                # create data frame to assist with summary plot generation 
                #   (uses pandas agg function)
                agg_df=temp_df.groupby(ind_vars).agg(
                        [numpy.mean,len,scipy.stats.sem]).reset_index()
                agg_cols=agg_df.columns
                agg_df.columns=[i[0] if i[0]!=c else i[1] for i in agg_cols]
                agg_df['axis']=agg_df[ind_vars].astype(str).agg('_'.join,axis=1)
            
            
                temp_plot=agg_df.plot(
                        kind='barh',
                        title=c+' [mean+/-sem]',
                        legend=True,
                        y='mean',
                        x='axis')
                temp_plot.errorbar(agg_df['mean'],agg_df['axis'],xerr=agg_df['sem'],
                                   ecolor='black',linewidth=0,elinewidth=1,capsize=4)
                temp_plot.set(xlabel=c,ylabel='_'.join(ind_vars))
                temp_plot=temp_plot.get_figure()
            
                temp_plot.savefig(
                    output_path+'_'+re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', '', c)+'.png', 
                    bbox_inches='tight')
                worksheet.insert_image(
                    'B{}'.format(2+counter*20),
                    output_path+'_'+re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', '', c)+'.png')
                counter+=1
                
                # produce pairwise comparisons
                pairwise_list=[]
                
                for i in range(len(iv_dict_rev)):
                    pairwise_list+=list(itertools.combinations(iv_dict_rev,i+1))
                
                
                for i in pairwise_list:    
                    if len(i)>1:
                        temp_df['*'.join(i)]=temp_df[
                            [j for j in i]
                            ].astype(str).agg(' * '.join,axis=1)
                    
                    pairs=list(itertools.combinations(temp_df['*'.join(i)].unique(),2))
                    for p in pairs:
            
                        if len(temp_df[temp_df['*'.join(i)]==p[0]])<2 or \
                                len(temp_df[temp_df['*'.join(i)]==p[1]])<2:
                            pairwise_df=pairwise_df.append(
                                pandas.DataFrame(
                                    {'outcome_measure':[c],'comparison':[
                                        ' vs '.join(
                                            [str(q) for q in p]
                                            )
                                    ],'notes':['cannot compare']}))
                            continue
                        
                        temp_p_df=pingouin.ttest(
                            temp_df[temp_df['*'.join(i)]==p[0]]['om'],
                            temp_df[temp_df['*'.join(i)]==p[1]]['om']
                            )
                        temp_np_df=pingouin.mwu(
                            temp_df[temp_df['*'.join(i)]==p[0]]['om'],
                            temp_df[temp_df['*'.join(i)]==p[1]]['om']
                            )
                        pw_stats_df=temp_p_df
                        pw_stats_df.index=['PAIRWISE']
                        pw_stats_df['ttest pval']=temp_p_df['p-val']
                        pw_stats_df['mwu pval']=temp_np_df['p-val'].values[0]
                        pw_stats_df.pop('p-val')
                        pw_stats_df['outcome_measure']=c
                        pw_stats_df['comparison']=' vs '.join([str(q) for q in (p)])
                        pairwise_df=pairwise_df.append(pw_stats_df)
                        
                        
            pairwise_df=pairwise_df.reset_index()
            pairwise_df=pairwise_df[
                    ['outcome_measure','comparison']+ \
                    [j for j in pairwise_df if j not in \
                    ['outcome_measure','comparison','notes']]+['notes']]
                    
            stats_df.to_excel(writer,'stats', index=False)
            pairwise_df.to_excel(writer,'pairwise',index=False)
    except Exception:
        logging.error('unable to process data')
        logging.error(traceback.format_exc())
    try:
        writer.save()
        logging.info('\n\nOutput Saved - {}'.format(output_path))
    except:
        logging.error('\n\nUnable to save file')
        logging.error(traceback.format_exc())
    #%%
    input('All Done - Press Enter to Exit, type "R" to restart')
    
if __name__=="__main__":
    main()
