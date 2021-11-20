# -*- coding: utf-8 -*-
"""
VevoLab Data Extraction Helper
@author: Chris Ward (C) 2021

Version 2.0

This program will ask the user for exported data from VevoLab as well as
metadata and analysis settings information

inputs...
report_path : a list of the paths to all VevoLab files that contain data
    (GUI tool helps user to select these files)

metadata_path  : path to the excel file (.xlsx) that contains metadata and
    settings - the file should contain at least 3 sheets...
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
V 1.0 -> 2.0
    * modified import of measurement data to accomodate AUTOLV data or other 
      measurements that are labeled with similar prefixes but end with an 
      unpredicatable number
                ...
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


#%% main
#%%

def main():
    # get input files (report, groups, analysis settings)
    report_path = guiOpenFileNames({'title':'Select VevoLab Report'})
    # get study metadata
    metadata_path = guiOpenFileName({'title':'Select Metadata File'})
    # set output file (xlsx)
    output_path = guiSaveFileName(
            {'title':'Select Name and Location to Save Output'}
            )


    #%
    Errors = []
    
    #% grab meta data
    try:
        animal_data = pandas.read_excel(
            metadata_path,sheet_name = 'animal data', dtype={'Animal ID':str}
            )
        skip_animal_data=0
    except Exception:
        Errors.append('No Animal Data Found')
        #Errors.append(traceback.format_exc())
        skip_animal_data=1
    
    try:    
        timepoint_data = pandas.read_excel(
            metadata_path,sheet_name='timepoint data'
            )
        skip_timepoint_data=0
    except Exception:
        Errors.append('No Timepoint Data Found')
        #Errors.append(traceback.format_exc())
        skip_timepoint_data=1
    
    try:
        model_data = pandas.read_excel(metadata_path,sheet_name='model')
        skip_model_data=0
    except Exception:
        Errors.append('No Model Information Found')
        #Errors.append(traceback.format_exc())
        skip_model_data=1
        
    try:
        column_names = pandas.read_excel(metadata_path,sheet_name='ColumnNames')
    except Exception:
        Errors.append('No Column Names Found')
        #Errors.append(traceback.format_exc())
        
    try:    
        derived_data_settings = pandas.read_excel(
            metadata_path,sheet_name='DerivedData'
            )
        skip_derived_data=0
    except Exception:
        Errors.append('No Settings For Derived Data Found')
        #Errors.append(traceback.format_exc())
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
                
#                if 'Series' not in rows[0]:
#                    continue
                
                report_dict[rows[0]] = {}
                FLAG_calculations = 0
                FLAG_measurements = 0
                
                
                
                for r in b.split('\n'):
                    columns = []
                    columns = r.split(',')
                    
                    #collect calculation values
                    if columns[0] == '':
                        FLAG_calculations = 0
                        FLAG_measurements = 0
                        continue
                    
                    elif columns[0] == 'Calculation':
                        FLAG_calculations = 1
                        FLAG_measurements = 0
                        continue
                    
                    elif columns[0] == 'Measurement':
                        FLAG_measurements = 1
                        FLAG_calculations = 0
                        continue
                    
                    if FLAG_calculations == 1:
                        report_dict[rows[0]][
                            columns[0]] = columns[3]
                        
                    if FLAG_measurements == 1:
                        if columns[0][-1].isdigit():
                            columns[0]=re.search(
                                '(?P<text>.*?)(?P<digit>\d+$)',
                                columns[0]
                                ).group('text')
                        report_dict[rows[0]][
                            '_'.join(columns[0:3])] = columns[4]
                        
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
        Errors.append('ERROR: Unable to collect data')
        Errors.append(traceback.format_exc())
        
    #%% perform derived data calculations if selected
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
        Errors.append('Unable to calculate Age Data')    
        #Errors.append(traceback.format_exc())
        
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
        Errors.append('Unable to calculate PostTreatment time')        
        #Errors.append(traceback.format_exc())
    
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
        Errors.append('Unable to calculate Time In Study')    
        #Errors.append(traceback.format_exc())
    
    try:
        if skip_model_data==0:
            primary_df = primary_df.sort_values(
                by = list(model_data['factors'].values)+['Animal ID'])
    except Exception:
        Errors.append('Unable to calculate Time In Study')    
        #Errors.append(traceback.format_exc())
    try:    
        if skip_animal_data==0 and \
                skip_timepoint_data==0 and \
                skip_model_data==0 and \
                skip_derived_data==0:
            #% repeated measures style output
            
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
        
        writer=pandas.ExcelWriter(output_path+'.xlsx',engine='xlsxwriter')
        
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
        Errors.append('unable to process data')
        Errors.append(traceback.format_exc())
    try:
        writer.save()
        Errors.append('\n\nOutput Saved - {}'.format(output_path))
    except:
        Errors.append('\n\nUnable to save file')
        Errors.append(traceback.format_exc())
    #%%
    print('\n'.join(Errors))
    
    input('All Done - Press Enter to Exit')
    
if __name__=="__main__":
    main()
