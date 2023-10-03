# -*- coding: utf-8 -*-
"""
VDEH_model
"""

__component_version__ = "1.1"
__license__ = "MIT License"

#%% import modules/libraries
from dataclasses import dataclass

import pandas
import re
import numpy
import scipy
import pingouin
import itertools
import logging
import traceback

# import os
# import sys
# import datetime

#%% define functions


def collect_data(report_paths, logger=None):
    """
    Parameters
    ----------
    report_paths : list of strings
        list of filepaths to check for candidate column names within

    Returns
    -------
    column_names : dict of lists
        dict with 2 entries.
            'MetaData Fields' - fields that are likely metadata containing
            'VevoLab Measurement_Mode_Parameter or Calculation' - fields that
                appear to contain measurements of calculations

    """

    column_names = {}
    column_names["MetaData Fields"] = []
    column_names["VevoLab Measurement_Mode_Parameter or Calculation"] = []
    df = pandas.DataFrame()

    # iterate through files
    for f in report_paths:
        # open files
        if logger:
            logger.log("info", f"collecting data from {f}")
        with open(f, "r") as opfi:
            report_text = opfi.read().replace('"', "")
            report_dict = {}
            study_dict = {}
            # iterate through series (split using 'Series Name' - first row
            # will contain text attributable to Series Name)
            for i, b in enumerate(report_text.split("Series Name,")):

                # iterate through sections
                rows = []
                rows = b.split("\n")

                # iterate through rows

                FLAG_calculations = 0
                FLAG_measurements = 0
                FLAG_version = 0

                # prepare the report dict
                report_dict[rows[0]] = {"Series Name":rows[0]}
                for r in rows[1:]:
                    columns = []
                    columns = r.split(",")

                    # populate report_dict, key will be the series name
                    # 1st block will be study info metadata

                    # clear flags indicating calculation or measurement section
                    if columns[0] == "":
                        FLAG_calculations = 0
                        FLAG_measurements = 0
                        FLAG_version = 0
                        continue

                    # check and set flags for whether the line indicates
                    # transition between calculation, measurement, or other
                    # section
                    elif columns[0] == "Calculation":
                        FLAG_calculations = 1
                        FLAG_measurements = 0
                        FLAG_version = 0
                        continue

                    elif columns[0] == "Measurement":
                        FLAG_measurements = 1
                        FLAG_calculations = 0
                        FLAG_version = 0
                        continue

                    elif columns[0] == "Version Information":
                        FLAG_version = 1
                        FLAG_calculations = 0
                        FLAG_measurements = 0
                        continue

                    # if row is not a transition indicator, extract data if
                    # row is calculation or measurement data (add to list)
                    if (
                        FLAG_calculations == 0
                        and FLAG_measurements == 0
                        and FLAG_version == 0
                    ):
                        column_names["MetaData Fields"].append(columns[0])
                        # try:
                        report_dict[rows[0]][columns[0]] = columns[1]
                        # except IndexError: # no 2nd column
                            # report_dict[rows[0]][columns[0]] = '###'
                        if i == 0:
                            # try:
                                study_dict[columns[0]] = columns[1]
                            # except IndexError: # no 2nd column
                                # study_dict[columns[0]] = '###'

                    elif FLAG_version > 0:
                        if FLAG_version == 1:
                            version_header = [c for c in columns]

                        else:
                            study_dict[','.join(version_header)] = \
                                ','.join(columns)
                        FLAG_version += 1

                    elif FLAG_calculations == 1:
                        column_names[
                            "VevoLab Measurement_Mode_Parameter or Calculation"
                        ].append(columns[0])
                        report_dict[rows[0]][columns[0]] = columns[3]

                    elif FLAG_measurements == 1:
                        # screen for cases of measurements with number suffix
                        if columns[0][-1].isdigit():
                            # if measurement is number suffixed, grab the
                            # initial portion
                            columns[0] = re.search(
                                "(?P<text>.*?)(?P<digit>\d+$)", columns[0]
                            ).group("text")
                        column_names[
                            "VevoLab Measurement_Mode_Parameter or Calculation"
                        ].append("_".join(columns[0:3]))
                        # place the data
                        if "_".join(columns[0:3]) in report_dict[rows[0]]:
                            report_dict[rows[0]][
                                "_".join(columns[0:3])
                            ].append(columns[4])
                        else:
                            report_dict[rows[0]]["_".join(columns[0:3])] = [
                                columns[4]
                            ]
                for k, v in study_dict.items():
                    report_dict[rows[0]][k] = v

        for first_key in report_dict:
            for second_key in report_dict[first_key]:
                if type(report_dict[first_key][second_key]) is list:
                    data_list = []
                    try:
                        data_list = [
                            float(i)
                            for i in report_dict[first_key][second_key]
                        ]
                        report_dict[first_key][second_key] = sum(
                            data_list
                        ) / len(data_list)

                    except Exception:
                        logger.log(
                            "error",
                            (
                                "issue summarizing collected data "
                                + "{f}:{first_key} - {second_key}"
                            ),
                        )
                        logger.log("error", traceback.format_exc())
                        report_dict[first_key][second_key] = "ERROR_NA"

        current_df = pandas.DataFrame.from_dict(
            report_dict, orient="index"
        )

        df = pandas.concat([df,current_df],axis=0, join='outer')

    # clean up columns names to remove duplicates
    for key in column_names:
        column_names[key] = list(set(column_names[key]))

    return column_names, df



def simple_export(dict_of_dfs, output_path, logger=None):
    writer = pandas.ExcelWriter(output_path, engine="xlsxwriter")

    try:
        for k,v in dict_of_dfs.items():
            v.to_excel(writer, k, index=False)
        writer.close()
        if logger: logger.log(f'Data Saved to file - {output_path}')
        
    except Exception as e:
        if logger:
            logger.log(
                "error", f"Unable to save output - {e}"
            )
        if logger:
            logger.log("error", traceback.format_exc())
    
    

#%% define classes


@dataclass
class vdeh_model:
    # logging queue:
    logger: None = None

    # paths
    input_paths: list = None
    output_path: str = str()
    settings_path: str = str()

    # settings
    animal_data: pandas.DataFrame = pandas.DataFrame()
    timepoint_data: pandas.DataFrame = pandas.DataFrame()
    derived_data: pandas.DataFrame = pandas.DataFrame()
    column_names: pandas.DataFrame = pandas.DataFrame()
    model_data: pandas.DataFrame = pandas.DataFrame()
    model: pandas.DataFrame = pandas.DataFrame()

    settings_changed: bool = False
    version_info: str = str()
    log_level: str = "INFO"
    log_file_path: str = str()

    def load_logger(self, logger):
        self.logger = logger

    def load_settings_from_file(self):
        try:
            self.animal_data = pandas.read_excel(
                self.settings_path,
                sheet_name="animal data",
                dtype={"Animal ID": str},
            )
        except Exception:
            if self.logger:
                self.logger.log("info", "No Animal Data Found")

        try:
            self.timepoint_data = pandas.read_excel(
                self.settings_path, sheet_name="timepoint data"
            )
        except Exception:
            if self.logger:
                self.logger.log("info", "No Timepoint Data Found")

        try:
            self.model = pandas.read_excel(
                self.settings_path, sheet_name="model"
            )
        except Exception:
            if self.logger:
                self.logger.log("info", "No Model Information Found")

        try:
            self.column_names = pandas.read_excel(
                self.settings_path, sheet_name="column names"
            )
        except Exception:
            if self.logger:
                self.logger.log(
                    "warning",
                    "No Column Names Found - default columns will be used",
                )
            self.column_names,self.model_data = collect_data(
                self.input_paths,self.logger
            )

        try:
            self.derived_data = pandas.read_excel(
                self.settings_path, sheet_name="derived data"
            )
        except Exception:
            if self.logger:
                self.logger.log("info", "No Settings For Derived Data Found")

    def save_settings_to_file(self, new_settings_path):
        writer = pandas.ExcelWriter(
            self.new_settings_path, engine="xlsxwriter"
        )
        if self.animal_data.shape[0] > 0:
            self.animal_data.to_excel(writer, "animal data", index=False)

        if self.timepoint_data.shape[0] > 0:
            self.timepoint_data.to_excel(writer, "timepoint data", index=False)

        if self.derived_data.shape[0] > 0:
            self.derived_data.to_excel(writer, "derived data", index=False)

        if self.column_names.shape[0] > 0:
            self.column_names.to_excel(writer, "column names", index=False)

        if self.model.shape[0] > 0:
            self.model.to_excel(writer, "model", index=False)

        writer.save()

    def check_data(self):
        self.column_names,self.model_data = collect_data(
            self.input_paths,self.logger
        )
        
        

    def generate_full_report(self):
        # grab column name settings
        try:
            ColumnStyles = dict(
                zip(
                    self.column_names[
                        "VevoLab Measurement_Mode_Parameter or Calculation"
                    ],
                    self.column_names["Output Name"],
                )
            )

            #% grab data from the reports
            primary_df = pandas.DataFrame()
            Study_Name = ""
            for current_file in self.report_path:
                if self.logger:
                    self.logger.log("info", f"working on {current_file}")
                with open(current_file, "r") as opfi:
                    report_text = opfi.read().replace('"', "")

                # parse the report into blocks (b), and rows (r) and columns (c)
                report_dict = {}
                for b in report_text.split("Series Name,"):

                    rows = []
                    rows = b.split("\n")

                    for r in b.split("\n"):
                        columns = []
                        columns = r.split(",")
                        if columns[0] == "Study Name":
                            Study_Name = columns[1]

                    report_dict[rows[0]] = {}
                    FLAG_calculations = 0
                    FLAG_measurements = 0

                    for r in b.split("\n"):
                        columns = []
                        columns = r.split(",")

                        # collect calculation values
                        if columns[0] == "":
                            FLAG_calculations = 0
                            FLAG_measurements = 0
                            continue

                        # check and set flags for whether the line indicates
                        # transition between calculation, measurement, or other
                        # section
                        elif columns[0] == "Calculation":
                            FLAG_calculations = 1
                            FLAG_measurements = 0
                            continue

                        elif columns[0] == "Measurement":
                            FLAG_measurements = 1
                            FLAG_calculations = 0
                            continue

                        # if row is not a transition indicator, extract data if
                        # row is calculation or measurement data (add to list) -...
                        # take average at the end
                        if FLAG_calculations == 1:
                            report_dict[rows[0]][columns[0]] = columns[3]

                        if FLAG_measurements == 1:
                            # screen for cases of measurements with number suffix
                            if columns[0][-1].isdigit():
                                # if measurement is number suffixed, grab the
                                # initial portion
                                columns[0] = re.search(
                                    "(?P<text>.*?)(?P<digit>\d+$)", columns[0]
                                ).group("text")
                            # place the data
                            if "_".join(columns[0:3]) in report_dict[rows[0]]:
                                report_dict[rows[0]][
                                    "_".join(columns[0:3])
                                ].append(columns[4])
                            else:
                                report_dict[rows[0]][
                                    "_".join(columns[0:3])
                                ] = [columns[4]]

                        if columns[0] == "Series Date":
                            report_dict[rows[0]][
                                columns[0]
                            ] = pandas.to_datetime(columns[1])

                        if columns[0] == "Animal ID":
                            report_dict[rows[0]][columns[0]] = columns[1]
                            report_dict[rows[0]]["Study Name"] = Study_Name
                            report_dict[rows[0]]["Series Name"] = rows[0]
                        if columns[0] == "Sex":
                            report_dict[rows[0]][columns[0]] = columns[1]

                        if columns[0] == "Weight":
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
                                    float(i)
                                    for i in report_dict[first_key][second_key]
                                ]
                                report_dict[first_key][second_key] = sum(
                                    data_list
                                ) / len(data_list)

                            except Exception:
                                logging.error(
                                    "ERROR: issue summarizing collected data"
                                )
                                logging.error(traceback.format_exc())
                                report_dict[first_key][second_key] = "ERROR_NA"

                current_df = pandas.DataFrame.from_dict(
                    report_dict, orient="index"
                )

                current_df = current_df.rename(columns=ColumnStyles)
                output_df_columns = ["Animal ID", "Series Date"] + list(
                    ColumnStyles.values()
                )

                output_df = current_df[output_df_columns]

                if self.timepoint_data.shape[0] > 0:
                    output_df = pandas.merge(
                        self.timepoint_data,
                        output_df,
                        how="right",
                        left_on="date",
                        right_on="Series Date",
                    )

                if self.animal_data.shape[0] > 0:
                    output_df = pandas.merge(
                        self.animal_data,
                        output_df,
                        how="right",
                        on="Animal ID",
                    )

                primary_df = primary_df.append(output_df)

        except Exception as e:
            if self.logger:
                self.logger.log("error", f"ERROR: Unable to collect data: {e}")
            if self.logger:
                self.logger.log("error", traceback.format_exc())

        # perform derived data calculations if selected
        # calculate ages
        if self.derived_data.shape[0] > 0:
            try:
                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "Age(days)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["Age(days)"] = (
                        (primary_df["date"] - primary_df["DOB"])
                        / numpy.timedelta64(1, "D")
                    ).astype(int)

                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "Age(wks)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["Age(wks)"] = (
                        (primary_df["date"] - primary_df["DOB"])
                        / numpy.timedelta64(7, "D")
                    ).astype(int)

                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "Age(Mo)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["Age(Mo)"] = (
                        (primary_df["date"] - primary_df["DOB"])
                        / numpy.timedelta64(28, "D")
                    ).astype(int)
            except Exception as e:
                if self.logger:
                    self.logger.log(
                        "error", f"Unable to calculate Age Data: {e}"
                    )
                if self.logger:
                    self.logger.log("error", traceback.format_exc())

            # calculate days post treatment
            try:
                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "PostTreat(days)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["PostTreat(days)"] = (
                        (primary_df["date"] - primary_df["Treatment Date"])
                        / numpy.timedelta64(1, "D")
                    ).astype(int)

                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "PostTreat(wks)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["PostTreat(wks)"] = (
                        (primary_df["date"] - primary_df["Treatment Date"])
                        / numpy.timedelta64(7, "D")
                    ).astype(int)

                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "PostTreat(Mo)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["PostTreat(Mo)"] = (
                        (primary_df["date"] - primary_df["Treatment Date"])
                        / numpy.timedelta64(28, "D")
                    ).astype(int)
            except Exception as e:
                if self.logger:
                    self.logger.log(
                        "error", f"Unable to calculate PostTreatment time: {e}"
                    )
                if self.logger:
                    self.logger.log("error", traceback.format_exc())

            # calculate days within study
            try:
                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "TimeInStudy(days)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["TimeInStudy(days)"] = (
                        (primary_df["date"] - primary_df["Study Start Date"])
                        / numpy.timedelta64(1, "D")
                    ).astype(int)

                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "TimeInStudy(wks)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["TimeInStudy(wks)"] = (
                        (primary_df["date"] - primary_df["Study Start Date"])
                        / numpy.timedelta64(7, "D")
                    ).astype(int)

                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "TimeInStudy(Mo)"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df["TimeInStudy(Mo)"] = (
                        (primary_df["date"] - primary_df["Study Start Date"])
                        / numpy.timedelta64(28, "D")
                    ).astype(int)
            except Exception as e:
                if self.logger:
                    self.logger.log(
                        "error", f"Unable to calculate Time In Study: {e}"
                    )
                if self.logger:
                    self.logger.log("error", traceback.format_exc())

        try:
            if self.model.shape[0] > 0:
                primary_df = primary_df.sort_values(
                    by=list(self.model["factors"].values) + ["Animal ID"]
                )
        except Exception as e:
            if self.logger:
                self.logger.log("error", f"Unable to parse data model: {e}")
            if self.logger:
                self.logger.log("error", traceback.format_exc())

        # prepare summary ouputs

        try:
            if all(
                self.animal_data.shape[0] > 0,
                self.timepoint_data.shape[0] > 0,
                self.derived_data.shape[0] > 0,
                self.model.shape[0] > 0,
            ):

                # repeated measures style output for use with spss
                horiz_split_var = self.model["factors"].values[0]
                horiz_split_values = list(primary_df[horiz_split_var].unique())
                horiz_split_values.sort()

                secondary_df = primary_df[
                    primary_df[horiz_split_var] == horiz_split_values[0]
                ]
                secondary_df.columns = [
                    "{}_[{}]".format(c, horiz_split_values[0])
                    if c not in self.animal_data.columns
                    else c
                    for c in secondary_df.columns
                ]

                for i in range(len(horiz_split_values) - 1):
                    t = horiz_split_values[i + 1]

                    temp_df = primary_df[primary_df[horiz_split_var] == t][
                        ["Animal ID"]
                        + ["Series Date"]
                        + list(ColumnStyles.values())
                    ]
                    temp_df.columns = [
                        "{}_[{}]".format(c, t) if c != "Animal ID" else c
                        for c in temp_df.columns
                    ]

                    secondary_df = pandas.merge(
                        secondary_df,
                        temp_df,
                        how="outer",
                        on="Animal ID",
                        suffixes=(
                            "_[{}]".format(horiz_split_values[i]),
                            "_[{}]".format(t),
                        ),
                    )
                #% prism style output
                split_var = self.model["factors"].values[-1]
                group_splits = list(secondary_df[split_var].unique())
                group_splits.sort()

                col_split = re.compile(
                    "(((?P<col>.+)_\[(?P<tp>.*)\])|((?P<alt>.+)))"
                )

                tertiery_df = secondary_df[
                    secondary_df[split_var] == group_splits[0]
                ]
                new_cols = []
                for c in tertiery_df.columns:
                    temp_re = re.search(col_split, c)
                    if temp_re["col"] is not None:
                        new_cols.append(
                            "{}_[{}]_[{}]".format(
                                re.search(col_split, c)["col"],
                                group_splits[0],
                                re.search(col_split, c)["tp"],
                            )
                        )
                    else:
                        new_cols.append("{}_[{}]".format(c, group_splits[0]))
                tertiery_df.columns = new_cols

                for i in range(len(group_splits) - 1):
                    g = group_splits[i + 1]

                    temp_df = secondary_df[secondary_df[split_var] == g]
                    new_cols = []
                    for c in temp_df.columns:
                        temp_re = re.search(col_split, c)
                        if temp_re["col"] is not None:
                            new_cols.append(
                                "{}_[{}]_[{}]".format(
                                    re.search(col_split, c)["col"],
                                    g,
                                    re.search(col_split, c)["tp"],
                                )
                            )
                        else:
                            new_cols.append("{}_[{}]".format(c, g))
                    temp_df.columns = new_cols

                    tertiery_df = pandas.concat(
                        [tertiery_df, temp_df],
                        axis=1,
                        sort=True,  # added because of future warning
                    )
                tc = list(tertiery_df.columns)
                tc.sort()
                tertiery_df = tertiery_df[tc]
                tertiery_df = tertiery_df.fillna("")

            #% prepare for excel export
            stats_df = pandas.DataFrame()
            graphs_df = pandas.DataFrame()

            writer = pandas.ExcelWriter(self.output_path, engine="xlsxwriter")

            try:
                if (
                    self.derived_data[
                        self.derived_data["calculation"] == "KOMP_STYLE"
                    ]["Include"].values[0]
                    == 1
                ):
                    primary_df = primary_df.rename(
                        columns={
                            "Animal ID": "Animal_ID",
                            "Series Date": "Study_Date",
                        }
                    )
                    primary_df["Respiration"] = ""
                    primary_df.loc[:, "Study_Date"] = primary_df[
                        "Study_Date"
                    ].dt.strftime("%d-%b-%y")

                    primary_df.to_csv(self.output_path + ".csv", index=False)
            except Exception as e:
                if self.logger:
                    self.logger.log(
                        "error", f"Unable to produce KOMP stype summary: {e}"
                    )
                if self.logger:
                    self.logger.log("error", traceback.format_exc())

            primary_df.to_excel(writer, "vertical", index=False)

            if all(
                self.animal_data.shape[0] > 0,
                self.timepoint_data.shape[0] > 0,
                self.derived_data.shape[0] > 0,
                self.model.shape[0] > 0,
            ):

                secondary_df.to_excel(writer, "horizontal", index=False)
                tertiery_df.to_excel(writer, "split", index=False)
                graphs_df.to_excel(writer, "graphs", index=False)
                worksheet = writer.sheets["graphs"]

                #% run stats
                # get list of independent factors
                ind_vars = list(self.model["factors"].values)
                iv_dict = {}
                iv_dict_rev = {}

                # prepare stats dataframe to be used for easy export
                stats_df = pandas.DataFrame()
                pairwise_df = pandas.DataFrame()

                # prepare key for independent factors
                # - column names #note reserved format style
                for k in range(len(ind_vars)):
                    iv_dict[ind_vars[k]] = "__F{}__".format(k)
                    iv_dict_rev["__F{}__".format(k)] = ind_vars[k]

                # iterate through outcome measure columns and clean data for ANOVA
                counter = 0
                for c in ColumnStyles.values():
                    temp_df = primary_df[
                        [c] + list(self.model["factors"].values)
                    ]
                    temp_df.loc[:, c] = pandas.to_numeric(
                        temp_df[c], errors="coerce"
                    )
                    temp_df = temp_df.dropna()

                    temp_df["om"] = temp_df[c]
                    temp_df["gp"] = ""

                    for k in iv_dict_rev:
                        temp_df[k] = temp_df[iv_dict_rev[k]]
                        temp_df["gp"] += temp_df[k].astype(str)

                    homosced = pingouin.homoscedasticity(
                        temp_df, dv="om", group="gp"
                    )["pval"].values[0]
                    try:
                        normal = pingouin.normality(
                            temp_df, dv="om", group="gp"
                        )["pval"].values[0]
                    except:
                        normal = "unable to test"
                    table = temp_df.anova(
                        "om", between=list(iv_dict_rev.keys()), ss_type=3
                    )
                    table["levene pval"] = str(homosced)
                    table["shapiro pval"] = str(normal)

                    # replace independent factor placeholders with original names
                    for k in iv_dict_rev:
                        table["Source"] = table["Source"].replace(
                            k, iv_dict_rev[k], regex=True
                        )
                    table["outcome_measure"] = c
                    stats_df = stats_df.append(table)

                    # create data frame to assist with summary plot generation
                    #   (uses pandas agg function)
                    agg_df = (
                        temp_df.groupby(ind_vars)
                        .agg([numpy.mean, len, scipy.stats.sem])
                        .reset_index()
                    )
                    agg_cols = agg_df.columns
                    agg_df.columns = [
                        i[0] if i[0] != c else i[1] for i in agg_cols
                    ]
                    agg_df["axis"] = (
                        agg_df[ind_vars].astype(str).agg("_".join, axis=1)
                    )

                    temp_plot = agg_df.plot(
                        kind="barh",
                        title=c + " [mean+/-sem]",
                        legend=True,
                        y="mean",
                        x="axis",
                    )
                    temp_plot.errorbar(
                        agg_df["mean"],
                        agg_df["axis"],
                        xerr=agg_df["sem"],
                        ecolor="black",
                        linewidth=0,
                        elinewidth=1,
                        capsize=4,
                    )
                    temp_plot.set(xlabel=c, ylabel="_".join(ind_vars))
                    temp_plot = temp_plot.get_figure()

                    temp_plot.savefig(
                        self.output_path
                        + "_"
                        + re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', "", c)
                        + ".png",
                        bbox_inches="tight",
                    )
                    worksheet.insert_image(
                        "B{}".format(2 + counter * 20),
                        self.output_path
                        + "_"
                        + re.sub(r'[\\/\:*"<>\|\.%\$\^&£]', "", c)
                        + ".png",
                    )
                    counter += 1

                    # produce pairwise comparisons
                    pairwise_list = []

                    for i in range(len(iv_dict_rev)):
                        pairwise_list += list(
                            itertools.combinations(iv_dict_rev, i + 1)
                        )

                    for i in pairwise_list:
                        if len(i) > 1:
                            temp_df["*".join(i)] = (
                                temp_df[[j for j in i]]
                                .astype(str)
                                .agg(" * ".join, axis=1)
                            )

                        pairs = list(
                            itertools.combinations(
                                temp_df["*".join(i)].unique(), 2
                            )
                        )
                        for p in pairs:

                            if (
                                len(temp_df[temp_df["*".join(i)] == p[0]]) < 2
                                or len(temp_df[temp_df["*".join(i)] == p[1]])
                                < 2
                            ):
                                pairwise_df = pairwise_df.append(
                                    pandas.DataFrame(
                                        {
                                            "outcome_measure": [c],
                                            "comparison": [
                                                " vs ".join(
                                                    [str(q) for q in p]
                                                )
                                            ],
                                            "notes": ["cannot compare"],
                                        }
                                    )
                                )
                                continue

                            temp_p_df = pingouin.ttest(
                                temp_df[temp_df["*".join(i)] == p[0]]["om"],
                                temp_df[temp_df["*".join(i)] == p[1]]["om"],
                            )
                            temp_np_df = pingouin.mwu(
                                temp_df[temp_df["*".join(i)] == p[0]]["om"],
                                temp_df[temp_df["*".join(i)] == p[1]]["om"],
                            )
                            pw_stats_df = temp_p_df
                            pw_stats_df.index = ["PAIRWISE"]
                            pw_stats_df["ttest pval"] = temp_p_df["p-val"]
                            pw_stats_df["mwu pval"] = temp_np_df[
                                "p-val"
                            ].values[0]
                            pw_stats_df.pop("p-val")
                            pw_stats_df["outcome_measure"] = c
                            pw_stats_df["comparison"] = " vs ".join(
                                [str(q) for q in (p)]
                            )
                            pairwise_df = pairwise_df.append(pw_stats_df)

                pairwise_df = pairwise_df.reset_index()
                pairwise_df = pairwise_df[
                    ["outcome_measure", "comparison"]
                    + [
                        j
                        for j in pairwise_df
                        if j not in ["outcome_measure", "comparison", "notes"]
                    ]
                    + ["notes"]
                ]

                stats_df.to_excel(writer, "stats", index=False)
                pairwise_df.to_excel(writer, "pairwise", index=False)
        except Exception as e:
            if self.logger:
                self.logger.log("error", f"unable to process data: {e}")
            if self.logger:
                self.logger.log("error", traceback.format_exc())
        try:
            writer.save()
            if self.logger:
                self.logger.log("info", f"Output Saved - {self.output_path}")

        except Exception as e:
            if self.logger:
                self.logger.log("error", f"Unable to save file: {e}")
            if self.logger:
                self.logger.log("error", traceback.format_exc())
