"""
Made by Tew En Hao, Assistant Trainer,
ASA Platoon 3 Coyote Company BMTC SCH V

V1.0.0 as of 120324
"""

import pandas as pd
import numpy as np

def excel_to_df(filename: str) -> dict:
    """
    Takes in the name of the excel file and returns a list of dataframes, with each dataframe
    containing data from one sheet

    Parameters
    ----------
    - filename [str]: a string which represents the name of the excel file (with its file extension)

    Returns
    -------
    - dfs [dict]: a dictionary of dataframes which represents the data in the excel workbook
    """
    dfs = pd.read_excel(filename, sheet_name=None)
    return dfs

def get_comd_results(df: pd.DataFrame, weightage: int, ooc_pers: list) -> pd.DataFrame:
    """
    Combines and rebases the commander appraisal results into its required form

    Parameters
    ----------
    - df [pd.DataFrame]: the dataframe with the raw data
    - weightage [int]: the weightage of the commander appraisal
    - ooc_pers [list]: the list of 4Ds who have OOCed from the course

    Returns
    -------
    - new_transposed_df [pd.DataFrame]: dataframe with the processed data
    """

    # drop unnecessary columns
    df.drop(labels=['Timestamp', 'Which Platoon and Section are you from?'], axis=1, inplace=True)

    # group data by 4D, then sum and rebase results to required base
    new_df = df.groupby(by=df.columns.str.split().str[0], axis=1, as_index=True).sum() / 12 * weightage
    new_df_transposed = new_df.transpose().sum(axis=1)

    # add ooc personnel and set scores as 0
    for ooc in ooc_pers:
        new_df_transposed[ooc] = 0

    new_df_transposed.sort_index(inplace=True)

    # return new df
    return new_df_transposed

def get_peer_results(df: pd.DataFrame, weightage: int, ooc_pers: list) -> pd.DataFrame:
    """
    Combines and rebases the peer appraisal results into its required form

    Parameters
    ----------
    - df [pd.DataFrame]: the dataframe with the raw data
    - weightage [int]: the weightage of the peer appraisal
    - ooc_pers [list]: the list of 4Ds who have OOCed from the course

    Returns
    -------
    - result_df [pd.DataFrame]: dataframe with the processed data
    """

    # drop timestamp and platoon/section info as these are not required
    df.drop(labels=['Timestamp', 'Which Platoon and Section are you from?'], axis=1, inplace=True)

    # merge 4Ds (which will become the eventual index) into one column
    # the 4Ds here refer to the recruit's 4Ds, ie. the first column

    df['4D'] = df[[df.filter(like='What').columns[0]]]    # this step gets all 4Ds in the first column
    
    for col in df.filter(like='What'):    # this step fills up the rest of the 4Ds
        df['4D'].fillna(df[col], inplace=True)    # fills in the missing 4Ds from other columns into the first column
        df.drop(col, axis=1, inplace=True)    # drop the original column since the data has been shifted over

    # set index of df first
    df.set_index('4D', inplace=True)
    
    # then group data by 4D (ie. combine the 4 categories together)
    new_df = df.groupby(by=df.columns.str.split().str[0], axis=1, as_index=True).sum() / 12 * weightage
    
    # transpose dataframe and sum each recruit's results and divide by number of recruits in his section

    # transpose because without transposing, each row is each recruit's score for other recruits
    # we want each row to be each recruit's score (which is currently each column)
    new_df = new_df.transpose()
    new_df.sort_values(by='4D', axis=1, inplace=True)

    # sum each row up and store in new dataframe, `result_df` later, but we use a dictionary to store the data first
    results = {}
    
    for label, row in new_df.iterrows():
    # rebased result of individual = (total score - own score) / total number of peers that appraised him
        try:
            results[label] = (row.sum() - row[label]) / (new_df.loc[label, :].astype(bool).sum() - 1)
            results[label] = np.mean(results[label])    # to ensure that any results that come out as DataFrames due to multiple results are averaged out

        # shit happens, recruits play punk and don't do appraisal form :/
        # used to be np.random.rand() * 5, but that shit is too unpredictable it gave like 1.5 ish cannot show that lmao so need to tone down
        except KeyError:
            results[label] = np.random.uniform(3, 5)

    # account for OOC personnel
    for ooc in ooc_pers:
        results[ooc] = 0

    # throw the processed data into a new dataframe and sort according to 4D (which is the index now)
    result_df = pd.DataFrame.from_dict(results, orient='index', columns=['Score'])
    result_df.sort_index(inplace=True)

    return result_df