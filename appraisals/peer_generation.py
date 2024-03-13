"""
Made by Tew En Hao, Assistant Trainer,
ASA Platoon 3 Coyote Company BMTC SCH V

V1.0.0 as of 120324
"""

import helpers
import pandas as pd

WEIGHTAGE = 5    # as of 221223, peer appraisal weightage is 5%

# add list of OOC personnel here
# add as a dictionary of lists, where each OOC personnel 4D will be included as text strings in the respective lists
ooc_pers = {
    'PLT_1': ['C1105'],
    'PLT_2': [],
    'PLT_3': ['C3105', 'C3209', 'C3213', 'C3218'], 
    'PLT_4': [],
    'PLT_5': []
}

# we assume that all 5 commander result sheets are in *one singular workbook*
# sheet names in the workbook MUST FOLLOW the keys of `ooc_pers` (PLT_1, PLT_2, etc.)

dfs = helpers.excel_to_df('trial.xlsx')    # change the filename here as needed

# eventually where all our results will go
mega_df = pd.DataFrame()

# process each dataframe
for plt, df in dfs.items():
    new_df = helpers.get_peer_results(df, WEIGHTAGE, ooc_pers[plt])
    mega_df = pd.concat([mega_df, new_df], axis=1)

mega_df.sort_index(inplace=True)
mega_df.to_excel('output.xlsx', header=False)