# slide_30.py
# Update slide 6 with new quarter data

import logging
import pandas as pd
from pptx.chart.data import CategoryChartData
# from AE_311_AA_automation.config import REPORTING_PERIOD, REPORTING_YEAR


logger = logging.getLogger(__name__)


# TODO break out question concat routine as a module

def slide_30_updater(meta, df, df_labeled, prs):
    print("\n\nSlide 30 Updater\n")
    slide_index = 29
    logger.info(f'Slide {slide_index + 1} updated')
    slide = prs.slides[slide_index]

    # question_column = 'Q11'

    question_list = [
        'Q26_1',
        'Q26_2',
        'Q26_3',
        'Q26_4',
        'Q26_5',
        'Q26_6',
        'Q26_7',
        'Q26_8',
        'Q26_9',
        'Q26_10',
        'Q26_10_other',
    ]
    for question in question_list:
        dept_counts = None
        dept_name = meta.column_names_to_labels[question].split('? ', 1)[1]
        result = df_labeled[question].value_counts().get('Checked', 0)
        if result != 0:
            print(f'{dept_name}\t{result}')


        # result_df = pd.crosstab(df_labeled[question_column], df_labeled[question], normalize='columns')
        # if 'Checked' in result_df.columns:
        #     dept_counts = pd.crosstab(df_labeled[question_column], df_labeled[question])['Checked'].sum()
        #     result_series =result_df.get('Checked', 0.0)
        # else:
        #     continue
        # if 'Yes' in result_series.index:
        #     result_df_yes = result_series.loc['Yes']
        # else:
        #     continue
        # print(f'{dept_name} (n={dept_counts})\t{result_df_yes:.0%}\tDepartment counts = \t{dept_counts}')

    logger.info(f'Slide {slide_index + 1} updated')
