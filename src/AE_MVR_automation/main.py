# main.py
# The main entry point for the slide update automation script

import pandas as pd
import pptx
import logging
from config import DATASET_FILE_PATH, PPTX_INPUT_FILE, PPTX_OUTPUT_FILE
from helper_modules import read_data
from slide_updaters import (slide_30_updater,)
from config.logging_config import setup_logging


# TODO: Add code to update slide data automatically

logger = logging.getLogger(__name__)

def main():
    setup_logging()
    # logger = logging.getLogger(__name__)
    logger.info('Starting AE MVR - Automation')

    print(f'{DATASET_FILE_PATH = }')
    df, meta, df_labeled = read_data(DATASET_FILE_PATH)
    # print(df)

    prs = pptx.Presentation(PPTX_INPUT_FILE)

    # slide_1_updater(df, prs)
    # slide_2_updater(df, df_labeled, prs)
    # # slide_3_updater(df, prs)
    # slide_4_updater(meta, df, df_labeled, prs)
    # slide_5_updater(meta, df, df_labeled, prs)
    # slide_6_updater(meta, df, df_labeled, prs)
    # slide_7_updater(meta, df, df_labeled, prs)
    # slide_8_upper_chart_updater(meta, df_labeled, prs)
    # slide_8_lower_chart_updater(meta, df_labeled, prs)
    # slide_9_updater(meta, df_labeled, prs)
    # slide_10_updater(meta, df, df_labeled, prs)
    # # slide_11_updater(df, prs)
    # slide_18_updater(meta, df, df_labeled, prs)
    # slide_24_updater(meta, df, df_labeled, prs)
    # slide_25_updater(meta, df, df_labeled, prs)
    # slide_27_updater(meta, df, df_labeled, prs)
    # slide_29_updater(meta, df, df_labeled, prs)
    slide_30_updater(meta, df, df_labeled, prs)
    # slide_31_updater(df, prs)
    # slide_32_updater(df, prs)
    # slide_33_updater(df, prs)
    print('back in main.py')



    prs.save(PPTX_OUTPUT_FILE)
    logger.info('Finished AE 311 Additional Analysis - Automation')



if __name__ == "__main__":
    main()