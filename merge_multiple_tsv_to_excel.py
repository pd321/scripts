# Merge multiple given tsv files to an Excel sheet

import pandas as pd
import argparse
from pathlib import Path


def main(args):

    output = {}
    filename_order = []
    for input_tsv_file in args.input_tsv_files:
        tsv_name = Path(input_tsv_file).stem
        tsv_df = pd.read_csv(filepath_or_buffer=input_tsv_file, sep="\t")
        output[tsv_name] = tsv_df
        filename_order.append(tsv_name)

    writer = pd.ExcelWriter(args.output_xlsx_file, engine='xlsxwriter')

    for filename in filename_order:
        output[filename].to_excel(writer, sheet_name=filename, index=False)

    writer.save()


if __name__ == "__main__":

    parser = argparse.ArgumentParser(description="Merge multiple tsv files into an excel sheet")
    parser.add_argument('-i', '--input', dest='input_tsv_files', nargs='+')
    parser.add_argument('-o', '--output', dest='output_xlsx_file')
    cmd_args = parser.parse_args()
    main(cmd_args)