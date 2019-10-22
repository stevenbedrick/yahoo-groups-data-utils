"""
Convert a directory full of the .json files created by yahoo-group-archiver representing Yahoo Groups polls.

Each poll will end up as a worksheet in the output file, with responses, timestamps, etc. intact.

"""

import json
import os, glob
import logging
import argparse
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side
import html

BOTTOM_BORDER = Border(bottom=Side(border_style='thin', color="000000"))

def process_file(in_fname: str, w: Workbook) -> None:
    sheet = w.create_sheet(title=os.path.basename(in_fname)) # happens in-place
    j = json.load(open(in_fname, "rb"))
    # pull out the relevant top-level metadata:
    date_created = datetime.fromtimestamp(j['dateCreated'])
    date_ended = datetime.fromtimestamp(j['dateEnd'])
    survey_text = html.unescape(j['surveyText'])
    total_votes = j['voteCount']

    # populate the top part of the spreadsheet:
    sheet["A1"] = "Topic:"
    sheet["B1"] = survey_text

    sheet["A2"] = "Posted:"
    sheet["B2"] = date_created.isoformat()
    sheet["A3"] = "Closed:"
    sheet["B3"] = date_ended.isoformat()
    sheet["A4"] = "Total votes:"
    sheet["B4"] = total_votes

    # overview table of different responses and counts:

    sheet["A7"] = "Overall Responses:"
    sheet["A8"] = "Response"
    sheet["B8"] = "Count"
    sheet["A8"].border = BOTTOM_BORDER
    sheet["B8"].border = BOTTOM_BORDER

    row = 9
    for s in j["selections"]:
        sheet[f"A{row}"] = s["selectionText"]
        sheet[f"B{row}"] = s["selections"]
        row += 1

    # now do the full table of responses:

    row += 1 # skip a row
    sheet[f"A{row}"] = "Non-Anonymous Responses:"

    row += 1
    # header row for the responses themselves
    sheet[f"A{row}"] = "nickname"
    sheet[f"A{row}"].border = BOTTOM_BORDER

    sheet[f"B{row}"] = "email"
    sheet[f"B{row}"].border = BOTTOM_BORDER

    sheet[f"C{row}"] = "response"
    sheet[f"C{row}"].border = BOTTOM_BORDER


    row += 1

    for s in j["selections"]:
        value = s["selectionText"]
        for r in s["responses"]:
            if "nickname" in r:
                sheet[f"A{row}"] = r["nickname"]
            sheet[f"B{row}"] = r["email"]
            sheet[f"C{row}"] = value
            row += 1

if __name__ == "__main__":
    logging.basicConfig(level="INFO", format="%(levelname)s: %(message)s")

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input_directory', type=str, help="Path to directory containing JSON files.")
    parser.add_argument('output_file', type=str, help="Path to desired output file (e.g. \"polls.xlsx\")")
    args = parser.parse_args()

    infiles = glob.glob(os.path.join(args.input_directory, "*.json"))

    wb = Workbook()

    for fname in infiles:
        process_file(fname, wb)

    # clear out the first (original, blank) sheet that was created automatically
    wb.remove(wb["Sheet"])

    # save the workbook!
    wb.save(args.output_file)
