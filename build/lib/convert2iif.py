
# Script to convert CSV to IIF output.

import os
import sys, traceback, re
import csv

PROJECT_ROOT = os.path.dirname(os.path.realpath(__file__))


def error(trans):
    """Log errors to stderr with traceback."""
    sys.stderr.write(f"{trans}\n")
    traceback.print_exc(file=sys.stderr)


def main(inputFileName):
    # This is the IIF template

    head = "!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\n"\
        + "!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\n"\
        + "!ENDTRNS\n"

    tail = "ENDTRNS\n"

    iifFileName = os.path.join(PROJECT_ROOT, f"{inputFileName}.iif")

    count = 0

    try:
        # open the files
        with open(inputFileName, 'r', newline='', encoding='utf-8') as inputFile, \
             open(iifFileName, 'w', newline='', encoding='utf-8') as outputFile:

            outputFile.write(head)

            # the first line is a header, so we can use a DictReader
            csvReader = csv.DictReader(inputFile)

            for trans in csvReader:

                trnsDate = trans["Raw DateTime"][5:7] + "/" +trans["Raw DateTime"][8:10] + "/" + trans["Raw DateTime"][2:4]

                trnsDesc = trans["Description"].strip()
                trnsMerch = trans["Merchant Name"]
                trnsGlcode = trans["Glcode Name"]

                try:
                    trnsAmount = float(trans.get("Total Amount", 0))
                    trnsSubtotal = float(trans.get("Subtotal Amount", 0))
                    trnsTax = float(trans.get("Tax Amount", 0))
                except ValueError:
                    error("Invalid number format in transaction.")
                    continue

                trnsTemplate = f"TRNS\t\tCREDIT CARD\t{trnsDate}\tFloat\t{trnsMerch}\t\t{trnsAmount}\t\t{trnsDesc}\n"
                splTemplate = f"SPL\t\tCREDIT CARD\t{trnsDate}\t{trnsGlcode}\t{trnsMerch}\t\t{trnsSubtotal}\t\t\t\n"
                taxTemplate = f"SPL\t\tCREDIT CARD\t{trnsDate}\tGST Accounts Receivable\t\t\t{trnsTax}\t\tHalf of the GST\n"

                outputFile.write(trnsTemplate)
                outputFile.write(splTemplate)

                if trnsTax != 0:
                    outputFile.write(taxTemplate)
                
                outputFile.write(tail)
                count += 1

        print(f"Conversion complete, {count} transactions in {iifFileName}")

    except Exception as e:
        error(f"Failed to process {inputFileName}: {e}")


if __name__ == '__main__':

    if len(sys.argv) != 2:
        print("usage:   python convert2iif.py input.csv")

    main(sys.argv[1])
