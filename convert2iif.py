
# Script to convert CSV to IIF output.

from os import path
import sys, traceback
import csv

def error(trans):
    """Log errors to stderr with traceback."""
    sys.stderr.write(f"{trans}\n")
    traceback.print_exc(file=sys.stderr)

def loadListsFromQB():
    """Load lists from QuickBooks."""
    from win32com import client
    import QBComTypes as qb

    sessionManager: qb.IQBSessionManager = client.Dispatch("QBFC16.QBSessionManager")
    sessionManager.OpenConnection("", "Test App")
    sessionManager.BeginSession("", qb.constants.omDontCare)
    requestMsgSet: qb.IMsgSetRequest = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    accountQueryRq: qb.IAccountQuery = requestMsgSet.AppendAccountQueryRq()
    accountQueryRq.ORAccountListQuery.AccountListFilter.ActiveStatus.SetValue(qb.constants.asActiveOnly)

    vendorQueryRq: qb.IVendorQuery = requestMsgSet.AppendVendorQueryRq()
    vendorQueryRq.ORVendorListQuery.VendorListFilter.ActiveStatus.SetValue(qb.constants.asActiveOnly)

    responseMsgSet: qb.IMsgSetResponse = sessionManager.DoRequests(requestMsgSet)
    sessionManager.EndSession()
    sessionManager.CloseConnection()

    acctList = qb.IAccountRetList(responseMsgSet.ResponseList.GetAt(0).Detail)
    vendorList = qb.IVendorRetList(responseMsgSet.ResponseList.GetAt(1).Detail)

    validAccounts = [acct.FullName.GetValue() for acct in acctList]
    validVendors = [vendor.Name.GetValue() for vendor in vendorList]

    return validAccounts, validVendors

def preCheck(transactions):
    """Pre-check the CSV file for valid accounts and vendors."""
    validAccounts, validVendors = loadListsFromQB(validAccounts, validVendors)

    good = True

    vendors = set(t["Merchant Name"] for t in transactions)
    accounts = set(t["Glcode Name"] for t in transactions)

    badVendors = vendors - set(validVendors)
    badAccounts = accounts - set(validAccounts)

    for vendor in badVendors:
        error(f"Invalid vendor name {vendor}")
        good = False
    
    for account in badAccounts:
        error(f"Invalid account name {account}")
        good = False    
    
    return good


def main(inputFileName, iifFileName):

    # This is the IIF template

    head = "!TRNS\tTRNSID\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\tTOPRINT\n"\
        + "!SPL\tSPLID\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO\tCLEAR\n"\
        + "!ENDTRNS\n"

    tail = "ENDTRNS\n"

    if iifFileName == "":
        iifFileName = path.splitext(inputFileName)[0] + ".iif"

    count = 0

    try:
        # open the files
        with open(inputFileName, 'r', newline='', encoding='utf-8') as inputFile, \
             open(iifFileName, 'w', newline='', encoding='utf-8') as outputFile:

            # the first line is a header, so we can use a DictReader
            csvReader = csv.DictReader(inputFile)
            
            # load the full set of transactions into memory
            transactions = list(csvReader)

            if not preCheck(transactions):
                return

            outputFile.write(head)
            for trans in transactions:

                trnsDate = trans["Raw DateTime"][5:7] + "/" +trans["Raw DateTime"][8:10] + "/" + trans["Raw DateTime"][2:4]

                trnsDesc = trans["Description"].strip()
                trnsMerch = trans["Merchant Name"]
                trnsGlcode = trans["Glcode Name"]

                try:
                    trnsAmount = -1 * float(trans.get("Total Amount", 0))
                    trnsSubtotal = float(trans.get("Subtotal Amount", 0))
                    trnsTax = float(trans.get("Tax Amount", 0))
                except ValueError:
                    error("Invalid number format in transaction.")
                    continue

                if trnsAmount > 0:
                    trnsType = "DEPOSIT"
                else:
                    trnsType = "CHEQUE"

                trnsTemplate = f"TRNS\t\t{trnsType}\t{trnsDate}\tFloat Financial\t{trnsMerch}\t\t{trnsAmount}\t\t{trnsDesc}\tY\tN\n"
                splTemplate = f"SPL\t\t{trnsType}\t{trnsDate}\t{trnsGlcode}\t\t\t{trnsSubtotal}\t\t{trnsDesc}\tY\n"
                taxTemplate = f"SPL\t\t{trnsType}\t{trnsDate}\tGST Accounts Receivable\t\t\t{trnsTax}\t\tHalf of the GST\tY\n"

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

    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("usage:   Convert2IIF input.csv output.iif")    
    elif len(sys.argv) == 2:
        main(sys.argv[1], "")
    else:
        main(sys.argv[1], sys.argv[2])
