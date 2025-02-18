
# Script to convert CSV to IIF output.

import csv
import sys
import traceback
import typing
from datetime import datetime

import QBComTypes as qb

def error(trans):
    """Log errors to stderr with traceback."""
    sys.stderr.write(f"{trans}\n")
    traceback.print_exc(file=sys.stderr)

def loadListsFromQB(sessionManager):
    """Load lists from QuickBooks."""

    requestMsgSet: qb.IMsgSetRequest = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    accountQueryRq = qb.IAccountQuery(requestMsgSet.AppendAccountQueryRq())
    accountQueryRq.ORAccountListQuery.AccountListFilter.ActiveStatus.SetValue(qb.constants.asActiveOnly)

    vendorQueryRq = qb.IVendorQuery(requestMsgSet.AppendVendorQueryRq())
    vendorQueryRq.ORVendorListQuery.VendorListFilter.ActiveStatus.SetValue(qb.constants.asActiveOnly)

    responseMsgSet: qb.IMsgSetResponse = sessionManager.DoRequests(requestMsgSet)

    acctList = qb.IAccountRetList(responseMsgSet.ResponseList.GetAt(0).Detail)
    vendorList = qb.IVendorRetList(responseMsgSet.ResponseList.GetAt(1).Detail)

    validAccounts = [acct.FullName.GetValue() for acct in acctList]
    validVendors = [vendor.Name.GetValue() for vendor in vendorList]

    return validAccounts, validVendors

def preCheck(sessionManager, transactions, vendorName="Merchant Name"):
    """Pre-check the CSV file for valid accounts and vendors."""
    validAccounts, validVendors = loadListsFromQB(sessionManager)

    good = True

    vendors = set(t[vendorName] for t in transactions)
    accounts = set(t["GL Code ID"] for t in transactions)

    badVendors = vendors - set(validVendors)
    badAccounts = accounts - set(validAccounts)

    for vendor in badVendors:
        error(f"Invalid {vendorName}: {vendor}")
        good = False

    for account in badAccounts:
        error(f"Invalid account name {account}")
        good = False

    return good

def endSession(sessionManager):
    """End the QuickBooks session."""
    sessionManager.EndSession()
    sessionManager.CloseConnection()

def walkRs(respMsgSet: qb.IMsgSetResponse):
    """Walk the response message set."""
    if respMsgSet.ResponseList is None:
        return

    respList = qb.IResponseList(respMsgSet.ResponseList)
    if respList is None:
        return

    for resp in respList:
        if resp.StatusCode >= 0 and resp.Detail is not None:
            respType = typing.cast(int, resp.Type.GetValue())
            if respType == qb.ENResponseType.rtDepositAddRs:
                depositRet: qb.IDepositRet = qb.IDepositRet(resp.Detail)
                walkDepositRet(depositRet)
            elif respType == qb.ENResponseType.rtCheckAddRs:
                checkRet: qb.ICheckRet = qb.ICheckRet(resp.Detail)
                walkCheckRet(checkRet)
            else:
                error(f"Unknown response type {qb.ENResponseType(respType).name}")

def walkDepositRet(depositRet: qb.IDepositRet):
    """Walk the deposit return."""
    if depositRet is None:
        return

    # Get value of TxnDate
    txnDate = depositRet.TxnDate.GetValue()
    txnToAccount = depositRet.DepositToAccountRef.FullName.GetValue()
    txnMemo = depositRet.Memo.GetValue()
    txnTotal = depositRet.DepositTotal.GetValue()

    if depositRet.DepositLineRetList is not None:
        depositLineRetList = qb.IDepositLineRetList(depositRet.DepositLineRetList)
        for depositLineRet in depositLineRetList:
            lineAccount = ""
            if depositLineRet.AccountRef is not None:
                accountRef = depositLineRet.AccountRef
                lineAccount = accountRef.FullName.GetValue()
            lineMemo = depositLineRet.Memo.GetValue()
            lineAmount = depositLineRet.Amount.GetValue()
            error(f"Deposit {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} {lineMemo} {lineAmount}")

def walkCheckRet(checkRet: qb.ICheckRet):
    """Walk the check return."""
    if checkRet is None:
        return

    # Get value of TxnDate
    txnDate = checkRet.TxnDate.GetValue()
    txnToAccount = checkRet.DepositToAccountRef.FullName.GetValue()
    txnMemo = checkRet.Memo.GetValue()
    txnTotal = checkRet.Amount.GetValue()

    if checkRet.ExpenseLineRetList is not None:
        expenseLineRetList = qb.IExpenseLineRetList(checkRet.ExpenseLineRetList)
        for expenseLineRet in expenseLineRetList:
            lineAccount = ""
            if expenseLineRet.AccountRef is not None:
                accountRef = expenseLineRet.AccountRef
                lineAccount = accountRef.FullName.GetValue()
            lineMemo = expenseLineRet.Memo.GetValue()
            lineAmount = expenseLineRet.Amount.GetValue()
            error(f"Cheque {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} {lineMemo} {lineAmount}")


def main(inputFileName, iifFileName):

    count = 0

    try:
        # open the files
        with open(inputFileName, 'r', newline='', encoding='utf-8') as inputFile:

            # the first line is a header, so we can use a DictReader
            csvReader = csv.DictReader(inputFile)

            # load the full set of transactions into memory
            transactions = list(csvReader)

            # create a QuickBooks session
            with qb.IQBSessionManager() as sessionManager:
                if csvReader.fieldnames and "Report Name" in csvReader.fieldnames:
                    # this must be a reimbursement file
                    reimbursement = True
                    payeeNameField= "Requester"
                else:
                    # this must be a standard transactions file
                    reimbursement = False
                    payeeNameField = "Merchant Name"

                if not preCheck(sessionManager, transactions, payeeNameField):
                    return

                requestMsgSet = qb.IMsgSetRequest(sessionManager.CreateMsgSetRequest("CA", 16, 0))
                requestMsgSet.Attributes.OnError = qb.constants.roeContinue

                for trans in transactions:
                    if reimbursement:
                        trnsDate = datetime.strptime(trans["Expense Date"], '%d/%m/%y')
                    else:
                        trnsDate = datetime.strptime(trans["Expense Date"], '%y-%m-%d')

                    trnsDesc = trans["Description"].strip()
                    trnsMerch = trans[payeeNameField]
                    trnsGlcode = trans["GL Code ID"]

                    try:
                        trnsAmount = -1 * float(trans.get("Total", 0))
                        trnsSubtotal = float(trans.get("Subtotal", 0))
                        trnsTax = float(trans.get("Tax", 0))
                    except ValueError:
                        error("Invalid number format in transaction.")
                        continue

                    if trnsAmount < 0 and trnsGlcode == "Other Income:Interest Income" and not reimbursement:
                        # trnsType = "DEPOSIT"
                        depAddRq = qb.IDepositAdd(requestMsgSet.AppendDepositAddRq())
                        depAddRq.DepositToAccount.FullName.SetValue("Float Financial")
                        depAddRq.TxnDate.SetValue(trnsDate)
                        depAddRq.Memo.SetValue(trnsDesc)
                        depLineAddRq: qb.IDepositLineAdd = depAddRq.DepositLineAddList.Append()
                        depLineAddRq.ORDepositLineAdd.DepositInfo.AccountRef.FullName.SetValue(trnsGlcode)
                        depLineAddRq.ORDepositLineAdd.DepositInfo.Amount.SetValue(trnsAmount)

                    else:
                        # trnsType = "CHEQUE"
                        chkAddRq = qb.ICheckAdd(requestMsgSet.AppendCheckAddRq())
                        chkAddRq.AccountRef.FullName.SetValue("Float Financial")
                        chkAddRq.IsToBePrinted.SetValue(False)
                        chkAddRq.TxnDate.SetValue(trnsDate)
                        chkAddRq.PayeeEntityRef.FullName.SetValue(trnsMerch)
                        expAdd: qb.IExpenseLineAdd = chkAddRq.ExpenseLineAddList.Append()
                        expAdd.AccountRef.FullName.SetValue(trnsGlcode)
                        expAdd.Amount.SetValue(trnsSubtotal)
                        expAdd.Memo.SetValue(trnsDesc)
                        if trnsTax != 0:
                            expAddT: qb.IExpenseLineAdd = chkAddRq.ExpenseLineAddList.Append()
                            expAddT.AccountRef.FullName.SetValue("GST Accounts Receivable")
                            expAdd.Amount.SetValue(trnsTax)
                            expAddT.Memo.SetValue("Half of the GST")

                    count += 1

                respMsgSet = qb.IMsgSetResponse(sessionManager.DoRequests(requestMsgSet))

            walkRs(respMsgSet)

        print(f"Conversion complete, {count} transactions in {iifFileName}")

    except Exception as e:
        error(f"Failed to process {inputFileName}: {e}")


if __name__ == '__main__':

    if len(sys.argv) != 2:
        print("usage:   Float2QB input.csv")
        # read the input filename from the console
        inputFileName = input("Enter the name of the input file: ")
        main(inputFileName, "")
    else:
        main(sys.argv[1], "")

    # wait for keypress before closing the console window
    input("Press Enter to close this window")




