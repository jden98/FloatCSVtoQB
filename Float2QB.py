# Script to convert CSV to IIF output.

import csv
import sys
import traceback
from datetime import datetime
from pathlib import Path

import QBComTypes as qb

def DeQuote(s: str) -> str:
    """Remove quotes from a string."""
    if s.startswith('"') and s.endswith('"'):
        return s[1:-1]
    return s

def Error(trans):
    """Log errors to stderr with traceback."""
    sys.stderr.write(f"{trans}\n")
    traceback.print_exc(file=sys.stderr)


def LoadListsFromQB(
    sessionManager: qb.IQBSessionManager,
) -> tuple[list[str], list[str]]:
    """Load lists from QuickBooks."""

    requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    accountQueryRq = requestMsgSet.AppendAccountQueryRq()
    accountQueryRq.ORAccountListQuery.AccountListFilter.ActiveStatus.SetValue(
        qb.constants.asActiveOnly
    )

    vendorQueryRq = requestMsgSet.AppendVendorQueryRq()
    vendorQueryRq.ORVendorListQuery.VendorListFilter.ActiveStatus.SetValue(
        qb.constants.asActiveOnly
    )

    responseMsgSet = sessionManager.DoRequests(requestMsgSet)

    acctList = qb.IAccountRetList(responseMsgSet.ResponseList.GetAt(0).Detail)
    vendorList = qb.IVendorRetList(responseMsgSet.ResponseList.GetAt(1).Detail)

    validAccounts = [acct.FullName.GetValue() for acct in acctList]
    validVendors = [vendor.Name.GetValue() for vendor in vendorList]

    return validAccounts, validVendors


def PreCheck(
    sessionManager: qb.IQBSessionManager,
    transactions: list[dict],
    vendorName="Merchant Name",
) -> bool:
    """Pre-check the CSV file for valid accounts and vendors."""
    validAccounts, validVendors = LoadListsFromQB(sessionManager)

    good = True

    vendors = set(t[vendorName] for t in transactions)
    accounts = set(t["GL Code ID"] for t in transactions)

    badVendors = vendors - set(validVendors)
    badAccounts = accounts - set(validAccounts)

    for vendor in badVendors:
        Error(f"Invalid {vendorName}: {vendor}")
        good = False

    for account in badAccounts:
        Error(f"Invalid account name {account}")
        good = False

    return good


def WalkRs(respMsgSet: qb.IMsgSetResponse) -> None:
    """Walk the response message set."""
    if respMsgSet.responseList is None:
        return

    for resp in respMsgSet.responseList:
        if resp.StatusCode >0:
            Error(f"Error: Code:{resp.StatusCode} Severity: {resp.StatusSeverity} Message: {resp.StatusMessage}")
        if resp.StatusCode >= 0 and resp.Detail is not None:
            respType = int(resp.Type.GetValue())
            if respType == qb.ENResponseType.rtDepositAddRs:
                depositRet: qb.IDepositRet = qb.IDepositRet(resp.Detail)
                WalkDepositRet(depositRet, resp.StatusCode, resp.StatusSeverity, resp.StatusMessage)
            elif respType == qb.ENResponseType.rtCheckAddRs:
                checkRet: qb.ICheckRet = qb.ICheckRet(resp.Detail)
                WalkCheckRet(checkRet, resp.StatusCode, resp.StatusSeverity, resp.StatusMessage)
            elif respType == qb.ENResponseType.rtBillAddRs:
                billRet: qb.IBillRet = qb.IBillRet(resp.Detail)
                WalkBillRet(billRet, resp.StatusCode, resp.StatusSeverity, resp.StatusMessage)
            else:
                Error(f"Unknown response type {qb.ENResponseType(respType).name}")

def WalkBillRet(billRet: qb.IBillRet, statusCode: int, statusSeverity: str, statusMessage: str) -> None:
    """Walk the bill return."""
    if billRet is None:
        return

    # Get value of TxnDate
    txnDate = billRet.TxnDate.GetValue()
    txnToAccount = billRet.VendorRef.FullName.GetValue()
    txnMemo = billRet.Memo.GetValue()
    txnTotal = billRet.AmountDue.GetValue()

    if statusCode == 0:
        print(f"Created bill from {txnToAccount} for {txnTotal}")
    else:
        if billRet.ExpenseLineRetList is not None:
            expenseLineRetList = qb.IExpenseLineRetList(billRet.ExpenseLineRetList)
            for expenseLineRet in expenseLineRetList:
                lineAccount = ""
                if expenseLineRet.AccountRef is not None:
                    accountRef = expenseLineRet.AccountRef
                    lineAccount = accountRef.FullName.GetValue()
                lineMemo = expenseLineRet.Memo.GetValue()
                lineAmount = expenseLineRet.Amount.GetValue()
                Error(
                    f"Error creating Bill {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} "
                    f"{lineMemo} {lineAmount}"
                )


def WalkDepositRet(depositRet: qb.IDepositRet, statusCode: int, statusSeverity: str, statusMessage: str) -> None:
    """Walk the deposit return."""
    if depositRet is None:
        return

    # Get value of TxnDate
    txnDate = depositRet.TxnDate.GetValue()
    txnToAccount = depositRet.DepositToAccountRef.FullName.GetValue()
    txnMemo = depositRet.Memo.GetValue()
    txnTotal = depositRet.DepositTotal.GetValue()

    if statusCode == 0:
        print(f"Created deposit to {txnToAccount} for {txnTotal}")
    else:
        if depositRet.depositLineRetList is not None:
            for depositLineRet in depositRet.depositLineRetList:
                lineAccount = ""
                if depositLineRet.AccountRef is not None:
                    accountRef = depositLineRet.AccountRef
                    lineAccount = accountRef.FullName.GetValue()
                lineMemo = depositLineRet.Memo.GetValue()
                lineAmount = depositLineRet.Amount.GetValue()
                Error(
                    f"Error creating Deposit {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} "
                    f"{lineMemo} {lineAmount}"
                )


def WalkCheckRet(checkRet: qb.ICheckRet, statusCode: int, statusSeverity: str, statusMessage: str) -> None:
    """Walk the check return."""
    if checkRet is None:
        return

    # Get value of TxnDate
    txnDate = checkRet.TxnDate.GetValue()
    txnToAccount = checkRet.AccountRef.FullName.GetValue()
    txnMemo = checkRet.Memo.GetValue()
    txnTotal = checkRet.Amount.GetValue()
    txnRefNumber = checkRet.RefNumber.GetValue()
    txnPayee = checkRet.PayeeEntityRef.FullName.GetValue()

    if statusCode == 0:
        print(f"Created cheque Number {txnRefNumber} to {txnPayee} for {txnTotal}")
    else:
        if checkRet.ExpenseLineRetList is not None:
            expenseLineRetList = qb.IExpenseLineRetList(checkRet.ExpenseLineRetList)
            for expenseLineRet in expenseLineRetList:
                lineAccount = ""
                if expenseLineRet.AccountRef is not None:
                    accountRef = expenseLineRet.AccountRef
                    lineAccount = accountRef.FullName.GetValue()
                lineMemo = expenseLineRet.Memo.GetValue()
                lineAmount = expenseLineRet.Amount.GetValue()
                Error(
                    f"Error creating Cheque {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} "
                    f"{lineMemo} {lineAmount}"
                )


def ProcessTransactions(
    sessionManager: qb.IQBSessionManager,
    transactions: list[dict],
    payeeNameField: str
) -> tuple[int, qb.IMsgSetResponse]:
    """Process the transaction data."""
    requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    count = 0
    for trans in transactions:
        trnsDate = datetime.strptime(trans["Transaction DateTime"], "%Y-%m-%d %H:%M:%S.%f%z")

        trnsDesc = trans["Description"].strip()
        trnsMerch = trans[payeeNameField]
        trnsGlcode = trans["GL Code ID"]

        try:
            trnsAmount = float(trans.get("Total", 0))
            trnsSubtotal = float(trans.get("Subtotal", 0))
            trnsTax = float(trans.get("Tax", 0))
        except ValueError:
            Error("Invalid number format in transaction.")
            continue

        if trnsAmount < 0:
            # trnsType = "DEPOSIT"
            depAddRq = qb.IDepositAdd(requestMsgSet.AppendDepositAddRq())
            depAddRq.DepositToAccountRef.FullName.SetValue("Float Financial")
            depAddRq.TxnDate.SetValue(trnsDate)
            depAddRq.Memo.SetValue(trnsDesc)
            depLineAddRq: qb.IDepositLineAdd = depAddRq.DepositLineAddList.Append()
            depositInfo = depLineAddRq.ORDepositLineAdd.DepositInfo
            depositInfo.AccountRef.FullName.SetValue(trnsGlcode)
            depositInfo.Amount.SetValue(-1 * trnsAmount)

        else:
            # trnsType = "CHEQUE"
            chkAddRq = qb.ICheckAdd(requestMsgSet.AppendCheckAddRq())
            chkAddRq.AccountRef.FullName.SetValue("Float Financial")
            chkAddRq.IsToBePrinted.SetValue(False)
            chkAddRq.TxnDate.SetValue(trnsDate)
            chkAddRq.PayeeEntityRef.FullName.SetValue(trnsMerch)
            chkAddRq.Memo.SetValue(trnsDesc)

            expAdd: qb.IExpenseLineAdd = chkAddRq.ExpenseLineAddList.Append()
            expAdd.AccountRef.FullName.SetValue(trnsGlcode)
            expAdd.Amount.SetValue(trnsSubtotal)
            expAdd.Memo.SetValue(trnsDesc)
            if trnsTax != 0:
                expAddT: qb.IExpenseLineAdd = chkAddRq.ExpenseLineAddList.Append()
                expAddT.AccountRef.FullName.SetValue("GST Accounts Receivable")
                expAddT.Amount.SetValue(trnsTax)
                expAddT.Memo.SetValue("Half of the GST")

        count += 1

    respMsgSet = qb.IMsgSetResponse(sessionManager.DoRequests(requestMsgSet))

    return count, respMsgSet

def ProcessReimbursements(
    sessionManager: qb.IQBSessionManager,
    transactions: list[dict],
    payeeNameField: str
) -> tuple[int, qb.IMsgSetResponse]:
    """Process the transaction data."""
    requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    count = 0
    for trans in transactions:
        trnsDate = datetime.strptime(trans["Expense Date"], "%d/%m/%Y")

        trnsDesc = trans["Description"].strip()
        trnsMerch = trans[payeeNameField]
        trnsGlcode = trans["GL Code ID"]

        try:
            trnsAmount = float(trans.get("Total", 0))
            trnsSubtotal = float(trans.get("Subtotal", 0))
            trnsTax = float(trans.get("Tax", 0))
        except ValueError:
            Error("Invalid number format in transaction.")
            continue

        # trnsType = "BILL"
        billAddRq = qb.IBillAdd(requestMsgSet.AppendBillAddRq())
        billAddRq.APAccountRef.FullName.SetValue("Accounts Payable")
        billAddRq.TxnDate.SetValue(trnsDate)
        billAddRq.VendorRef.FullName.SetValue(trnsMerch)
        billAddRq.Memo.SetValue(trnsDesc)

        expAdd: qb.IExpenseLineAdd = billAddRq.ExpenseLineAddList.Append()
        expAdd.AccountRef.FullName.SetValue(trnsGlcode)
        expAdd.Amount.SetValue(trnsSubtotal)
        expAdd.Memo.SetValue(trnsDesc)
        if trnsTax != 0:
            expAddT: qb.IExpenseLineAdd = billAddRq.ExpenseLineAddList.Append()
            expAddT.AccountRef.FullName.SetValue("GST Accounts Receivable")
            expAddT.Amount.SetValue(trnsTax)
            expAddT.Memo.SetValue("Half of the GST")

        count += 1

    respMsgSet = qb.IMsgSetResponse(sessionManager.DoRequests(requestMsgSet))

    return count, respMsgSet


def main(inputFileName, iifFileName):
    count = 0
    inputFilePath = Path(inputFileName)

    try:
        # open the files
        with inputFilePath.open("r", newline="", encoding="utf-8") as inputFile:
            csvReader = csv.DictReader(inputFile)

            # load the full set of transactions into memory
            transactions = list(csvReader)

            if csvReader.fieldnames and "Report Name" in csvReader.fieldnames:
                # this must be a reimbursement file
                reimbursement = True
                payeeNameField = "Requester"
            else:
                # this must be a standard transactions file
                reimbursement = False
                payeeNameField = "Merchant Name"

        with qb.IQBSessionManager() as sessionManager:
            if not PreCheck(sessionManager, transactions, payeeNameField):
                return

            if reimbursement:
                count, respMsgSet = ProcessReimbursements(
                    sessionManager, transactions, payeeNameField
                )
            else:
                count, respMsgSet = ProcessTransactions(
                    sessionManager, transactions, payeeNameField
            )

        WalkRs(respMsgSet)

        print(f"Conversion complete, {count} transactions in {inputFileName}")

    except Exception as e:
        Error(f"Failed to process {inputFileName}: {e}")


if __name__ == '__main__':

    if len(sys.argv) != 2:
        print("usage:   Float2QB input.csv")
        # read the input filename from the console
        inputFileName = DeQuote(input("Enter the name of the input file: ").strip())
        main(inputFileName, "")
    else:
        main(sys.argv[1], "")

    # wait for keypress before closing the console window
    input("Press Enter to close this window")
    