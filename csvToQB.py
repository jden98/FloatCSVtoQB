# Script to convert CSV to IIF output.

import csv
import sys
import traceback
from datetime import datetime
from pathlib import Path
from types import NoneType

import QBComTypes as qb


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
        if resp.StatusCode >= 0 and resp.Detail is not None:
            respType = int(resp.Type.GetValue())
            if respType == qb.ENResponseType.rtDepositAddRs:
                depositRet: qb.IDepositRet = qb.IDepositRet(resp.Detail)
                WalkDepositRet(depositRet)
            elif respType == qb.ENResponseType.rtCheckAddRs:
                checkRet: qb.ICheckRet = qb.ICheckRet(resp.Detail)
                WalkCheckRet(checkRet)
            else:
                Error(f"Unknown response type {qb.ENResponseType(respType).name}")


def WalkDepositRet(depositRet: qb.IDepositRet) -> None:
    """Walk the deposit return."""
    if depositRet is None:
        return

    # Get value of TxnDate
    txnDate = depositRet.TxnDate.GetValue()
    txnToAccount = depositRet.DepositToAccountRef.FullName.GetValue()
    txnMemo = depositRet.Memo.GetValue()
    txnTotal = depositRet.DepositTotal.GetValue()

    if depositRet.depositLineRetList is not None:
        for depositLineRet in depositRet.depositLineRetList:
            lineAccount = ""
            if depositLineRet.AccountRef is not None:
                accountRef = depositLineRet.AccountRef
                lineAccount = accountRef.FullName.GetValue()
            lineMemo = depositLineRet.Memo.GetValue()
            lineAmount = depositLineRet.Amount.GetValue()
            Error(
                f"Deposit {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} "
                f"{lineMemo} {lineAmount}"
            )


def WalkCheckRet(checkRet: qb.ICheckRet) -> NoneType:
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
            Error(
                f"Cheque {txnDate} {txnToAccount} {txnMemo} {txnTotal} {lineAccount} "
                f"{lineMemo} {lineAmount}"
            )


def ProcessTransactions(
    sessionManager: qb.IQBSessionManager,
    transactions: list[dict],
    payeeNameField: str,
    reimbursement: bool,
) -> int:
    """Process the input file."""
    requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    count = 0
    for trans in transactions:
        if reimbursement:
            trnsDate = datetime.strptime(trans["Expense Date"], "%d/%m/%Y")
        else:
            trnsDate = datetime.strptime(trans["Expense Date"], "%Y-%m-%d")

        trnsDesc = trans["Description"].strip()
        trnsMerch = trans[payeeNameField]
        trnsGlcode = trans["GL Code ID"]

        try:
            trnsAmount = -1 * float(trans.get("Total", 0))
            trnsSubtotal = float(trans.get("Subtotal", 0))
            trnsTax = float(trans.get("Tax", 0))
        except ValueError:
            Error("Invalid number format in transaction.")
            continue

        if (
            trnsAmount < 0
            and trnsGlcode == "Other Income:Interest Income"
            and not reimbursement
        ):
            # trnsType = "DEPOSIT"
            depAddRq = qb.IDepositAdd(requestMsgSet.AppendDepositAddRq())
            depAddRq.DepositToAccount.FullName.SetValue("Float Financial")
            depAddRq.TxnDate.SetValue(trnsDate)
            depAddRq.Memo.SetValue(trnsDesc)
            depLineAddRq: qb.IDepositLineAdd = depAddRq.DepositLineAddList.Append()
            depositInfo = depLineAddRq.ORDepositLineAdd.DepositInfo
            depositInfo.AccountRef.FullName.SetValue(trnsGlcode)
            depositInfo.Amount.SetValue(trnsAmount)

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

    WalkRs(respMsgSet)

    return count


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

            count = ProcessTransactions(
                sessionManager, transactions, payeeNameField, reimbursement
            )

        print(f"Conversion complete, {count} transactions in {inputFileName}")

    except Exception as e:
        Error(f"Failed to process {inputFileName}: {e}")


if __name__ == "__main__":
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("usage:   Convert2IIF input.csv output.iif")
    elif len(sys.argv) == 2:
        main(sys.argv[1], "")
    else:
        main(sys.argv[1], sys.argv[2])
