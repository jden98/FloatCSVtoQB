# Script to import CSV to QB.

import click
import csv
import locale
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path
from win32com.client import Dispatch

import QBComTypes as qb

def Error(message: str):
    """Log errors to stderr with traceback."""
    click.secho(f"Error: {message}", fg='red', err=True)
    if click.get_current_context().params['debug']:
        click.secho(traceback.format_exc(), fg='red', err=True)

def KeysLower(inDict: dict) -> dict:
    return {k.lower(): KeysLower(v) if isinstance(v, dict) else v for k, v in inDict.items()}

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


def VerifyCSVKeys(transactions: list[dict], reimbursement: bool, maxSplits: int | None = None) -> bool:
    """Verify that all required keys exist in the CSV transactions."""
    if not transactions:
        Error("No transactions found in CSV")
        return False

    first_trans = transactions[0]
    missing_keys = []

    # Common required keys for both types
    required_keys = ["description"]

    if reimbursement:
        # Reimbursement specific keys
        required_keys.extend([
            "expense date",
            "total",
            "subtotal",
            "tax",
            "requester",
            "gl code id"
        ])
    else:
        # Transaction specific keys
        required_keys.extend([
            "transaction date",
            "accounting vendor name",
            "total dollars",
            "transaction subtotal dollars",
            "transaction tax dollars",
            "gl code id"
        ])
        
        # Split transaction keys if needed
        if maxSplits:
            for i in range(1, maxSplits + 1):
                split_keys = [
                    f"line item {i} gl code id",
                    f"line item {i} description",
                    f"line item {i} amount",
                    f"line item {i} tax amount"
                ]
                required_keys.extend(split_keys)

    # Check for missing keys
    for key in required_keys:
        if key not in first_trans:
            missing_keys.append(key)

    if missing_keys:
        Error(f"Missing required keys in CSV: {', '.join(missing_keys)}")
        return False

    return True

def PreCheck(
    sessionManager: qb.IQBSessionManager,
    transactions: list[dict],
    reimbursement: bool,
    maxSplits: int | None = None,
) -> bool:
    """Pre-check the CSV file for valid accounts and vendors."""
    validAccounts, validVendors = LoadListsFromQB(sessionManager)

    good = True

    vendorName = "requester" if reimbursement else "accounting vendor name"
    
    # loop through transactions and check for bad accounts and vendors
    for t in transactions:
        if t[vendorName] not in validVendors:
            Error(f'Invalid {vendorName}: "{t[vendorName]}"')
            good = False

        if maxSplits and t['line item 1 amount'] != "": # if the csv has splits, check if this transaction does
            for i in range(1, maxSplits + 1):
                if t[f"line item {i} amount"] != "": # the determining factor for whether a split exist is if it has an amount
                    if t[f"line item {i} gl code id"] not in validAccounts:
                        Error(f'Invalid gl code id: "{t[f"line item {i} gl code id"]}" in split {i} in transaction for {t[vendorName]}')
                        good = False
        else: # this transaction doesn't have splits
            if t["gl code id"] not in validAccounts:
                Error(f'Invalid gl code id: {t["gl code id"]} in transaction for {t[vendorName]}')
                good = False

    return good


def WalkRs(respMsgSet: qb.IMsgSetResponse) -> bool:
    """Walk the response message set."""
    
    Success: bool = True
    if respMsgSet.responseList is None:
        return True

    for resp in respMsgSet.responseList:
        if resp.StatusCode >0:
            Error(f"Error: Code:{resp.StatusCode} Severity: {resp.StatusSeverity} Message: {resp.StatusMessage}")
            Success = False
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
                Success = False
    return Success

def WalkBillRet(billRet: qb.IBillRet, statusCode: int, statusSeverity: str, statusMessage: str) -> None:
    """Walk the bill return."""
    if billRet is None:
        return

    # Get value of TxnDate
    txnDate = billRet.TxnDate.GetValue()
    txnToAccount = billRet.VendorRef.FullName.GetValue()
    txnMemo = billRet.Memo.GetValue() if billRet.Memo is not None else ""
    txnTotal = billRet.AmountDue.GetValue()

    if statusCode == 0:
        click.echo(f"Created bill from {txnToAccount} for {locale.currency(txnTotal, grouping=True)}")
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
                    f"Error creating Bill {txnDate} {txnToAccount} {txnMemo} {locale.currency(txnTotal, grouping=True)} {lineAccount} "
                    f"{lineMemo} {locale.currency(lineAmount, grouping=True)}"
                )


def WalkDepositRet(depositRet: qb.IDepositRet, statusCode: int, statusSeverity: str, statusMessage: str) -> None:
    """Walk the deposit return."""
    if depositRet is None:
        return

    # Get value of TxnDate
    txnDate = depositRet.TxnDate.GetValue()
    txnToAccount = depositRet.DepositToAccountRef.FullName.GetValue()
    txnMemo = depositRet.Memo.GetValue() if depositRet.Memo is not None else ""
    txnTotal = depositRet.DepositTotal.GetValue()

    if statusCode == 0:
        click.echo(f"Created deposit to {txnToAccount} for {locale.currency(txnTotal, grouping=True)}")
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
                    f"Error creating Deposit {txnDate} {txnToAccount} {txnMemo} {locale.currency(txnTotal, grouping=True)} {lineAccount} "
                    f"{lineMemo} {locale.currency(lineAmount, grouping=True)}"
                )


def WalkCheckRet(checkRet: qb.ICheckRet, statusCode: int, statusSeverity: str, statusMessage: str) -> None:
    """Walk the check return."""
    if checkRet is None:
        return

    # Get value of TxnDate
    txnDate = checkRet.TxnDate.GetValue()
    txnToAccount = checkRet.AccountRef.FullName.GetValue()
    txnMemo = checkRet.Memo.GetValue() if checkRet.Memo is not None else ""
    txnTotal = checkRet.Amount.GetValue()
    txnRefNumber = checkRet.RefNumber.GetValue()
    txnPayee = checkRet.PayeeEntityRef.FullName.GetValue()

    if statusCode == 0:
        click.echo(f"Created cheque Number {txnRefNumber} to {txnPayee} for {locale.currency(txnTotal, grouping=True)}")
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
                    f"Error creating Cheque {txnDate} {txnToAccount} {txnMemo} {locale.currency(txnTotal,    grouping=True)} {lineAccount} "
                    f"{lineMemo} {locale.currency(lineAmount, grouping=True)}"
                )


def ProcessTransactions(
    sessionManager: qb.IQBSessionManager,
    transactions: list[dict],
    maxSplits: int | None,
) -> tuple[int, qb.IMsgSetResponse]:
    """Process the transaction data."""
    requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    count = 0
    for trans in transactions:
        trnsDate = datetime.strptime(trans["transaction date"], "%Y-%m-%d %H:%M:%S.%f%z")
        trnsMerch = trans["accounting vendor name"]
        trnsTotal = float(trans["transaction subtotal dollars"])
        trnsGlcode = trans["gl code id"]
        trnsTax = float(trans["transaction tax dollars"])
        trnsDesc = trans["description"]

        lineItems = []
        # if the file has any transactions with splits...
        if (not maxSplits) or trnsGlcode > "":
            # if this particular transaction has no splits
            lineItems.append({
                        'description': trnsDesc,
                        'total': trnsTotal,
                        'tax': trnsTax,
                        'glcode': trnsGlcode
                    })
        else:
            # if this particular transaction has splits
            for item in range(1, maxSplits + 1):
                # is there another split
                if trans.get(f"line item {item} gl code id", "") > "":
                    # Find corresponding fields for this line item
                    splitDesc = trans.get(f"line item {item} description", "")
                    splitTotal = float(trans.get(f"line item {item} amount", 0))
                    splitTax = float(trans.get(f"line item {item} tax Amount", 0))
                    splitGLCode = trans.get(f"line item {item} gl code id", "")

                    lineItems.append({
                        'description': splitDesc,
                        'total': splitTotal,
                        'tax': splitTax,
                        'glcode': splitGLCode
                    })

        if not lineItems:
            Error("Transaction has no detectable amounts or splits.")
            continue

        # if this is a deposit
        if trnsTotal < 0:
            # trnsType = "DEPOSIT"
            depAddRq = qb.IDepositAdd(requestMsgSet.AppendDepositAddRq())
            depAddRq.DepositToAccountRef.FullName.SetValue("Float Financial")
            depAddRq.TxnDate.SetValue(trnsDate)
            depAddRq.Memo.SetValue(trnsDesc)
            depLineAddRq: qb.IDepositLineAdd = depAddRq.DepositLineAddList.Append()
            depositInfo = depLineAddRq.ORDepositLineAdd.DepositInfo
            depositInfo.AccountRef.FullName.SetValue(trnsGlcode)
            depositInfo.Amount.SetValue(-1 * trnsTotal)
        else:
            # trnsType = "CHEQUE"
            chkAddRq = qb.ICheckAdd(requestMsgSet.AppendCheckAddRq())
            chkAddRq.AccountRef.FullName.SetValue("Float Financial")
            chkAddRq.IsToBePrinted.SetValue(False)
            chkAddRq.TxnDate.SetValue(trnsDate)
            chkAddRq.PayeeEntityRef.FullName.SetValue(trnsMerch)
            chkAddRq.Memo.SetValue(trnsDesc)

            for item in lineItems:
                expAdd: qb.IExpenseLineAdd = chkAddRq.ExpenseLineAddList.Append()
                expAdd.AccountRef.FullName.SetValue(item['glcode'])
                expAdd.Amount.SetValue(item['total'])
                expAdd.Memo.SetValue(item['description'])

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
) -> tuple[int, qb.IMsgSetResponse]:
    """Process the transaction data."""
    requestMsgSet = sessionManager.CreateMsgSetRequest("CA", 16, 0)
    requestMsgSet.Attributes.OnError = qb.constants.roeContinue

    count = 0
    for trans in transactions:
        trnsDate = datetime.strptime(trans["expense date"], "%d/%m/%Y")

        trnsDesc = trans["description"].strip()
        trnsMerch = trans["requester"]
        trnsGlcode = trans["gl code id"]

        try:
            trnsAmount = float(trans.get("total", 0))
            trnsSubtotal = float(trans.get("subtotal", 0))
            trnsTax = float(trans.get("tax", 0))
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


def ProcessFile(inputFileName):
    """Main processing function that handles the CSV import to QB"""
    count = 0
    inputFilePath = Path(inputFileName)

    try:
        with open(inputFilePath, "r", encoding="utf-8-sig") as inputFile:
            csvReader = csv.DictReader(inputFile)

            # load the full set of transactions into memory
            transactions = [KeysLower(d) for d in list(csvReader)]

        if "report name" in transactions[0]:
            # this must be a reimbursement file
            reimbursement = True
            maxSplits = None
        else:
            # this must be a standard transactions file
            reimbursement = False
            firstLine = transactions[0]
            pattern = r"^line item (\d+)"
            splits = [int(match.group(1)) for key in firstLine if (match := re.match(pattern, key))]
            maxSplits = max(splits) if splits else None

        # Verify CSV has all required keys
        if not VerifyCSVKeys(transactions, reimbursement, maxSplits):
            return

        with qb.IQBSessionManager() as sessionManager:
            if not PreCheck(sessionManager, transactions, reimbursement, maxSplits):
                return

            if reimbursement:
                count, respMsgSet = ProcessReimbursements(
                    sessionManager, transactions
                )
            else:
                count, respMsgSet = ProcessTransactions(
                    sessionManager, transactions, maxSplits
                )

        # if the response indicates success, prompt the user to delete input file
        if WalkRs(respMsgSet):
            click.echo(f"Conversion complete, processed {count} transactions from {inputFileName}")

            if click.confirm("Would you like to delete the input file?"):
                os.remove(inputFilePath) # delete the input file
        else:
            click.echo("Failed to import transactions to QuickBooks", err=True)
            
    except Exception as e:
        Error(f"Failed to process {inputFileName}: {e}")

@click.command()
@click.argument('input_file', type=click.Path(exists=True, dir_okay=False), required=False)
@click.option('--debug/--no-debug', default=False, help='Enable debug mode with full traceback')
def main(input_file, debug):
    """Import Float CSV file to QuickBooks.
    
    INPUT_FILE: Path to the CSV file to process
    """
    if not input_file:
        input_file = click.prompt('Please enter the path to your Float CSV file', type=str).strip().strip('"')
        if not os.path.exists(input_file):
            Error(f"Error: File '{input_file}' does not exist.")
            click.prompt('Press Enter to exit', default='', show_default=False)
            sys.exit(1)

    try:
        locale.setlocale(locale.LC_ALL, 'en_CA')
        ProcessFile(input_file)
        click.prompt('Press Enter to exit', default='', show_default=False)
    except Exception as e:
        Error(f"Error: {str(e)}")
        click.prompt('Press Enter to exit', default='', show_default=False)
        sys.exit(1)


if __name__ == '__main__':
    main()