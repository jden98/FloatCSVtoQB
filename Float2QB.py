# Script to convert CSV to IIF output.

import csv
import msvcrt
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path
import click
from win32com.client import Dispatch

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

def GetChar():
    return msvcrt.getch().decode("utf-8").lower()

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
    txnMemo = billRet.Memo.GetValue()
    txnTotal = billRet.AmountDue.GetValue()

    if statusCode == 0:
        click.echo(f"Created bill from {txnToAccount} for {txnTotal}")
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
        click.echo(f"Created deposit to {txnToAccount} for {txnTotal}")
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
        click.echo(f"Created cheque Number {txnRefNumber} to {txnPayee} for {txnTotal}")
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


def process_file(sessionManager, inputFileName):
    """Main processing function that handles the CSV to IIF conversion."""
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
@click.argument('input_file', type=click.Path(exists=True, dir_okay=False))
def main(input_file):
    """Convert Float CSV file to QuickBooks IIF format and import it.
    
    INPUT_FILE: Path to the CSV file to process
    """ 
    try:
        # Create session manager and begin session
        sessionManager = Dispatch("QBXMLRP2.RequestProcessor")
        sessionManager.OpenConnection("", "Float to QB Converter")
        
        process_file(sessionManager, input_file)
        
    except Exception as e:
        Error(f"Failed to process {input_file}: {e}")
    finally:
        try:
            sessionManager.CloseConnection()
        except:
            pass

if __name__ == '__main__':
    main()