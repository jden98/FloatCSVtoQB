Read exported CSV files produced by Float Financial, and, using the Quickbooks SDK version 16 QBFC, import transactions, deposits, and reimbursements to Quickbooks Desktop.

Float card transactions are imported as cheques (or checks in the US) written against the 'Float Financial' bank account.
Interest Payments and Float Card refunds are imported as Deposits to this account

Reimbursements are created as unpaid Bills against a vendor with the Float Spender's name. As of this writing, there's no indication in the exported CSV as to whether the reimbursement was paid through Float. If that gets added, then paid reimbursements could be imported as cheques, or as bills with cheques.

At this point, the Float 'Bank' account name is hard coded to 'Float Financial', and the accounts payable account for bills is hard coded as 'Accounts Payable'.

This application is written in Python 3, and the QBFC classes were generated with .makepy from win32com, then tweaked to add return typing information and iterators where needed.

The installer .iss file is for Inno Setup.
