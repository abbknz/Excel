AP Invoice Excel File – Help Guide

1. What is this file?

This Excel file checks company invoices.

An invoice is a bill from a vendor.

This file can check:
- Duplicate invoices
- Missing PO numbers
- Wrong tax
- Late invoices
- Approval problems

The file does not pay invoices.

The file only finds problems.


2. Excel Sheets

This file has four sheets:

README
This sheet explains the file.

Data
Enter invoice information in this sheet.
The formulas also check the invoices.

Exceptions
This sheet shows high-risk invoices.

Dashboard
This sheet shows a short report.


3. Create the Data Sheet

Open the Data sheet.

Write these names in row 1:

A - Invoice_ID
B - Vendor
C - Invoice_No
D - PO_No
E - Invoice_Date
F - Due_Date
G - Amount
H - Tax
I - Total
J - Payment_Status
K - Approval_Status
L - Duplicate_Flag
M - PO_Status
N - Tax_Check
O - Days_Overdue
P - Risk_Level
Q - Exception_Notes


4. Manual Columns

The user enters information in these columns:

A: Invoice ID
B: Vendor
C: Invoice Number
D: PO Number
E: Invoice Date
F: Due Date
G: Amount
H: Tax
J: Payment Status
K: Approval Status

Do not type information in the formula columns.


5. Formula Columns

These columns use formulas:

I: Total
L: Duplicate Flag
M: PO Status
N: Tax Check
O: Days Overdue
P: Risk Level
Q: Exception Notes


6. Total Formula

Go to cell I2.

Write this formula:

=G2+H2

This formula adds the amount and tax.

Example:
Amount: $1,000
Tax: $130
Total: $1,130


7. Duplicate Invoice Formula

Go to cell L2.

Write this formula:

=IF(COUNTIFS($B$2:$B$1000,B2,$C$2:$C$1000,C2,$I$2:$I$1000,I2)>1,"Duplicate Risk","OK")

This formula checks:
- Vendor name
- Invoice number
- Total amount

If the same invoice is in the file two times, the result is:

Duplicate Risk

If there is no problem, the result is:

OK


8. PO Check Formula

Go to cell M2.

Write this formula:

=IF(D2="","Missing PO","Matched")

If the PO number is empty, the result is:

Missing PO

If there is a PO number, the result is:

Matched


9. Tax Check Formula

Go to cell N2.

Write this formula:

=IF(ABS(H2-G2*0.13)<=1,"OK","Review Tax")

This formula checks 13% tax.

Example:
Amount: $1,000
Correct tax: $130

If the tax is correct, the result is:

OK

If the tax is not correct, the result is:

Review Tax

The 13% tax is only an example.

For real work, use the correct tax rate.


10. Late Invoice Formula

Go to cell O2.

Write this formula:

=IF(J2="Paid",0,MAX(0,TODAY()-F2))

This formula checks the due date.

If the invoice is paid, the result is zero.

If the invoice is late, the formula shows the number of late days.

Set column O to Number format.

Do not use Date format.


11. Risk Level Formula

Go to cell P2.

Write this formula:

=IF(OR(L2="Duplicate Risk",M2="Missing PO",N2="Review Tax",O2>30,K2<>"Approved"),"High","Low")

The invoice is High Risk when:
- It may be a duplicate.
- It has no PO number.
- The tax is not correct.
- It is more than 30 days late.
- It is not approved.

If there is no problem, the risk is:

Low


12. Exception Notes Formula

Go to cell Q2.

Write this formula:

=TEXTJOIN("; ",TRUE,IF(L2="Duplicate Risk","Duplicate risk",""),IF(M2="Missing PO","Missing PO",""),IF(N2="Review Tax","Tax problem",""),IF(O2>0,"Overdue",""),IF(K2<>"Approved","Approval problem",""))

This formula writes the invoice problems.

Example:

Missing PO; Tax problem; Overdue


13. Copy the Formulas

Select these cells:
- I2
- L2 to Q2

Copy the formulas down.

You can copy them to row 1000.

Do not change the formulas in each row.


14. Payment Status List

Select cells J2:J1000.

Go to:

Data → Data Validation

Choose:

List

Write:

Paid,Unpaid,On Hold

Now the user can select the payment status.


15. Approval Status List

Select cells K2:K1000.

Go to:

Data → Data Validation

Choose:

List

Write:

Approved,Pending,Rejected

Now the user can select the approval status.


16. Exceptions Sheet

Open the Exceptions sheet.

Copy the Data sheet headers to row 1.

In cell A2, write:

=FILTER(Data!A2:Q1000,Data!P2:P1000="High","No high-risk invoices")

This formula shows only High Risk invoices.

This formula works in Excel 365.

Do not use full columns like A:Q.

A small range makes the file faster.


17. Dashboard Sheet

Open the Dashboard sheet.

Write these KPI names:

A5 - Total Invoices
A6 - Total AP Amount
A7 - Duplicate Risks
A8 - Missing PO
A9 - Tax Reviews
A10 - High Risk

Use these formulas:

Total Invoices
In cell B5:

=COUNTA(Data!$A$2:$A$1000)

Total AP Amount
In cell B6:

=SUM(Data!$I$2:$I$1000)

Duplicate Risks
In cell B7:

=COUNTIF(Data!$L$2:$L$1000,"Duplicate Risk")

Missing PO
In cell B8:

=COUNTIF(Data!$M$2:$M$1000,"Missing PO")

Tax Reviews
In cell B9:

=COUNTIF(Data!$N$2:$N$1000,"Review Tax")

High Risk
In cell B10:

=COUNTIF(Data!$P$2:$P$1000,"High")


18. Test the File

Enter one good invoice.

The result should be:
- Duplicate Flag: OK
- PO Status: Matched
- Tax Check: OK
- Days Overdue: 0
- Risk Level: Low

Enter the same invoice two times.

The result should be:

Duplicate Risk

Leave the PO number empty.

The result should be:

Missing PO

Enter the wrong tax.

The result should be:

Review Tax

Use an old due date.

The file should show late days.

Change Approval Status to Pending.

The result should show:

Approval problem


19. How to Use the File at Work

Step 1: Receive an invoice.
Step 2: Enter the invoice in the Data sheet.
Step 3: Check the formula results.
Step 4: Open the Exceptions sheet.
Step 5: Review High Risk invoices.
Step 6: Fix the invoice problem.
Step 7: Send the correct invoice for payment.


20. Important Information

The Data sheet is the main sheet.

The user enters invoice information in the Data sheet.

The formulas check the invoice.

The Exceptions sheet shows problem invoices.

The Dashboard sheet shows the total numbers.

This file does not make payments.

This file does not create accounting entries.

It is an invoice checking tool.
