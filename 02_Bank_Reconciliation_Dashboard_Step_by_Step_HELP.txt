MONTH-END BANK RECONCILIATION DASHBOARD
STEP-BY-STEP EXCEL HELP GUIDE
English Level: A1-B1

PROJECT PURPOSE
This workbook compares the Bank Statement with the GL Cash Book.
It finds GL transactions that do not match the bank data.
It shows the problem items in an Exceptions sheet.
It also shows totals and counts in a Dashboard.

============================================================
1. CREATE THE WORKBOOK
============================================================

1. Open Microsoft Excel.
2. Create a Blank Workbook.
3. Create five worksheets.
4. Rename the worksheets exactly:

   README
   Bank_Statement
   GL_Cash_Book
   Exceptions
   Dashboard

5. Save the file as:

   02_Month_End_Bank_Reconciliation_Dashboard.xlsx


============================================================
2. BUILD THE README SHEET
============================================================

1. Open the README sheet.
2. Enter this information:

   A1: Month-End Bank Reconciliation & Close Dashboard

   A4: Section
   B4: Details

   A5: Business Problem
   B5: Accounting teams must compare bank transactions with
       GL cash records and find unmatched items.

   A6: How to Use
   B6: Add bank and GL data, copy the formulas, review the
       Exceptions sheet, and check the Dashboard.

   A7: Excel / Finance Skills
   B7: COUNTIF, IF, ROUND, helper columns, INDEX/MATCH,
       IFERROR, exception reporting, and dashboard design.

   A8: Workbook Sheets
   B8: README, Bank_Statement, GL_Cash_Book, Exceptions,
       Dashboard

3. Make A1 bold and large.
4. Make row 4 a colored header.
5. Turn on Wrap Text for column B.


============================================================
3. BUILD THE BANK_STATEMENT SHEET
============================================================

1. Open the Bank_Statement sheet.
2. Enter these headers in row 1:

   A1: Bank_ID
   B1: Date
   C1: Reference
   D1: Description
   E1: Bank_Amount
   F1: Match_Key

3. Enter or paste the bank data in columns A to E.

Example:

   BNK-0001 | 2026-05-17 | REF-1001 | Deposit | -4477.35

4. In cell F2, enter this formula:

   =C2&"|"&ROUND(E2,2)

5. Press Enter.
6. Copy the F2 formula down to the last bank transaction.

WHAT THIS FORMULA DOES

It joins the Reference and Amount.

Example:

   REF-1001|-4477.35

This Match_Key is used to compare bank and GL transactions.

7. Format column B as Date.
8. Format column E as Currency or Number with two decimals.
9. Make row 1 bold with a colored background.
10. Turn on Filter for row 1.
11. Freeze the top row.


============================================================
4. BUILD THE GL_CASH_BOOK SHEET
============================================================

1. Open the GL_Cash_Book sheet.
2. Enter these headers in row 1:

   A1: GL_ID
   B1: Posting_Date
   C1: Reference
   D1: Description
   E1: GL_Amount
   F1: Match_Key
   G1: Recon_Status
   H1: Variance
   I1: Exception_Seq

3. Enter or paste the GL data in columns A to E.

Example:

   GL-0001 | 2026-05-22 | REF-1001 | Cash posting | -4477.35

4. In cell F2, enter:

   =C2&"|"&ROUND(E2,2)

5. In cell G2, enter:

   =IF(COUNTIF(Bank_Statement!$F:$F,F2)>0,"Matched","Unmatched GL Item")

WHAT THIS FORMULA DOES

It searches for the GL Match_Key in the Bank_Statement Match_Key column.

Result:

   Matched
   or
   Unmatched GL Item

6. In cell H2, enter:

   =IF(G2="Matched",0,E2)

WHAT THIS FORMULA DOES

If the item is matched, the variance is zero.
If the item is not matched, it shows the GL amount.

7. In cell I2, enter:

   =IF(G2<>"Matched",COUNTIF($G$2:G2,"<>Matched"),"")

WHAT THIS FORMULA DOES

It gives each unmatched item a sequence number:

   1, 2, 3, 4, ...

The Exceptions sheet uses this number.

8. Copy formulas F2:I2 down to the last GL transaction.
9. Format column B as Date.
10. Format columns E and H as Currency or Number with two decimals.
11. Make row 1 bold with a colored background.
12. Turn on Filter.
13. Freeze the top row.


============================================================
5. ADD CONDITIONAL FORMATTING
============================================================

1. Select the Recon_Status cells in column G.
2. Go to Home > Conditional Formatting > New Rule.
3. Use a formula for matched items:

   =$G2="Matched"

4. Select a green fill.
5. Add another rule:

   =$G2="Unmatched GL Item"

6. Select a red or light red fill.

Now the status is easier to review.


============================================================
6. BUILD THE EXCEPTIONS SHEET
============================================================

1. Open the Exceptions sheet.
2. Enter these headers in row 1:

   A1: GL_ID
   B1: Posting_Date
   C1: Reference
   D1: Description
   E1: GL_Amount
   F1: Match_Key
   G1: Recon_Status
   H1: Variance

3. In cell A2, enter this formula:

   =IFERROR(INDEX(GL_Cash_Book!A$2:A$116,MATCH(ROW()-1,GL_Cash_Book!$I$2:$I$116,0)),"")

4. Copy the A2 formula across from A2 to H2.
5. Copy the formulas down for enough rows.

Example:

   Copy A2:H2 down to row 200.

IMPORTANT

If your GL data has more than 116 rows, change 116 in the formula.

Example for 500 GL rows:

   =IFERROR(INDEX(GL_Cash_Book!A$2:A$500,MATCH(ROW()-1,GL_Cash_Book!$I$2:$I$500,0)),"")

HOW THE FORMULA WORKS

1. Exception_Seq gives unmatched items numbers 1, 2, 3, and so on.
2. MATCH finds each sequence number.
3. INDEX returns the related GL information.
4. IFERROR shows a blank cell when there are no more exceptions.

6. Format column B as Date.
7. Format columns E and H as Currency or Number.
8. Make row 1 bold with a colored background.
9. Turn on Filter.
10. Freeze the top row.


============================================================
7. BUILD THE DASHBOARD SHEET
============================================================

1. Open the Dashboard sheet.
2. Merge cells A1:L1.
3. Enter this title:

   Bank Reconciliation Dashboard

4. Merge cells A2:L2.
5. Enter this subtitle:

   Self-directed portfolio project with realistic sample data,
   formulas, controls, and dashboard

6. Enter these KPI labels:

   A4: KPI
   B4: Value

   A5: Bank Balance
   A6: GL Balance
   A7: Difference
   A8: Matched GL Items
   A9: Unmatched GL Items

7. Enter these formulas:

   B5:
   =SUM(Bank_Statement!$E$2:$E$121)

   B6:
   =SUM(GL_Cash_Book!$E$2:$E$116)

   B7:
   =B5-B6

   B8:
   =COUNTIF(GL_Cash_Book!$G:$G,"Matched")

   B9:
   =COUNTIF(GL_Cash_Book!$G:$G,"Unmatched GL Item")

8. Create the chart data:

   D4: Category
   E4: Value

   D5: Matched
   E5:
   =COUNTIF(GL_Cash_Book!$G:$G,"Matched")

   D6: Unmatched GL Item
   E6:
   =COUNTIF(GL_Cash_Book!$G:$G,"Unmatched GL Item")

9. Select D4:E6.
10. Go to Insert > Column Chart > Clustered Column.
11. Move the chart to the right side of the Dashboard.
12. Format B5:B7 as Currency.
13. Format B8:B9 and E5:E6 as whole numbers.
14. Make the title and headers bold.
15. Use simple professional colors.

IMPORTANT

If your data has more rows, change the last row in the SUM formulas.

Example:

   =SUM(Bank_Statement!$E$2:$E$500)
   =SUM(GL_Cash_Book!$E$2:$E$500)


============================================================
8. TEST THE WORKBOOK
============================================================

1. Find one matched GL transaction.
2. Change its GL amount by a small value.

Example:

   Bank amount: 1000.00
   GL amount:   1005.00

3. Check the Recon_Status column.
4. The result must change to:

   Unmatched GL Item

5. Open the Exceptions sheet.
6. The transaction must appear there.
7. Open the Dashboard.
8. The unmatched count must increase.

9. Change the GL amount back to the correct value.
10. The item must return to Matched.


============================================================
9. MONTH-END USE
============================================================

Every month:

1. Make a backup copy of the workbook.
2. Delete the old bank data in columns A to E.
3. Paste the new Bank Statement data.
4. Copy the Match_Key formula down.
5. Delete the old GL data in columns A to E.
6. Paste the new GL Cash Book data.
7. Copy formulas F to I down.
8. Review the Exceptions sheet.
9. Check the Bank Balance and GL Balance.
10. Investigate the Difference.
11. Correct accounting errors when needed.
12. Save the final month-end file.

Possible reasons for unmatched items:

   Bank fee not recorded in GL
   Deposit in transit
   Outstanding cheque
   Wrong amount
   Wrong reference
   Duplicate transaction
   Missing GL entry
   Timing difference


============================================================
10. IMPORTANT CONTROL POINTS
============================================================

1. The Reference must be correct.
2. The Amount sign must be correct.

   Positive and negative amounts are different.

3. Amounts are rounded to two decimals.
4. Extra spaces in the Reference can cause a problem.
5. Copy formulas to all new rows.
6. Check that the formula ranges include all data.
7. Do not type over formula columns F to I.
8. Review all unmatched items before closing the month.


============================================================
11. SIMPLE LIMITATION
============================================================

This workbook checks whether the same Reference and Amount exist
in the bank data.

It is a simple portfolio model.

It does not fully control duplicate one-to-one matching.
It also reports GL unmatched items only.
A more advanced version can also report bank-only items,
duplicate items, date differences, and amount tolerances.


============================================================
12. SKILLS SHOWN
============================================================

Finance and Accounting Skills:

   Bank Reconciliation
   General Ledger Review
   Month-End Close
   Variance Analysis
   Exception Reporting
   Transaction Investigation

Excel Skills:

   COUNTIF
   IF
   ROUND
   IFERROR
   INDEX
   MATCH
   ROW
   Text Concatenation
   Absolute and Relative References
   Helper Columns
   Conditional Formatting
   Data Filters
   Number and Date Formatting
   Dashboard Design
   Column Charts


============================================================
13. SHORT INTERVIEW EXPLANATION
============================================================

I created an Excel bank reconciliation dashboard.
It compares bank transactions with GL cash records.
It uses a Match_Key and COUNTIF to find unmatched items.
INDEX and MATCH create an automatic exception report.
The dashboard shows balances, differences, and match counts.
