# receiptSorter
We classify and justify our company expenses. It currently takes 9h to gather each month's receipts (3h) and find their corresponding entry in bank statements (3h) and then group cash receipts (2h).
I'm working on an automation solution that takes no more than 6h. 
- ultiple users can input receipts in any order, 
- only at the end do we open the spreadsheet for errors. 
- Each month is overwritten in COMPTES_2021-2022.xlsx

# Changelog.
### 1 June '22
Creating a Menu for the End-User. One menu item asks to import Bank Statements as New Sheets.

### 5 June '22
One menu item searches for bank entries with the same cost as the receipts. 

### 8 June '22
The **Bank Linking Sheet** goes through all possible matches and outputs COMPTES_2021-2022.xlsx In case the same receipt is scanned twice, a script detects it and signals to user. 

### 13 June '22
The **Smart Receipts App** takes receipts and outputs a ZIP (photos) and a CSV file (custom categories). 
SmartReceipt's output CSV does not join a receipt image so I set up a pipeline with Make.com that pushes all photos to Drive. 
You email the ZIP to our email address to trigger the pipeline. 
I extended the pipeline to multiple users. I reckon this solves the conundrum of having to open the doc for every receipt.

### 15 June '22
The **Bank Linking Sheet** goes through all possible matches and outputs COMPTES_2021-2022.xlsx However, the search scripts are made to ask permission before a match is joined.

### 20 June '22
The **Smart Receipts App** creates a ZIP file of receipts. Send it to our email address and a spreadsheet will organise them among existing receipts. CSVs of all three users are merged and organised by date right before Bank Linking. If new receipts are found, the duplicata script does a check.  

### 21 July '22
For now, Master Accounts is linked as a Google Sheet. main.testListLinks creates a centralised month, puts photo hyperlink and populates Master Accounts.

### 22 July '22
LinkerFunction (triggered from Sheet UI) uses main's findBankEntry() and Ticks Validated Receipts / cross out in Bank Statement.


# User Testing.
4 Users are registered and their Smart Receipts App is customised.
TODO: Issues 1 and 2.
<!--
<blockquote class="trello-board-compact">
    <a href="{https://trello.com/b/aMz841An/receipts-sorter}">Changelog</a>
    </blockquote>
    <script src="https://p.trellocdn.com/embed.min.js"></script> -->
