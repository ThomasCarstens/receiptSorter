# receiptSorter
We classify and justify our company expenses. It currently takes 3h to gather each month's receipts with an additional 3-6h to link to bank statements.
I'm working on a solution that takes no more than 6h because multiple users can input receipts in any order and Bank Linking is semi-automated.

-> The **Smart Receipts App** takes receipts and outputs a ZIP (photos) and a CSV file (custom categories).
-> Send these to [email] and a spreadsheet will organise them among existing receipts.
-> The **Bank Linking Sheet** goes through all possible matches and outputs COMPTES_2021-2022.xlsx

## Exploration Phase.
### 1 June '22
Creating a Menu for the End-User. One menu item asks to import Bank Statements as New Sheets.

### 5 June '22
One menu item searches for bank entries with the same cost as the receipts. 

### 8 June '22
In case the same receipt is scanned twice, a script detects it and signals to user. 

### 13 June '22
SmartReceipt's output CSV does not join a receipt image so I set up a pipeline with Make.com that pushes all photos to Drive. You email the ZIP to a specific address to trigger the pipeline. I was able to get rid of the one-person-at-a-time bottleneck by extending the pipeline to multiple users.

### 15 June '22
The search scripts are made to ask permission before a match is joined.

### 20 June '22
CSVs of all three users are merged and organised by date right before Bank Linking. Duplicata trigger if new receipts are added. The search scripts focus on one month, to populate OUTPUT sheet. A progress bar is added to the search scripts. 

### 21 June '22
Bank Statements are tagged by user during Import, and extra email trigger is added. Finally, OUTPUT is linked to COMPTES_2021-2022.xlsx in a way that can be manually superceded. 

## User Testing.
...
<!-- 
<blockquote class="trello-board-compact">
    <a href="{https://trello.com/b/aMz841An/receipts-sorter}">Changelog</a>
    </blockquote>
    <script src="https://p.trellocdn.com/embed.min.js"></script> -->
