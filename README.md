# (VBA) EXCEL-MT940-GENERATOR (VBA)

## ABOUT THIS TOOL:
- THIS WAS FIRST CREATED FOR TESCO, A UK GROCERY STORE CHAIN. 

## Assumptions:
- Column D has a total row. All green cells have been filled out, yellow are optional. No cells should be deleted (unless deleting entire row). New rows should be added above the "total" row. Only delete rows with Data Type "Content". 											
																		
## Step 1:	
- Paste in your desired output path into cell "C4". 
- Example: C:\Users\avisb\Downloads\

## Step 2:
- Fill out all mandatory cells (green).																			

## Step 3:
- Highlight entire row 7 & Click "Export".
- 
## Step 4:
- Open file path set in step 2 and find new file saved with file name "MT940-Export_{timestamp}.txt".															

## Notes:
1. When clicking export, if you have any row highlighted other than the first statement transaction, it will start generating entries from that highlighted row and not take into consideration anything above it.
2. Totals row is generated within the macro. If the balance does not look correct on the totals row, ignore it - this is calculated and replaced within the macro itself.
3. Do not leave empty "CONTENT" rows or this will cause breaks in the VBA logic.
4. Input Debits as (-) Negatives and Credits as (+) Positives 																	
