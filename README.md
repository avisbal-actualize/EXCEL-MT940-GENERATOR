# EXCEL-MT940-GENERATOR
VBA MT940 GENERATOR

ABOUT THIS TOOL:
- THIS WAS FIRST CREATED FOR TESCO, A UK GROCERY STORE CHAIN. 

Assumptions:	Column K has a total row. All green cells have been filled out, yellow are optional. No cells should be deleted (unless deleting entire row). New rows should be added above the "total" row. Only delete rows with Data Type "Content". 																														
																		
Step 1:	Start top to bottom. Fill out all mandatory cells (green). Fill out any optional cells (yellow).

Step 2:	Highlight entire row 7 & Click "Export".																																	

Step 3:	Paste in your desired file path. To set new default path, open vba --> open DoExport module --> go to row 16 and paste desired default value. 	Example: C:\Users\avisb\Downloads\ 																		

Step 4:	Open file path used above and find new file saved with file name "MT940-Export". 																	
																		
Note: 	When clicking export, if you have any row highlighted other than the first statement transaction, it will start generating entries from that highlighted row and not take into consideration anything above it.  I do not suggest using this as the closing balance will still take into consideration all transaction amounts, not just from starting point.							
Note: 	Totals row is generated within the macro. If the balance does not look correct on the totals row, ignore it - this is calculated and replaced within the macro itself. 																																		
Note: 	Do not leave empty "CONTENT" rows or this will cause breaks in the VBA logic.																																	
Note: 	Input Debits as (-) Negatives and Credits as (+) Positives 																	
