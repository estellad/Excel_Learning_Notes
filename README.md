A markdown file dedicated to everything about Excel :sleepy: on Windows.

## Basics
### General
1. **Toggle between different sheets of the same file**: View > New Window
2. **Increase summary stats**: Highlight a range, then right click the bottom bar to check Average, Min, Max, Sum, etc.
3. **Conditional formating**: Home > Styles. This is for things like color coding/bold text on every other row = MOD(ROW(),2)=1, highlighting cells with "ERROR", cell value > 5.
4. **Group**: Data > Outline. This is to quickly hide/unhide things you want to present at a higher level or more detailed level. 
5. **Freeze top row or left column**: View > Freeze Panes. This is to fix header and rownames while scroll through. 
6. **Spell check**: highlight a range (or click on one cell for checking the entire sheet, or click Ctrl and select multiple tabs for checking the entire workbook), then F7 (fn + F7 if on laptop). If checking for the entire workbook, don't forget to ungroup the sheets after spell check, as any formatting afterwards would affect the entire workbook.
7. **Delimit data**: If separated by "|": Data > Text to Coluns > Delimited > Next > Other: | > Next > Finish. \
                     Side note for **exporting data from pdf**: Inside Adobe pdf document > Tools > Export PDF to any format > Spreadsheet > Export > Save.              
8. **Check format of data**: Home > Number > displayed in drop down box. To convert Text into General (number), use =value() function.
9. **References**: 
   - Relative reference: The cell reference moves the same distance and direction as you drag. Row and column free to change with fill - A1
   - Absolute reference: Row and column locked - $A$1
   - Partial absolute/Relative: 
    * Row free to change, column locked - $A1
    * Column free to change, row locked - A$1

### Like a Pro: No Mouse
**Press Alt**, then press a letter to go into a tab, then press another letter to go into a field. If you need to go back one level, press Esc. Do this for repetitive work, and it will be much faster than using mouse. 

### Highlighting, Inserting or Deleting
1.  **Highlighting**:
  - Highlight individual selection of cell by cell in one direction: Shift + Arrow
  - Navigate quickly to last populated cell in one direction: Ctrl + Arrow
  - Highlight and navigate to last populated cell: Ctrl + Shift + Arrow

2. **Insert or delete one/more cell between two cells**
  - Insert: Highlight second cell (for inserting one cell, if need to insert more, highlight N cells starting from the second cell), then Ctrl + Shift + (+)
  - Delete: Highlight the cell between (for deleting one cell, if need to delete more, highlight N cells in between), then Ctrl + (-)
 
3. **Insert or delete a column or a row**\
   *Highlight a column*: Ctrl + Space
  - Insert a column: Ctrl + Space, then Ctrl + Shift + (+)
  - Delete a column: Ctrl + Space, then Ctrl + (-) \
   *Highlight a row*: Shift + Space
  - Insert a row: Shift + Space, then Ctrl + Shift + (-)
  - Delete a row: Shift + Space, then Ctrl + (-)

4.  **Fill formula down and fill formula across**
    Manually highlight the range and then:
    - Down: Ctrl + D
    - Right: Ctrl + R
    Additional shortcuts:
    - Step into formula: F2
    - Change reference type: F4, keep hitting until you get to the correct type of reference you want. 


## Linking, Matching and Lookups
### Linking
1. **Paste link from output result**: To update an exhibit from a constantly updating file: copy cells from the old file, in the exhibit file, right click on first cell > Paste Special > Paste Link.
2. **Update links**: Data > Edit Links > click link > Change Source > select new file, then back in edit link pop up > Update Values > Close. 

### Matching
#### VLOOKUP
1. VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) could be used for simple matching on **content**. Example = VLOOKUP($G6, $B$5:$E$21, 2, 0).
  * lookup_value: usual a unique identifier, e.g. person id, drug din, etc.
  * table_array: source data. The data you are looking for will always be the first column in the selected table_array. 
  * col_index_num: the index in the table_array that you are trying to match from. 
  * range_lookup: TRUE - approx match; FALSE or 0 exact match. Often use 0 exact match.

2. Dragging with VLOOKUP using references
  For VLOOKUP, always fix the column and move the row, so the lookup_value is $G6

#### MATCH
MATCH(lookup_value, lookup_array, \[match_type\]) returns the **position** of this cell value located in lookup_array. It is rarely used singly by itself, but it is often used with INDEX-MATCH-MATCH. Example = MATCH($G8, $B$5:$F$5, 0).
  * lookup_value: lookup value.
  * lookup_array: source data, one row or one column. 
  * match_type: 1 by default, set to 0 for exact match.

#### INDEX
INDEX(array, row_number, column_number) returns the **content** of this cell value located in array, using coordinates. It is rarely used singly by itself, but it is often used with INDEX-MATCH-MATCH. Example = INDEX($E$8:$G$9, 4, 3). It is more dynamic than VLOOKUP, and it does not require the look up value to be in the first column. 

#### INDEX-MATCH-MATCH
INDEX(array, MATCH(lookup_value, lookup_array, \[match_type\]), MATCH(lookup_value, lookup_array, \[match_type\])) returns the **content**. This method is better than VLOOKUP, because we can match on any column, not necessarily the first column. Instead of using row number and column number, here we use match() to return them. \
Example = INDEX(Data!$1:$1048576, MATCH(Exhibit!$A6, Data!$A:$A, 0), MATCH(Exhibit!B$1, Data!$1:$1, 0)). Note that the entire sheet is selected, and the first match returns row index so it has column locked, and the second match returns column index so it has row locked. Now it can be dragged to both directions, because everything is locked properly. \
Note: a fast way to enter this formula without flipping between sheets is put INDEX(Data!$1:$1048576, MATCH(*XX*, Data!$A:$A, 0), MATCH(*XX*, Data!$1:$1, 0)) first, and then fill in $A6 and B$1 later (equivalent to without sheet name Exhibit!).

## Tables and Charts
### Tables
#### Table Tricks
1. **Make all columns wide enough** to display text: ALT + H + O + I. To make all columns as wide as the widest one, select column + right click + column width, check the value, then select all columns + right click + put the column width value in. 
2. **Add line break inside a cell**: (Enable by press Windows +) ALT + ENTER, then wrap text, widen the row. 
3. **Underline the cell**: Different than bottom border, as this is within the cell but same width as cell not as text. Ctrl + Shift + U

#### Table Formatting
1. **Align**: Text same length center it; text different length left align it. 
2. **Format align integer number**: 
  - Home > Number > ,
  - To format align the .00 decimals: right click > format (Ctrl + 1) > Custom > in Type enter: ?,??0 (the zeros will disappear after alignment)
3. **Format align float as percentage**:
  - Ctrl + 1 > Custom > in Type enter: ?0.0%
4. **Format a row number with a .**:
  - Ctrl + 1 > Custom > in Type enter: #.
5. **Superscript**: in a cell enter \[1\] at where you want to superscript, then highlight this, right click > Format > check superscript.

#### Table Print
1. Select all content > Page Layout > Page Setup > Print Area > Set Print Area > Ctrl + P > No Scaling > Fit Sheet on One Page 
2. Go back, Page Layout > Expand (bottom right) > Margin > check Horizontally > OK
3. Add Header/Footer: Insert > Text > Header & Footer
4. Page Layout > Header & Footer > uncheck Scale with document > OK
5. Spell check: F7

To have header appear on multiple pages: Page Layout > Print Titles > Rows to repeat at top > go select the title rows > OK.\
Also, can add a Total = SUM() row for calculation at the end.

### Charts
#### Chart Formatting
Select the data > Insert > Charts > so many up to you
1. **To make the graph on an independent sheet** from data: right click on the chart > Move chart > New sheet > name it > OK
2. Check with Ctrl + P, go back to **remove light grey outline**: right click > Format plot area > Plot Area Options > Chart Area > Border > No line 
3. **Change title**: click title > click 'Plus' sign > replace Chart title 
4. **Add vertical axis**: click 'Plus' sign > Axis Titles > uncheck Primary Horizontal > type in vertical axis. Make it read as horizontal: move it up > click on it > Format axis title > title options > Alignment > Text Direction > Horizontal.
5. **Format grid lines**: click on a grid line > Format Major Gridlines > Dash type > the fourth one
6. **Black text**: dark grey is default, but select the chart > then black text
7. **Font**: calibri is default, but select the chart > then Times New Roman
8. **Format horizontal dates**: select horizontal axis > Format Axis > Axis Options > last one (looks like bar) > go down > Number > Date > Type > Choose Mar-12 as March 2012. Also, Axis position > On tick marks. Also to label less densely every 6 moths, Axis Options > Units > Major > 6 Months
9. **Add legend**: click 'Plus' > Legend > Top. 
10. **Update legend name**: right click chart > Select Data > click on name of the series > Edit > change Series name > OK.
11. **Add Note**: Insert > Text Box
12. Then drag the graph around, drag axis label around, drag notes around, bold title, bold axis, black axis, and solid line on axis.
13. **Get Quater**: = CONCATENATE(YEAR(A4), "Q", CEILING((MONTH(A4)/3,1)), where A4 is 1/31/2003. Result is 2003Q1.
14. **Multilevel categorical axis lable**: One row by year, then one row by quater. Then plot the chart to see two layers of horizontal axis. Check in Format Axis > Labels > Multi-level Categorical Labels
15. **NA do not show in plot**: change 0 entry in line chart to NA = IF(C4=0, NA(), C4)

#### Advanced Chart Formatting


## Examples of File Structure
### Literature Review Sheet Structure
Title Page (tab white) | \
Search Strategy --> (tab dark blue; empty content)| Search strategy (tab light blue for the rest) | PRISMA | Level 1 Screening | Level 2 Screening | \
Extraction --> (tab dark orange, empty content)| Publication Details (tab light orange for the rest) | Method 1 | Method 2 | Method 3 | \
Narrative Synthesis --> (tab dark green; empty content)| Method 1 Findings (tab light green for the rest) | Method 2 Findings | Method 3 Findings | \
Appendix (tab dark grey; empty content) | Indications covered (tab light grey)


















