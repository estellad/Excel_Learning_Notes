A markdown file dedicated to everything about Excel :sleepy:  on Windows. Life as a consultant, you know ~

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
VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) could be used for simple matching on **content**. \
  *Example = VLOOKUP($G6, $B$5:$E$21, 2, 0)*.
  * lookup_value: usual a unique identifier, e.g. person id, drug din, etc.
  * table_array: source data. The data you are looking for will always be the first column in the selected table_array. 
  * col_index_num: the index in the table_array that you are trying to match from. 
  * range_lookup: TRUE - approx match; FALSE or 0 exact match. Often use 0 exact match.

Dragging with VLOOKUP using references: For VLOOKUP, always fix the column and move the row, so the lookup_value is $G6

#### MATCH
MATCH(lookup_value, lookup_array, \[match_type\]) returns the **position** of this cell value located in lookup_array. It is rarely used singly by itself, but it is often used with INDEX-MATCH-MATCH. \
  *Example = MATCH($G8, $B$5:$F$5, 0)*.
  * lookup_value: lookup value.
  * lookup_array: source data, one row or one column. 
  * match_type: 1 by default, set to 0 for exact match.

#### INDEX
INDEX(array, row_number, column_number) returns the **content** of this cell value located in array, using coordinates. It is rarely used singly by itself, but it is often used with INDEX-MATCH-MATCH. It is more dynamic than VLOOKUP, and it does not require the look up value to be in the first column. \
  *Example = INDEX($E$8:$G$9, 4, 3)*. 

#### INDEX-MATCH-MATCH
INDEX(array, MATCH(lookup_value, lookup_array, \[match_type\]), MATCH(lookup_value, lookup_array, \[match_type\])) returns the **content**. This method is better than VLOOKUP, because we can match on any column, not necessarily the first column. Instead of using row number and column number, here we use match() to return them. \
  *Example = INDEX(Data!$1:$1048576, MATCH(Exhibit!$A6, Data!$A:$A, 0), MATCH(Exhibit!B$1, Data!$1:$1, 0)).* 
- Note that the entire sheet is selected, and the first match returns row index so it has column locked, and the second match returns column index so it has row locked. Now it can be dragged to both directions, because everything is locked properly. 
- Note: a fast way to enter this formula without flipping between sheets is put INDEX(Data!$1:$1048576, MATCH(*XX*, Data!$A:$A, 0), MATCH(*XX*, Data!$1:$1, 0)) first, and then fill in $A6 and B$1 later (equivalent to without sheet name Exhibit!).

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
8. **Format horizontal dates**: select horizontal axis > Format Axis > Axis Options > last one (looks like bar) > go down > Number > Date > Type > Choose Mar-12 as March 2012. Also, Axis position > On tick marks. Also to label less densely every 6 months, Axis Options > Units > Major > 6 Months
9. **Add legend**: click 'Plus' > Legend > Top. 
10. **Update legend name**: right click chart > Select Data > click on name of the series > Edit > change Series name > OK.
11. **Add Note**: Insert > Text Box
12. Then drag the graph around, drag axis label around, drag notes around, bold title, bold axis, black axis, and solid line on axis.
13. **Get Quater**: = CONCATENATE(YEAR(A4), "Q", CEILING((MONTH(A4)/3,1)), where A4 is 1/31/2003. Result is 2003Q1.
14. **Multilevel categorical axis lable**: One row by year, then one row by quater. Then plot the chart to see two layers of horizontal axis. Check in Format Axis > Labels > Multi-level Categorical Labels
15. **NA do not show in plot**: change 0 entry in line chart to NA = IF(C4=0, NA(), C4)

#### Advanced Chart Formatting
- **Overlay a bar chart to an existing line chart**:
  * Create a new line: right click plot > select data > add series, etc.
  * Convert the new line to bar plot: right click > Change chart type > bar > Combo > select the series > Clustered Column > check Secondary Axis
  * Add a axis title by click the plus sign, then right click the axis title > format axis > Title options > Text direction > Horizontal. Then align with left axis.
  * Good practice :joy: change bar graph color from green to purple: right click the bar > Fill and Outline both to purple. 

- **Error bar**:
  * right click plot > select data > add series > add error series already made > enter Series Name = "Error Bars" and Series Value >
  * Insert error bar: Design > Add Chart Element > Error Bars > More Error Bar Options > select series name "Error Bars"
  * Format error bat: right click error bar > Format Error Bars > Error Bar Options > End Style > Cap/Minus \
                                                                                  > Error Amount > Percentage > 100% \
                                                                                  > Dash type > dotted 
  * Delete error bar from the legend
  * Format the legend to center: right click on legend > Format legend > Legend Options > Legend Position > Top > Press Shift and Down to make sure it's perfectly centered
 
- **Linked text box for the error bar to reflect updates and explain**: 
  * click plot > Insert > Text Box > drag a text box on the plot at the error bar > click into the text box to enter a formula to link > = "some cell name"
  * (Sometimes it's easier to just have a not linked text box for formatting and character limit reason)

- **Streamline a template** 
  A template will save all your font, all of your formatting, notes, and sources 
  * Repeatitive plot of this type: right click the plot > Save as Template > save it somewhere
  * To access a template: Design > Change Chart Type > Templates > select it

### Box and Whisker Plot
You will need to create rows of summary stats (\[A\]min, \[B\]25%, \[C\]median, \[D\]75%, \[E\]max), and then rows of ploting stats based on summary stats(\[B\] - \[A\]min whisker, \[B\]blank dummy series (exist for the purpose of spacing), \[C\]-\[B\]25-50, \[D\]-\[C\]50-75, \[E\]-\[D\]max whisker) for each category :sweat: 
1. Highlight rows except min and max whiskers, and then Insert > Charts > Stacked Column Charts 
2. right click on chart > Move chart > New sheet: Box Intermediate
3. click on the bottom chunk of stacked bars > Format Data Series > Series Options > Fill > No Fill \
                                                                                   > Border > No line
4. Remove from legend "Blank dummy series"
5. Click Plus sign for Add Chart Element > Error Bars > More Options > Blank dummy series (for first adding the error bar at the bottom)
6. click on error bar > Format Error Bars > Error Bar Options > Directions > Minus \
                                                              > End Style > Cap \
                                                              > Error Amount > Custom > Specify Value > Negative Error Value > Select data series Min Whisker > OK
7. Do the same for top whisker, but instead of doing it for Blank dummy series, do it for 50-75. Also, add error bar for Positive Error Value > select Max Whisker. 
9. Color: light blue for the 50-75 top half box, and dark blue for bottom half box. 
10. Formatting, thicken the outline of boxes and whiskers, light grey dotted grid, add title, axis and footnote. 

### Timelines / Swimlane Plot
Plot time and height (for the purpose of not overlaying the text) of the events as a scattered plot, solidify the verticle lines to create landmarks. Next to each dot, add some data label text to specify the event. 

### Waterfalls Plot
Just stacked bar chart, but hiding some chunks to invisible.

### Shading Plot
To shade area between two lines, just add Area components to right click > Change Chart Type > Combo > for each line add a shading variable with Chart Type: Area > OK\
The shading variables are just the copy of the line variables.

### Combine multiple charts as an Exhibit
Suppose the scenario that you have 2 charts on 2 different sheets
1. Copy the 2 sheets A, B by right click > Move or Copy > check Create a copy > OK, then you have A(2) and B(2), and also create an Exhibit sheet.
2. In A, click on chart > right click > Move Chart > Object in > Exhibit > OK. (Original A tab disappeared)
3. Do the same for B
4. In Exhibit > View > Workbook Views > Page Break > Preview > move both plots to page 1 (and page 2 disappear).
5. Add a title and stuff, and do Printing checks

## Text Manipulation
### Text Basics
In my honest opionion, just learn regular expression in R or Python, but:
1. =LEN(cell) returns the number of characters in the cell, including spaces in the middle and on both ends
2. =PROPER(cell or CELL or "cell cell") returns the cell as first letter capitalized, but the following letters in small case for each word. Similar for UPPER() and LOWER()
3. =CONCATENATE(cell&" "&cell&" "&cell&" "&cell)
4. =FIND(",", cell) returns common is the 6th character. FIND() is case sensitive, while SEARCH() is not
5. =LEFT(cell, FIND()-1) returns the first several characters from the left.
6. =MID(cell, start=FIND()+1, end=FIND()-FIND()-1) equals substring() in R :joy:
7. =SUBSTITUTE(cell, "old", "new") replace text
8. =RIGHT(cell, 4) returns last 4 characters starting from the right
9. =TRIM() removes any leading or trailing spaces from a character string

### Filter and Sorting
#### Filter
1. Apply filter: Ctrl + Shift + L
2. Note: If a data has a break row, then filter will not be applied to all the data. Also if there is an empty column between the last two columns, and you press Ctrl + Shift + L, the last column is not added with a filter. Therefore, highlight all the rows and columns you want to apply a filter to
3. There can also be filter by color, filter by number ranges, etc
4. Filter can only be removed one by one. To remove all, click Clear in Data > Sort & Filter > Clear

#### Sorting
1. After adding filter, click the small drop down, and sort from oldest to newest/vice versa/by color
2. Once sorted, impossible to go back unless pressing Ctrl + Z
3. To sort by multiple variables: Home > Sort & Filter > Custom Sort > Add Level > Sort by: , Then by: > OK

### Conditional Logic
1. =IF(cell..., T, F), use together with AND(cond1,cond2) and OR()
2. istext(), isnumber(), isblank(), isna(), islogical() can be used with if()
3. iferror() to trap errors that has NA or blank or etc

### Sumifs



### Exhibits to Word
#### Paste Chart
1. **As a picture**
  * select plot > Home > Clipboard > Copy > Copy as Picutre > As shown when printed > OK, and Ctrl + V. -> This will remove the dark grey dotted grey bakground. 
  * Paste Special: Ctrl + C, in Word (> Home > Clipboard > Paste > Paste Special >) Ctrl + Alt + V > Picture Enhanced Metafile > OK. -> Sharper than standard plot. It preserves the formatting, size, and layout. However, it would make chart text hard to read and make adjusting plot size in Word difficult. 
2. **As an object**
  * Ctrl + C plot, then in Word > Ctrl + Alt + V > Microsoft Office Graphic Objecct. This allows the plot to be easily resized and reformatted. However, text might get messy when pasted in. Can help with linking, but not recommended.

**Formatting Tips for Chart**
1. Figure title, notes and sources should be typed in Word, so no need to copy together with the chart
2. Adjust font size of axis labels, text boxes, etc. to 12pt before pasting
3. A standard chart copied would be pasted with dimension 6.5" width \* 4.72" height. To make chart vertically smaller, **make it an object** in Excel before copying:
  *  right click chart > Move Chart > Object in > Chart - As Object > OK. This allow to adjust the size of the chart in Excel.

#### Pasting Table
1. **As a picture**: same as chart's. Not recommended. 
2. **Using Word's formatting**, which would disgard any Excel formatting: Ctrl + C and Ctrl + V

**Formatting Tips for Tables**
1. Same as chart's
2. **Adjusting the table gridline formatting in Excel**, go to View > Normal View > uncheck Gridlines

These guidelines can also apply for PowerPoints. 



## Examples of File Structure
### Literature Review Sheet Structure
- Title Page (tab white) | 
- Search Strategy --> (tab dark blue; empty content)| Search strategy (tab light blue for the rest) | PRISMA | Level 1 Screening | Level 2 Screening |
- Extraction --> (tab dark orange, empty content)| Publication Details (tab light orange for the rest) | Method 1 | Method 2 | Method 3 | 
- Narrative Synthesis --> (tab dark green; empty content)| Method 1 Findings (tab light green for the rest) | Method 2 Findings | Method 3 Findings | 
- Appendix (tab dark grey; empty content) | Indications covered (tab light grey)


















