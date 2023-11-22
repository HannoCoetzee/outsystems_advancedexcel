using System;
using System.Collections;
using System.Data;
using OutSystems.HubEdition.RuntimePlatform;

namespace OutSystems.NssAdvanced_Excel {

	public interface IssAdvanced_Excel {

		/// <summary>
		/// Opens an existing workbook for editing by either specifying a name or the binary data.
		/// </summary>
		/// <param name="ssFileName">Location of the file that you want to open. Set to empty string &quot;&quot; when using binary data</param>
		/// <param name="ssBinary_Data">Binary data of the file that you want to open. Set to nullbinary() if using FileName</param>
		/// <param name="ssWorkbook">The workbook that you want to work with.</param>
		void MssWorkbook_Open(string ssFileName, byte[] ssBinary_Data, out object ssWorkbook);

		/// <summary>
		/// Select a worksheet by its index or by its name
		/// </summary>
		/// <param name="ssWorkbook">The workbook wherein the worksheet exists</param>
		/// <param name="ssWorksheetIndex">The index of the worksheet to find. Indexes start at 1</param>
		/// <param name="ssWorksheetName">The name of the worksheet to find</param>
		/// <param name="ssWorksheet">This is the worksheet object that you have been looking for,</param>
		void MssWorksheet_Select(object ssWorkbook, int ssWorksheetIndex, string ssWorksheetName, out object ssWorksheet);

		/// <summary>
		/// Creates a new excel workbook, optionally specifying the name of the fiirst sheet.
		/// </summary>
		/// <param name="ssNumberOfSheets">The number of sheets to add. Sheet names will be auto generated, i.e. Sheet1, Sheet2.</param>
		/// <param name="ssFirstSheetName">Specify the name of the initial sheet in the workbook. Default = &quot;Sheet1&quot;</param>
		/// <param name="ssSheetNames">List of new sheets to add, with at least a name specified. The index, if specified, will be used to add sheets in that order.
		/// FirstSheetName and NrSheets are ignored if SheetNames is populated</param>
		/// <param name="ssWorkbook">The newly created workbook</param>
		void MssWorkbook_Create(int ssNumberOfSheets, string ssFirstSheetName, RLNewSheetRecordList ssSheetNames, out object ssWorkbook);

		/// <summary>
		/// Get the in-memory binary data of the specified workbook
		/// </summary>
		/// <param name="ssWorkbook">The workbook you want the binary data for</param>
		/// <param name="ssBinaryData">The binary data of the file</param>
		void MssWorkbook_GetBinaryData(object ssWorkbook, out byte[] ssBinaryData);

		/// <summary>
		/// Rename a worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssName">The new name for the spreadsheet</param>
		void MssWorksheet_Rename(object ssWorksheet, string ssName);

		/// <summary>
		/// Closes the excel workbook
		/// </summary>
		/// <param name="ssWorkbook"></param>
		void MssWorkbook_Close(object ssWorkbook);

		/// <summary>
		/// Hides / Shows a Column passed by index
		/// </summary>
		/// <param name="ssWorksheet">The worksheet you want to work with.</param>
		/// <param name="ssColumn">The index of the column within the worksheet that you want to hide/show.</param>
		/// <param name="ssHidden">A Boolean value, set to True to hide the column, and to False to show the column.</param>
		void MssColumn_Hide_Show(object ssWorksheet, int ssColumn, bool ssHidden);

		/// <summary>
		/// Reads the value of a cell.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Name of the cell to read from, i.e. A4. Required if CellRow and CellNumber set to 0.</param>
		/// <param name="ssCellRow">Row number of the cell to read from. Required if CellName not set.</param>
		/// <param name="ssCellColumn">Column number of the cell to read from. Required if CellName not set.</param>
		/// <param name="ssCellValue">The value in the cell, as text.</param>
		/// <param name="ssReadText">If true always reads the cell value as text</param>
		void MssCell_Read(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, out string ssCellValue, bool ssReadText);

		/// <summary>
		/// Set protection on an Excel Worksheet
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to protect</param>
		/// <param name="ssPassword">DEPRECATED
		/// Can be used for backwards compatibility with Excel_Package
		/// 
		/// Password to protect the worksheet with.
		/// </param>
		/// <param name="ssProtectionOptions">Options to set when protecting the worksheet</param>
		void MssWorksheet_Protect(object ssWorksheet, string ssPassword, RCProtectionRecord ssProtectionOptions);

		/// <summary>
		/// Write a converted value to a cell.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides </param>
		/// <param name="ssCellName">Name of the cell to write to, i.e. A4. Required if CellRow and CellColumn not set</param>
		/// <param name="ssCellRow">Row number of the cell to write to. Required if CellName not set.</param>
		/// <param name="ssCellColumn">Column number of the cell to write to. Required if CellName not set.</param>
		/// <param name="ssCellValue">The value to write to the cell</param>
		/// <param name="ssCellType">Type can be:
		/// general (default if empty)
		/// text,
		/// datetime,
		/// integer,
		/// decimal,
		/// boolean,
		/// formula</param>
		/// <param name="ssCellFormat">CellFormat for the target cell</param>
		void MssCell_Write(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat);

		/// <summary>
		/// Write a dataset to a range of cells.
		/// Accepts format for the target cells
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssRowStart">Start row (integer)</param>
		/// <param name="ssColumnStart">Start column (integer)</param>
		/// <param name="ssDataSet">Data to write</param>
		/// <param name="ssCellFormat">CellFormat for the target cells</param>
		/// <param name="ssExportHeaders">True to include headers in export file. Default value = False</param>
		void MssCell_WriteRange(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord ssCellFormat, bool ssExportHeaders);

		/// <summary>
		/// Get the name of the given worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssWorksheetName"></param>
		void MssWorksheet_GetName(object ssWorksheet, out string ssWorksheetName);

		/// <summary>
		/// Get the properties of all of the worksheets in the workbook
		/// </summary>
		/// <param name="ssWorkbook">The workbook</param>
		/// <param name="ssProperties"></param>
		void MssWorksheet_GetPropertiesAll(object ssWorkbook, out RCWorkbookRecord ssProperties);

		/// <summary>
		/// Get the properties of the given worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssProperties"></param>
		void MssWorksheet_GetProperties(object ssWorksheet, out RCWorksheetRecord ssProperties);

		/// <summary>
		/// Add a worksheet to an existing workbook, optionally at the index specified. Specifying only a name will create a blank sheet. Specifying  a name with binary data, will add the sheet from the existing binary data, and then rename to the newly provided name
		/// </summary>
		/// <param name="ssWorkbook">The workbook that you want to add the sheet to</param>
		/// <param name="ssWorksheetName">The name of the worksheet you want to add. If binary data is nullbinary(), an empty sheet will be added</param>
		/// <param name="ssWorksheet">The worksheet object that you want to add. Set to nullbinary() if adding a new sheet by name</param>
		/// <param name="ssIndexWhereToAdd">The index where to add the new sheet. Default will be highest sheet index plus 1</param>
		void MssWorkBook_AddSheet(object ssWorkbook, string ssWorksheetName, object ssWorksheet, int ssIndexWhereToAdd);

		/// <summary>
		/// Delete a worksheet in a workbook by specifying either the index, or the name of the worksheet.
		/// </summary>
		/// <param name="ssWorkbook">The workbook from which you want to delete the worksheet</param>
		/// <param name="ssIndexToDelete">The index of the worksheet to delete. Set to 0 if using the worksheet name to delete</param>
		/// <param name="ssNameToDelete">The name of the worksheet to delete. Set to empty string &quot;&quot; if using the index to delete.</param>
		void MssWorksheet_Delete(object ssWorkbook, int ssIndexToDelete, string ssNameToDelete);

		/// <summary>
		/// Create a chart
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssChartType">Receives the chart type in text, possible types:
		/// Area3D
		/// AreaStacked3D
		/// AreaStacked1003D
		/// BarClustered3D
		/// BarStacked3D
		/// BarStacked1003D
		/// Column3D
		/// ColumnClustered3D
		/// ColumnStacked3D
		/// ColumnStacked1003D
		/// Line3D
		/// Pie3D
		/// PieExploded3D
		/// Area
		/// AreaStacked
		/// AreaStacked100
		/// BarClustered
		/// BarOfPie
		/// BarStacked
		/// BarStacked100
		/// Bubble
		/// Bubble3DEffect
		/// ColumnClustered
		/// ColumnStacked
		/// ColumnStacked100
		/// ConeBarClustered
		/// ConeBarStacked
		/// ConeBarStacked100
		/// ConeCol
		/// ConeColClustered
		/// ConeColStacked
		/// ConeColStacked100
		/// CylinderBarClustered
		/// CylinderBarStacked
		/// CylinderBarStacked100
		/// CylinderCol
		/// CylinderColClustered
		/// CylinderColStacked
		/// CylinderColStacked100
		/// Doughnut
		/// DoughnutExploded
		/// Line
		/// LineMarkers
		/// LineMarkersStacked
		/// LineMarkersStacked100
		/// LineStacked
		/// LineStacked100
		/// Pie
		/// PieExploded
		/// PieOfPie
		/// PyramidBarClustered
		/// PyramidBarStacked
		/// PyramidBarStacked100
		/// PyramidCol
		/// PyramidColClustered
		/// PyramidColStacked
		/// PyramidColStacked100
		/// Radar
		/// RadarFilled
		/// RadarMarkers
		/// StockHLC
		/// StockOHLC
		/// StockVHLC
		/// StockVOHLC
		/// Surface
		/// SurfaceTopView
		/// SurfaceTopViewWireframe
		/// SurfaceWireframe
		/// XYScatter
		/// XYScatterLines
		/// XYScatterLinesNoMarkers
		/// XYScatterSmooth
		/// XYScatterSmoothNoMarkers=73</param>
		/// <param name="ssChartName"></param>
		/// <param name="ssDataSeries_List">List Of DataSeries</param>
		/// <param name="ssHeight">Expressed in pixels</param>
		/// <param name="ssWidth">Expressed in pixels</param>
		/// <param name="ssRowPos">Row position to place the upper left corner graph</param>
		/// <param name="ssColPos">Column position to place the upper left corner graph</param>
		void MssChart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos);

		/// <summary>
		/// Inserts a new row into the spreadsheet.  Existing rows below the position are shifted down.  All formula are updated to take account of the new row.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to insert the row(s) into</param>
		/// <param name="ssInsertAt">The position of the new row
		/// </param>
		/// <param name="ssNrRows">Number of rows to insert</param>
		/// <param name="ssCopyStyleFromRow">Copy Styles from this row. Applied to all inserted rows. 0 will not copy any styles</param>
		void MssRow_Insert(object ssWorksheet, int ssInsertAt, int ssNrRows, int ssCopyStyleFromRow);

		/// <summary>
		/// Change the index of a worksheet in the document
		/// </summary>
		/// <param name="ssWorkbook">The workbook in which the change is to be made.</param>
		/// <param name="ssCurrentIndex">The current index(position) of the sheet in question</param>
		/// <param name="ssNewIndex">The new index for the sheet</param>
		void MssWorkbook_ChangeSheetIndex(object ssWorkbook, int ssCurrentIndex, int ssNewIndex);

		/// <summary>
		/// Apply a specified cell format to the range specified for the given worksheet
		/// </summary>
		/// <param name="ssWorksheet">Worksheet object where formatting is to be applied</param>
		/// <param name="ssCellFormat">CellFormat to apply</param>
		/// <param name="ssRange">Range that CellFormat is to be applied to</param>
		void MssCellFormat_ApplyToRange(object ssWorksheet, RCCellFormatRecord ssCellFormat, RCRangeRecord ssRange);

		/// <summary>
		/// Find all cells that contain the specified value in the given worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet in which to search</param>
		/// <param name="ssValueToFind">The value to search for</param>
		/// <param name="ssListOfCells">List of cells (ranges) where the value has been found</param>
		void MssCells_FindByValue(object ssWorksheet, string ssValueToFind, out RLRangeRecordList ssListOfCells);

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssRange"></param>
		/// <param name="ssValue"></param>
		/// <param name="ssParameter1"></param>
		/// <param name="ssFound"></param>
		/// <param name="ssRowIndex"></param>
		/// <param name="ssColumnIndex"></param>
		void MssContainInRange(object ssWorksheet, string ssRange, string ssValue, string ssParameter1, out bool ssFound, out int ssRowIndex, out int ssColumnIndex);

		/// <summary>
		/// Calculate all formulae for the entire workbook provided.
		/// </summary>
		/// <param name="ssWorkbook">The workbook to work with</param>
		void MssWorkbook_Calculate(object ssWorkbook);

		/// <summary>
		/// Calculate all formulae on the provided worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		void MssWorksheet_Calculate(object ssWorksheet);

		/// <summary>
		/// Hides / Shows Row passed by index
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to work with</param>
		/// <param name="ssRowIndex">Index of the Row to show/hide</param>
		/// <param name="ssHidden">A Boolean value, set to True to hide the row and to False to show the row</param>
		void MssRow_Hide_Show(object ssWorksheet, int ssRowIndex, bool ssHidden);

		/// <summary>
		/// Hide / Show a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssHidden">Visible = 0 - The worksheet is visible
		/// Hidden = 1 - The worksheet is hidden but can be shown by the user via the user interface
		/// VeryHidden = 2 - The worksheet is hidden and cannot be shown by the user via the user interface</param>
		void MssWorksheet_Hide_Show(object ssWorksheet, int ssHidden);

		/// <summary>
		/// Add a rule for conditionally formatting a range of cells.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssConditionalFormatRecord">The conditional formatting to apply to the Address Range</param>
		void MssConditionalFormatting_AddRule(object ssWorksheet, RCConditionalFormatItemRecord ssConditionalFormatRecord);

		/// <summary>
		/// Get a list of all the conditional formatting rules in a worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssListOfRule">List of conditional formatting rules</param>
		void MssConditionalFormatting_GetAllRules(object ssWorksheet, out RLConditionalFormatItemRecordList ssListOfRule);

		/// <summary>
		/// Merge cells in the range provided
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssRangeToMerge">The range of the cells to merge</param>
		void MssCell_Merge(object ssWorksheet, RCRangeRecord ssRangeToMerge);

		/// <summary>
		/// Un-Merge cells in the range provided
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssRangeToUnmerge">The range of cell to un-merge</param>
		void MssCell_UnMerge(object ssWorksheet, RCRangeRecord ssRangeToUnmerge);

		/// <summary>
		/// Delete row(s) from a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssStartRowNumber">Row number where to start deleting rows.</param>
		/// <param name="ssNumberOfRows">The number of rows to delete. Default = 1.</param>
		void MssRow_Delete(object ssWorksheet, int ssStartRowNumber, int ssNumberOfRows);

		/// <summary>
		/// Delete column(s) from a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssStartColumnNumber">Column number where to start deleting columns.</param>
		/// <param name="ssNumberOfColumns">The number of rows to delete. Default = 1.</param>
		void MssColumn_Delete(object ssWorksheet, int ssStartColumnNumber, int ssNumberOfColumns);

		/// <summary>
		/// Delete comment(s) in a specified range
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRange">Range to delete comments from.</param>
		void MssComment_Delete(object ssWorksheet, RCRangeRecord ssRange);

		/// <summary>
		/// Add a comment to a cell
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRowNumber">The row number of the cell to add the comment to.</param>
		/// <param name="ssColumnNumber">The column number of the cell to add the comment to.</param>
		/// <param name="ssText">The comment.</param>
		/// <param name="ssAuthor">The author of the comment.</param>
		/// <param name="ssAutofit">True to autofit the comment window to the comment text</param>
		void MssComment_Add(object ssWorksheet, int ssRowNumber, int ssColumnNumber, string ssText, string ssAuthor, bool ssAutofit);

		/// <summary>
		/// Inserts a new column into the spreadsheet.  Existing columns to the right of the insert index will be shifted right.  All formula are updated to take account of the new column.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssInsertAt">Column number where to insert new column.</param>
		/// <param name="ssNumberOfColumns">The number of columns to insert.</param>
		/// <param name="ssCopyStylesFrom">Copy Styles from this column. Applied to all inserted columns. 0 (default) will not copy any styles</param>
		void MssColumn_Insert(object ssWorksheet, int ssInsertAt, int ssNumberOfColumns, int ssCopyStylesFrom);

		/// <summary>
		/// Delete a specified Conditional Formatting rule on a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRuleToDeleteIndex">The index of the rule to be deleted.</param>
		void MssConditionalFormatting_DeleteRule(object ssWorksheet, int ssRuleToDeleteIndex);

		/// <summary>
		/// Delete ALL Conditional Formatting rules for a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		void MssConditionalFormatting_DeleteAllRules(object ssWorksheet);

		/// <summary>
		/// Insert an image into a Worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssImageFile">Binary data of the image to be inserted</param>
		/// <param name="ssImageType">File type. BMP, PNG, JPG</param>
		/// <param name="ssImageName">Name reference for the image in the Worksheet</param>
		/// <param name="ssRowNumber">Row index where to insert image. Ignored if CellName is specified</param>
		/// <param name="ssColumnNumber">Column index where to insert image. Ignored if CellName is specified</param>
		/// <param name="ssCellName">Cell Name where to insert image</param>
		/// <param name="ssImageWidth">The width of the image in pixels</param>
		/// <param name="ssImageHeight">The height of the image in pixels</param>
		/// <param name="ssMarginTop"> Offset in pixels </param>
		/// <param name="ssMarginLeft"> Offset in pixels</param>
		void MssImage_Insert(object ssWorksheet, byte[] ssImageFile, string ssImageType, string ssImageName, int ssRowNumber, int ssColumnNumber, string ssCellName, int ssImageWidth, int ssImageHeight, int ssMarginTop, int ssMarginLeft);

		/// <summary>
		/// Apply the column autofit action to the specified range of cells specified in the given worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		void MssWorksheet_AutofitColumns(object ssWorksheet);

		/// <summary>
		/// Add the automatic filter option of Excel to the specified range of cells.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRangeToFilter">The range where to add the filter. If not supplied, the dimension of the worksheet will be used.</param>
		void MssWorksheet_AddAutoFilter(object ssWorksheet, RCRangeRecord ssRangeToFilter);

		/// <summary>
		/// Set protection on the workbook level
		/// </summary>
		/// <param name="ssWorkbook">The workbook to work with</param>
		/// <param name="ssPassword">The password to set for the workbook. This does not encrypt the workbook.</param>
		/// <param name="ssLockStructure">Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.</param>
		/// <param name="ssLockWindows">Locks the position of the workbook window.</param>
		/// <param name="ssLockRevision">Lock the workbook for revision</param>
		void MssWorkbook_Protect(object ssWorkbook, string ssPassword, bool ssLockStructure, bool ssLockWindows, bool ssLockRevision);

		/// <summary>
		/// Calculates the formula of a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		void MssCell_CalculateByIndex(object ssWorksheet, int ssRow, int ssColumn);

		/// <summary>
		/// Calculates the formula of a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		void MssCell_CalculateByName(object ssWorksheet, string ssCellName);

		/// <summary>
		/// Apply format to a range of cells.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssRowStart">Start row (integer)</param>
		/// <param name="ssColumnStart">Start column (integer)</param>
		/// <param name="ssRowEnd">End row (integer)</param>
		/// <param name="ssColumnEnd">End column (integer)</param>
		/// <param name="ssCellFormat">CellFormat for the target cells</param>
		void MssCell_FormatRange(object ssWorksheet, int ssRowStart, int ssColumnStart, int ssRowEnd, int ssColumnEnd, RCCellFormatRecord ssCellFormat);

		/// <summary>
		/// Reads the value of a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">row number</param>
		/// <param name="ssColumn">column number</param>
		/// <param name="ssReadText">If true always reads the cell value as text</param>
		/// <param name="ssCellValue">text-value</param>
		void MssCell_ReadByIndex(object ssWorksheet, int ssRow, int ssColumn, bool ssReadText, out string ssCellValue);

		/// <summary>
		/// Reads the value of a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssReadText">If true always reads the cell value as text</param>
		/// <param name="ssCellValue">text-value</param>
		void MssCell_ReadByName(object ssWorksheet, string ssCellName, bool ssReadText, out string ssCellValue);

		/// <summary>
		/// Write a formula to a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">rownumber</param>
		/// <param name="ssColumn">columnnumber</param>
		/// <param name="ssFormula">Formula</param>
		void MssCell_SetFormulaByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssFormula);

		/// <summary>
		/// Write a formula to a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssFormula">Formula</param>
		void MssCell_SetFormulaByName(object ssWorksheet, string ssCellName, string ssFormula);

		/// <summary>
		/// Adds a copy of a worksheet
		/// </summary>
		/// <param name="ssWorkbook">The workbook in which the worksheet is to be copied
		/// </param>
		/// <param name="ssWorksheetName">The name of the spreadsheet to create</param>
		/// <param name="ssWorksheetToCopy">The worksheet to be copied</param>
		/// <param name="ssWorksheet">The copied worksheet</param>
		void MssWorkbook_AddCopyWorksheet(object ssWorkbook, string ssWorksheetName, object ssWorksheetToCopy, out object ssWorksheet);

		/// <summary>
		/// Get all images in a worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssImages"></param>
		void MssWorksheet_GetImages(object ssWorksheet, out RLImageRecordList ssImages);

		/// <summary>
		/// Select a worksheet by its index
		/// </summary>
		/// <param name="ssWorkbook">The worksheet to work with</param>
		/// <param name="ssWorksheetNumber">The index of the spreadsheet to select, starting at 1</param>
		/// <param name="ssWorksheet">The selected worksheet</param>
		void MssWorksheet_SelectByIndex(object ssWorkbook, int ssWorksheetNumber, out object ssWorksheet);

		/// <summary>
		/// Select a worksheet to work on by its name
		/// </summary>
		/// <param name="ssWorkbook">The workbook to work with</param>
		/// <param name="ssWorksheetName">The name of the spreadsheet to select</param>
		/// <param name="ssWorksheet">The selected worksheet</param>
		void MssWorksheet_SelectByName(object ssWorkbook, string ssWorksheetName, out object ssWorksheet);

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ssWorkbook"></param>
		/// <param name="ssNameToDelete"></param>
		void MssWorksheet_DeleteByName(object ssWorkbook, string ssNameToDelete);

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ssWorkbook"></param>
		/// <param name="ssIndexToDelete"></param>
		void MssWorksheet_DeleteByIndex(object ssWorkbook, int ssIndexToDelete);

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ssWorksheet">The worksheet you want to work with.</param>
		/// <param name="ssChartType">Receives the chart type in text, possible types:
		/// Area3D
		/// AreaStacked3D
		/// AreaStacked1003D
		/// BarClustered3D
		/// BarStacked3D
		/// BarStacked1003D
		/// Column3D
		/// ColumnClustered3D
		/// ColumnStacked3D
		/// ColumnStacked1003D
		/// Line3D
		/// Pie3D
		/// PieExploded3D
		/// Area
		/// AreaStacked
		/// AreaStacked100
		/// BarClustered
		/// BarOfPie
		/// BarStacked
		/// BarStacked100
		/// Bubble
		/// Bubble3DEffect
		/// ColumnClustered
		/// ColumnStacked
		/// ColumnStacked100
		/// ConeBarClustered
		/// ConeBarStacked
		/// ConeBarStacked100
		/// ConeCol
		/// ConeColClustered
		/// ConeColStacked
		/// ConeColStacked100
		/// CylinderBarClustered
		/// CylinderBarStacked
		/// CylinderBarStacked100
		/// CylinderCol
		/// CylinderColClustered
		/// CylinderColStacked
		/// CylinderColStacked100
		/// Doughnut
		/// DoughnutExploded
		/// Line
		/// LineMarkers
		/// LineMarkersStacked
		/// LineMarkersStacked100
		/// LineStacked
		/// LineStacked100
		/// Pie
		/// PieExploded
		/// PieOfPie
		/// PyramidBarClustered
		/// PyramidBarStacked
		/// PyramidBarStacked100
		/// PyramidCol
		/// PyramidColClustered
		/// PyramidColStacked
		/// PyramidColStacked100
		/// Radar
		/// RadarFilled
		/// RadarMarkers
		/// StockHLC
		/// StockOHLC
		/// StockVHLC
		/// StockVOHLC
		/// Surface
		/// SurfaceTopView
		/// SurfaceTopViewWireframe
		/// SurfaceWireframe
		/// XYScatter
		/// XYScatterLines
		/// XYScatterLinesNoMarkers
		/// XYScatterSmooth
		/// XYScatterSmoothNoMarkers=73</param>
		/// <param name="ssChartName"></param>
		/// <param name="ssDataSeries_List">List Of DataSeries</param>
		/// <param name="ssHeight">Expressed in pixels</param>
		/// <param name="ssWidth">Expressed in pixels</param>
		/// <param name="ssRowPos">Row position to place the upper left corner graph</param>
		/// <param name="ssColPos">Column position to place the upper left corner graph</param>
		void MssWorksheet_Chart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos);

		/// <summary>
		/// Create a defined &quot;Name&quot; (a word or string of characters in Excel that represents a cell, range of cells, formula, or constant value) in excel, starting in the RowStart / ColumnStart cell.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssName">&quot;Name&quot;</param>
		/// <param name="ssDataSet">Values to assigned the name</param>
		/// <param name="ssRowStart">Start row number</param>
		/// <param name="ssColumnStart">Start column number</param>
		void MssWorksheet_AddName(object ssWorksheet, string ssName, object ssDataSet, int ssRowStart, int ssColumnStart);

		/// <summary>
		/// Opens an existing workbook for editing and keeps it in memory
		/// </summary>
		/// <param name="ssBinaryData"></param>
		/// <param name="ssWorkbook"></param>
		void MssWorkbook_Open_BinaryData(byte[] ssBinaryData, out object ssWorkbook);

		/// <summary>
		/// Set the pixel width of a column on a specific worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssColumnNumber">The column number, starting at 1</param>
		/// <param name="ssDesiredWidth">The pixel width you desire for the column.</param>
		void MssColumn_SetWidth(object ssWorksheet, int ssColumnNumber, decimal ssDesiredWidth);

		/// <summary>
		/// Set the pixel height for a specific row in a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with
		/// </param>
		/// <param name="ssRowNumber">The number of the row to set the height for</param>
		/// <param name="ssDesiredHeight">The desired pixel height for the row</param>
		void MssRow_SetHeight(object ssWorksheet, int ssRowNumber, decimal ssDesiredHeight);

		/// <summary>
		/// Write a converted value to a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		/// <param name="ssCellValue">Text Value</param>
		/// <param name="ssCellType">Type can be:
		/// general (default if empty)
		/// text,
		/// datetime,
		/// integer,
		/// decimal,
		/// boolean,
		/// formula</param>
		void MssCell_WriteByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssCellValue, string ssCellType);

		/// <summary>
		/// Write a converted value to a cell, defined by its index.
		/// Accepts format for the target cell
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		/// <param name="ssCellValue">Text Value</param>
		/// <param name="ssCellType">Type can be:
		/// general (default if empty)
		/// text,
		/// datetime,
		/// integer,
		/// decimal,
		/// boolean,
		/// formula</param>
		/// <param name="ssCellFormat">CellFormat for the target cell</param>
		void MssCell_WriteByIndexWithFormat(object ssWorksheet, int ssRow, int ssColumn, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat);

		/// <summary>
		/// Write a converted value to a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet in which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssCellValue">Value to write</param>
		/// <param name="ssCellType">Type can be:
		/// general (default if empty)
		/// text,
		/// datetime,
		/// integer,
		/// decimal,
		/// boolean,
		/// formula</param>
		void MssCell_WriteByName(object ssWorksheet, string ssCellName, string ssCellValue, string ssCellType);

		/// <summary>
		/// Write a converted value to a cell, defined by its name.
		/// Accepts format for the target cell
		/// </summary>
		/// <param name="ssWorksheet">Worksheet in which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssCellValue">Value to write</param>
		/// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
		/// <param name="ssCellFormat">CellFormat for the target cell</param>
		void MssCell_WriteByNameWithFormat(object ssWorksheet, string ssCellName, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat);

		/// <summary>
		/// Write a dataset to a range of column cells
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssRow"></param>
		/// <param name="ssColumnStart"></param>
		/// <param name="ssValueList"></param>
		/// <param name="ssCellType"></param>
		void MssCell_WriteColumnRange(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType);

		/// <summary>
		/// Write a dataset to a range of column cells
		/// Accepts format for the target cells
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssRow">rownumber</param>
		/// <param name="ssColumnStart">Start column (integer)</param>
		/// <param name="ssValueList">Values to write to columns</param>
		/// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
		/// <param name="ssCellFormat">CellFormat for the target cells</param>
		void MssCell_WriteColumnRangeWithFormat(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType, RCCellFormatRecord ssCellFormat);

		/// <summary>
		/// Write a image on a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">row number</param>
		/// <param name="ssColumn">column number</param>
		/// <param name="ssImageName">The image name</param>
		/// <param name="ssImage">The image to write.</param>
		void MssCell_WriteImageByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssImageName, byte[] ssImage);

		/// <summary>
		/// Write a image on a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssImageName">The image name</param>
		/// <param name="ssImage">The image to write.</param>
		void MssCell_WriteImageByName(object ssWorksheet, string ssCellName, string ssImageName, byte[] ssImage);

		/// <summary>
		/// Write a dataset to a range of cells.
		/// Accepts format for the target cells
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssRowStart">Start row (integer)</param>
		/// <param name="ssColumnStart">Start column (integer)</param>
		/// <param name="ssDataSet">Data to write</param>
		/// <param name="ssCellFormat">CellFormat for the target cells</param>
		void MssCell_WriteRangeWithFormat(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord ssCellFormat);

		/// <summary>
		/// Add a worksheet to work on by its name
		/// </summary>
		/// <param name="ssWorkbook">Workbook where the sheet is to be added</param>
		/// <param name="ssWorksheetName">The name of the spreadsheet to create</param>
		/// <param name="ssWorksheet">The newly added worksheet</param>
		void MssWorkbook_AddName(object ssWorkbook, string ssWorksheetName, out object ssWorksheet);

		/// <summary>
		/// Set the active sheet
		/// </summary>
		/// <param name="ssWorkbook"></param>
		/// <param name="ssWorksheetName"></param>
		/// <param name="ssWorksheetIndex"></param>
		void MssWorksheet_SetActive(object ssWorkbook, string ssWorksheetName, int ssWorksheetIndex);

		/// <summary>
		/// Action is used to add Drop down list.
		/// </summary>
		/// <param name="ssWorksheet">Current worksheet the values to be added.</param>
		/// <param name="ssItemsList">Items values to be added to dropdown.</param>
		/// <param name="ssItemsAddress">Instead of using the itemslist to make a list of values, you can refer to a location within your Excel sheet for the list of values. Example: &quot;=B10:B20&quot; or &quot;=Sheet2!$C$1:$C$1000&quot;</param>
		/// <param name="ssCellRange">Sheet Cell range on which dropdown to be added, e.g. &quot;B:B&quot;</param>
		/// <param name="ssTitleMessage">Dropdown title message to be shown.</param>
		/// <param name="ssPromptMessage">Dropdown propmt message to be shown.</param>
		/// <param name="ssShowError">Show error when using invalid input on dropdown</param>
		/// <param name="ssCustomErrorMessage">Error to be shown when using invalid input on dropdown</param>
		/// <param name="ssCustomErrorTitle">Title of error popup to be shown when using invalid input on dropdown</param>
		void MssWorksheet_AddDropdown(object ssWorksheet, RLItemsRecordList ssItemsList, string ssItemsAddress, string ssCellRange, string ssTitleMessage, string ssPromptMessage, bool ssShowError, string ssCustomErrorMessage, string ssCustomErrorTitle);

		/// <summary>
		/// Set the footer on the specified worksheet.
		/// To insert fields use the following:
		/// Filename: &amp;F
		/// Sheet name: &amp;A
		/// Last saved date: &amp;D
		/// Last saved time: &amp;T
		/// Page number: &amp;P
		/// Number of pages: &amp;N
		/// To set the color of the following part of the section using &amp;K immediately followed by a hexadecimal RGB color
		/// To set the color to red for example use &amp;KFF0000
		/// eg Set the LeftSection to red and include the filename, set LeftSection to &quot;&amp;KFF0000Here is the filename &amp;F&quot;
		/// </summary>
		/// <param name="ssWorksheet">The worksheet for which the footer is to be set.</param>
		/// <param name="ssLeftSection">The content for the left section.</param>
		/// <param name="ssCenterSection">The content for the center section.</param>
		/// <param name="ssRightSection">The content for the right section.</param>
		void MssWorksheet_SetFooter(object ssWorksheet, string ssLeftSection, string ssCenterSection, string ssRightSection);

		/// <summary>
		/// Set the header of the specified worksheet.
		/// To insert fields use the following:
		/// Filename: &amp;F
		/// Sheet name: &amp;A
		/// Last saved date: &amp;D
		/// Last saved time: &amp;T
		/// Page number: &amp;P
		/// Number of pages: &amp;N
		/// To set the color of the following part of the section using &amp;K immediately followed by a hexadecimal RGB color
		/// To set the color to red for example use &amp;KFF0000
		/// eg Set the LeftSection to red and include the filename, set LeftSection to &quot;&amp;KFF0000Here is the filename &amp;F&quot;
		/// </summary>
		/// <param name="ssWorksheet">The worksheet for which the header is to be set.</param>
		/// <param name="ssLeftSection">The content for the left section.</param>
		/// <param name="ssCenterSection">The content for the center section.</param>
		/// <param name="ssRightSection">The content for the right section.</param>
		void MssWorksheet_SetHeader(object ssWorksheet, string ssLeftSection, string ssCenterSection, string ssRightSection);

		/// <summary>
		/// Get the left, center and right sections for the odd or even page header of the specified worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet from which to retrieve the header.</param>
		/// <param name="ssIsEven">If True, retrieves the even page header, otherwise the odd page header.</param>
		/// <param name="ssLeftSection">The left section of the header.</param>
		/// <param name="ssCenterSection">The center section of the header.</param>
		/// <param name="ssRightSection">The right section of the header.</param>
		void MssWorksheet_GetHeader(object ssWorksheet, bool ssIsEven, out string ssLeftSection, out string ssCenterSection, out string ssRightSection);

		/// <summary>
		/// Get the left, center and right sections for the odd or even page footer of the specified worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet for which to get the footer.</param>
		/// <param name="ssIsEven">If True, retrieves the even page footer, otherwise the odd page footer.</param>
		/// <param name="ssLeftSection">The left section of the footer.</param>
		/// <param name="ssCenterSection">The center section of the footer.</param>
		/// <param name="ssRightSection">The right section of the footer.</param>
		void MssWorksheet_GetFooter(object ssWorksheet, bool ssIsEven, out string ssLeftSection, out string ssCenterSection, out string ssRightSection);

		/// <summary>
		/// Clear value of a cell, defined by its index.
		/// Option to specify whether the cell is part of a merged group or not.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssStartColumn">Column Number</param>
		/// <param name="ssEndColumn">Column Number, Mandatory if IsMerged is True</param>
		/// <param name="ssIsMerged">If True, cells are merged and will be unmerged.</param>
		void MssCell_ClearValueByIndex(object ssWorksheet, int ssRow, int ssStartColumn, int ssEndColumn, bool ssIsMerged);

		/// <summary>
		/// Clear value clear the value of a specific cell by its name.
		/// Option to specify whether the cell is part of a merged group or not.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell name (eg A1:B1, if cells are merged; eg A1, if single cell)</param>
		/// <param name="ssIsMerged">If True cells are merged and will be unmerged.</param>
		void MssCell_ClearValueByName(object ssWorksheet, string ssCellName, bool ssIsMerged);

		/// <summary>
		/// Reads formula of a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		/// <param name="ssFormula">The formula</param>
		void MssCell_ReadFormulaByIndex(object ssWorksheet, int ssRow, int ssColumn, out string ssFormula);

		/// <summary>
		/// Input text address and get back the Row/Col values
		/// </summary>
		/// <param name="ssAddress">Text address, e.g. AB47 or A11:AB47</param>
		/// <param name="ssRowStart">Address row or range start row</param>
		/// <param name="ssColStart">Address col or range start column</param>
		/// <param name="ssRowEnd">Range end row</param>
		/// <param name="ssColEnd">Range end column</param>
		void MssAddress_From_Text(string ssAddress, out int ssRowStart, out int ssColStart, out int ssRowEnd, out int ssColEnd);

		/// <summary>
		/// Input Row/Col values and get the text address
		/// </summary>
		/// <param name="ssRowStart">Start row of the address</param>
		/// <param name="ssColStart">Start column of the address</param>
		/// <param name="ssRowEnd">End row of the address or zero</param>
		/// <param name="ssColEnd">End column of the address or zero</param>
		/// <param name="ssAddress">Text address, e.g. AB47 or C11:AB47</param>
		void MssAddress_From_RowCol(int ssRowStart, int ssColStart, int ssRowEnd, int ssColEnd, out string ssAddress);

		/// <summary>
		/// Get the Microsoft Office properties of the Excel document.
		/// </summary>
		/// <param name="ssWorkbook">The workbook</param>
		/// <param name="ssProperties">The Microsoft Office properties of the Excel document.</param>
		void MssWorkbook_GetProperties(object ssWorkbook, out RCOfficePropertiesRecord ssProperties);

		/// <summary>
		/// Set the Microsoft Office properties of the Excel document.
		/// </summary>
		/// <param name="ssWorkbook">The workbook</param>
		/// <param name="ssProperties">The Microsoft Office properties of the Excel document.</param>
		/// <param name="ssIgnoreBlank">If True, any blank properties in the Properties structure provided will be left with their existing values. If False, any blank properties in the Properties structure provided will be set to blank.</param>
		void MssWorkbook_SetProperties(object ssWorkbook, RCOfficePropertiesRecord ssProperties, bool ssIgnoreBlank);

		/// <summary>
		/// Clear all Microsoft Office properties of the Excel document. To only clear some properties, set the associated &quot;Clear&quot; attribute to True for the properties to clear, and the remaining ones false. The default behaviour is to clear all properties.
		/// </summary>
		/// <param name="ssWorkbook">The workbook.</param>
		/// <param name="ssClearTitle">If True, clears the Title property.</param>
		/// <param name="ssClearSubject">If True, clears the Subject property.</param>
		/// <param name="ssClearAuthor">If True, clears the Author property.</param>
		/// <param name="ssClearComments">If True, clears the Comments property.</param>
		/// <param name="ssClearKeywords">If True, clears the Keywords property.</param>
		/// <param name="ssClearLastModifiedBy">If True, clears the LastModifiedBy  property.</param>
		/// <param name="ssClearCategory">If True, clears the Category property.</param>
		/// <param name="ssClearStatus">If True, clears the Status property.</param>
		/// <param name="ssClearCompany">If True, clears the Company property.</param>
		/// <param name="ssClearManager">If True, clears the Manager property.</param>
		void MssWorkbook_ClearProperties(object ssWorkbook, bool ssClearTitle, bool ssClearSubject, bool ssClearAuthor, bool ssClearComments, bool ssClearKeywords, bool ssClearLastModifiedBy, bool ssClearCategory, bool ssClearStatus, bool ssClearCompany, bool ssClearManager);

	} // IssAdvanced_Excel

} // OutSystems.NssAdvanced_Excel
