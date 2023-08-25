using System;
using System.Collections;
using System.Data;
using OutSystems.HubEdition.RuntimePlatform;
using OutSystems.RuntimePublic.Db;

namespace OutSystems.NssAdvanced_Excel {

	public class CssAdvanced_Excel: IssAdvanced_Excel {

		/// <summary>
		/// Opens an existing workbook for editing by either specifying a name or the binary data.
		/// </summary>
		/// <param name="ssFileName">Location of the file that you want to open. Set to empty string &quot;&quot; when using binary data</param>
		/// <param name="ssBinary_Data">Binary data of the file that you want to open. Set to nullbinary() if using FileName</param>
		/// <param name="ssWorkbook">The workbook that you want to work with.</param>
		public void MssWorkbook_Open(string ssFileName, byte[] ssBinary_Data, out object ssWorkbook) {
			ssWorkbook = null;
			// TODO: Write implementation for action
		} // MssWorkbook_Open

		/// <summary>
		/// Select a worksheet by its index or by its name
		/// </summary>
		/// <param name="ssWorkbook">The workbook wherein the worksheet exists</param>
		/// <param name="ssWorksheetIndex">The index of the worksheet to find. Indexes start at 1</param>
		/// <param name="ssWorksheetName">The name of the worksheet to find</param>
		/// <param name="ssWorksheet">This is the worksheet object that you have been looking for,</param>
		public void MssWorksheet_Select(object ssWorkbook, int ssWorksheetIndex, string ssWorksheetName, out object ssWorksheet) {
			ssWorksheet = null;
			// TODO: Write implementation for action
		} // MssWorksheet_Select

		/// <summary>
		/// Creates a new excel workbook, optionally specifying the name of the fiirst sheet.
		/// </summary>
		/// <param name="ssNumberOfSheets">The number of sheets to add. Sheet names will be auto generated, i.e. Sheet1, Sheet2.</param>
		/// <param name="ssFirstSheetName">Specify the name of the initial sheet in the workbook. Default = &quot;Sheet1&quot;</param>
		/// <param name="ssSheetNames">List of new sheets to add, with at least a name specified. The index, if specified, will be used to add sheets in that order.
		/// FirstSheetName and NrSheets are ignored if SheetNames is populated</param>
		/// <param name="ssWorkbook">The newly created workbook</param>
		public void MssWorkbook_Create(int ssNumberOfSheets, string ssFirstSheetName, RLNewSheetRecordList ssSheetNames, out object ssWorkbook) {
			ssWorkbook = null;
			// TODO: Write implementation for action
		} // MssWorkbook_Create

		/// <summary>
		/// Get the in-memory binary data of the specified workbook
		/// </summary>
		/// <param name="ssWorkbook">The workbook you want the binary data for</param>
		/// <param name="ssBinaryData">The binary data of the file</param>
		public void MssWorkbook_GetBinaryData(object ssWorkbook, out byte[] ssBinaryData) {
			ssBinaryData = new byte[] {};
			// TODO: Write implementation for action
		} // MssWorkbook_GetBinaryData

		/// <summary>
		/// Rename a worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssName">The new name for the spreadsheet</param>
		public void MssWorksheet_Rename(object ssWorksheet, string ssName) {
			// TODO: Write implementation for action
		} // MssWorksheet_Rename

		/// <summary>
		/// Closes the excel workbook
		/// </summary>
		/// <param name="ssWorkbook"></param>
		public void MssWorkbook_Close(object ssWorkbook) {
			// TODO: Write implementation for action
		} // MssWorkbook_Close

		/// <summary>
		/// Hides / Shows a Column passed by index
		/// </summary>
		/// <param name="ssWorksheet">The worksheet you want to work with.</param>
		/// <param name="ssColumn">The index of the column within the worksheet that you want to hide/show.</param>
		/// <param name="ssHidden">A Boolean value, set to True to hide the column, and to False to show the column.</param>
		public void MssColumn_Hide_Show(object ssWorksheet, int ssColumn, bool ssHidden) {
			// TODO: Write implementation for action
		} // MssColumn_Hide_Show

		/// <summary>
		/// Reads the value of a cell.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Name of the cell to read from, i.e. A4. Required if CellRow and CellNumber set to 0.</param>
		/// <param name="ssCellRow">Row number of the cell to read from. Required if CellName not set.</param>
		/// <param name="ssCellColumn">Column number of the cell to read from. Required if CellName not set.</param>
		/// <param name="ssCellValue">The value in the cell, as text.</param>
		/// <param name="ssReadText">If true always reads the cell value as text</param>
		public void MssCell_Read(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, out string ssCellValue, bool ssReadText) {
			ssCellValue = "";
			// TODO: Write implementation for action
		} // MssCell_Read

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
		public void MssWorksheet_Protect(object ssWorksheet, string ssPassword, RCProtectionRecord ssProtectionOptions) {
			// TODO: Write implementation for action
		} // MssWorksheet_Protect

		/// <summary>
		/// Write a converted value to a cell.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides </param>
		/// <param name="ssCellName">Name of the cell to write to, i.e. A4. Required if CellRow and CellColumn not set</param>
		/// <param name="ssCellRow">Row number of the cell to write to. Required if CellName not set.</param>
		/// <param name="ssCellColumn">Column number of the cell to write to. Required if CellName not set.</param>
		/// <param name="ssCellValue">The value to write to the cell</param>
		/// <param name="ssCellType">Type can be:
		/// text (default),
		/// datetime,
		/// integer,
		/// decimal,
		/// boolean,
		/// formula</param>
		/// <param name="ssCellFormat">CellFormat for the target cell</param>
		public void MssCell_Write(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat) {
			// TODO: Write implementation for action
		} // MssCell_Write

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
		public void MssCell_WriteRange(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord ssCellFormat, bool ssExportHeaders) {
			// TODO: Write implementation for action
		} // MssCell_WriteRange

		/// <summary>
		/// Get the name of the given worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssWorksheetName"></param>
		public void MssWorksheet_GetName(object ssWorksheet, out string ssWorksheetName) {
			ssWorksheetName = "";
			// TODO: Write implementation for action
		} // MssWorksheet_GetName

		/// <summary>
		/// Get all properties of the workbook
		/// </summary>
		/// <param name="ssWorkbook">The workbook</param>
		/// <param name="ssProperties"></param>
		public void MssWorkbook_GetProperties(object ssWorkbook, out RCWorkbookRecord ssProperties) {
			ssProperties = new RCWorkbookRecord(null);
			// TODO: Write implementation for action
		} // MssWorkbook_GetProperties

		/// <summary>
		/// Get the properties of the given worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssProperties"></param>
		public void MssWorksheet_GetProperties(object ssWorksheet, out RCWorksheetRecord ssProperties) {
			ssProperties = new RCWorksheetRecord(null);
			// TODO: Write implementation for action
		} // MssWorksheet_GetProperties

		/// <summary>
		/// Add a worksheet to an existing workbook, optionally at the index specified. Specifying only a name will create a blank sheet. Specifying  a name with binary data, will add the sheet from the existing binary data, and then rename to the newly provided name
		/// </summary>
		/// <param name="ssWorkbook">The workbook that you want to add the sheet to</param>
		/// <param name="ssWorksheetName">The name of the worksheet you want to add. If binary data is nullbinary(), an empty sheet will be added</param>
		/// <param name="ssWorksheet">The worksheet object that you want to add. Set to nullbinary() if adding a new sheet by name</param>
		/// <param name="ssIndexWhereToAdd">The index where to add the new sheet. Default will be highest sheet index plus 1</param>
		public void MssWorkBook_AddSheet(object ssWorkbook, string ssWorksheetName, object ssWorksheet, int ssIndexWhereToAdd) {
			// TODO: Write implementation for action
		} // MssWorkBook_AddSheet

		/// <summary>
		/// Delete a worksheet in a workbook by specifying either the index, or the name of the worksheet.
		/// </summary>
		/// <param name="ssWorkbook">The workbook from which you want to delete the worksheet</param>
		/// <param name="ssIndexToDelete">The index of the worksheet to delete. Set to 0 if using the worksheet name to delete</param>
		/// <param name="ssNameToDelete">The name of the worksheet to delete. Set to empty string &quot;&quot; if using the index to delete.</param>
		public void MssWorksheet_Delete(object ssWorkbook, int ssIndexToDelete, string ssNameToDelete) {
			// TODO: Write implementation for action
		} // MssWorksheet_Delete

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
		public void MssChart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos) {
			// TODO: Write implementation for action
		} // MssChart_Create

		/// <summary>
		/// Inserts a new row into the spreadsheet.  Existing rows below the position are shifted down.  All formula are updated to take account of the new row.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to insert the row(s) into</param>
		/// <param name="ssInsertAt">The position of the new row
		/// </param>
		/// <param name="ssNrRows">Number of rows to insert</param>
		/// <param name="ssCopyStyleFromRow">Copy Styles from this row. Applied to all inserted rows. 0 will not copy any styles</param>
		public void MssRow_Insert(object ssWorksheet, int ssInsertAt, int ssNrRows, int ssCopyStyleFromRow) {
			// TODO: Write implementation for action
		} // MssRow_Insert

		/// <summary>
		/// Change the index of a worksheet in the document
		/// </summary>
		/// <param name="ssWorkbook">The workbook in which the change is to be made.</param>
		/// <param name="ssCurrentIndex">The current index(position) of the sheet in question</param>
		/// <param name="ssNewIndex">The new index for the sheet</param>
		public void MssWorkbook_ChangeSheetIndex(object ssWorkbook, int ssCurrentIndex, int ssNewIndex) {
			// TODO: Write implementation for action
		} // MssWorkbook_ChangeSheetIndex

		/// <summary>
		/// Apply a specified cell format to the range specified for the given worksheet
		/// </summary>
		/// <param name="ssWorksheet">Worksheet object where formatting is to be applied</param>
		/// <param name="ssCellFormat">CellFormat to apply</param>
		/// <param name="ssRange">Range that CellFormat is to be applied to</param>
		public void MssCellFormat_ApplyToRange(object ssWorksheet, RCCellFormatRecord ssCellFormat, RCRangeRecord ssRange) {
			// TODO: Write implementation for action
		} // MssCellFormat_ApplyToRange

		/// <summary>
		/// Find all cells that contain the specified value in the given worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet in which to search</param>
		/// <param name="ssValueToFind">The value to search for</param>
		/// <param name="ssListOfCells">List of cells (ranges) where the value has been found</param>
		public void MssCells_FindByValue(object ssWorksheet, string ssValueToFind, out RLRangeRecordList ssListOfCells) {
			ssListOfCells = new RLRangeRecordList();
			// TODO: Write implementation for action
		} // MssCells_FindByValue

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
		public void MssContainInRange(object ssWorksheet, string ssRange, string ssValue, string ssParameter1, out bool ssFound, out int ssRowIndex, out int ssColumnIndex) {
			ssFound = false;
			ssRowIndex = 0;
			ssColumnIndex = 0;
			// TODO: Write implementation for action
		} // MssContainInRange

		/// <summary>
		/// Calculate all formulae for the entire workbook provided.
		/// </summary>
		/// <param name="ssWorkbook">The workbook to work with</param>
		public void MssWorkbook_Calculate(object ssWorkbook) {
			// TODO: Write implementation for action
		} // MssWorkbook_Calculate

		/// <summary>
		/// Calculate all formulae on the provided worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		public void MssWorksheet_Calculate(object ssWorksheet) {
			// TODO: Write implementation for action
		} // MssWorksheet_Calculate

		/// <summary>
		/// Hides / Shows Row passed by index
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to work with</param>
		/// <param name="ssRowIndex">Index of the Row to show/hide</param>
		/// <param name="ssHidden">A Boolean value, set to True to hide the row and to False to show the row</param>
		public void MssRow_Hide_Show(object ssWorksheet, int ssRowIndex, bool ssHidden) {
			// TODO: Write implementation for action
		} // MssRow_Hide_Show

		/// <summary>
		/// Hide / Show a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssHidden">Visible = 0 - The worksheet is visible
		/// Hidden = 1 - The worksheet is hidden but can be shown by the user via the user interface
		/// VeryHidden = 2 - The worksheet is hidden and cannot be shown by the user via the user interface</param>
		public void MssWorksheet_Hide_Show(object ssWorksheet, int ssHidden) {
			// TODO: Write implementation for action
		} // MssWorksheet_Hide_Show

		/// <summary>
		/// Add a rule for conditionally formatting a range of cells.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssConditionalFormatRecord">The conditional formatting to apply to the Address Range</param>
		public void MssConditionalFormatting_AddRule(object ssWorksheet, RCConditionalFormatItemRecord ssConditionalFormatRecord) {
			// TODO: Write implementation for action
		} // MssConditionalFormatting_AddRule

		/// <summary>
		/// Get a list of all the conditional formatting rules in a worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssListOfRule">List of conditional formatting rules</param>
		public void MssConditionalFormatting_GetAllRules(object ssWorksheet, out RLConditionalFormatItemRecordList ssListOfRule) {
			ssListOfRule = new RLConditionalFormatItemRecordList();
			// TODO: Write implementation for action
		} // MssConditionalFormatting_GetAllRules

		/// <summary>
		/// Merge cells in the range provided
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssRangeToMerge">The range of the cells to merge</param>
		public void MssCell_Merge(object ssWorksheet, RCRangeRecord ssRangeToMerge) {
			// TODO: Write implementation for action
		} // MssCell_Merge

		/// <summary>
		/// Un-Merge cells in the range provided
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssRangeToUnmerge">The range of cell to un-merge</param>
		public void MssCell_UnMerge(object ssWorksheet, RCRangeRecord ssRangeToUnmerge) {
			// TODO: Write implementation for action
		} // MssCell_UnMerge

		/// <summary>
		/// Delete row(s) from a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssStartRowNumber">Row number where to start deleting rows.</param>
		/// <param name="ssNumberOfRows">The number of rows to delete. Default = 1.</param>
		public void MssRow_Delete(object ssWorksheet, int ssStartRowNumber, int ssNumberOfRows) {
			// TODO: Write implementation for action
		} // MssRow_Delete

		/// <summary>
		/// Delete column(s) from a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssStartColumnNumber">Column number where to start deleting columns.</param>
		/// <param name="ssNumberOfColumns">The number of rows to delete. Default = 1.</param>
		public void MssColumn_Delete(object ssWorksheet, int ssStartColumnNumber, int ssNumberOfColumns) {
			// TODO: Write implementation for action
		} // MssColumn_Delete

		/// <summary>
		/// Delete comment(s) in a specified range
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRange">Range to delete comments from.</param>
		public void MssComment_Delete(object ssWorksheet, RCRangeRecord ssRange) {
			// TODO: Write implementation for action
		} // MssComment_Delete

		/// <summary>
		/// Add a comment to a cell
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRowNumber">The row number of the cell to add the comment to.</param>
		/// <param name="ssColumnNumber">The column number of the cell to add the comment to.</param>
		/// <param name="ssText">The comment.</param>
		/// <param name="ssAuthor">The author of the comment.</param>
		/// <param name="ssAutofit">True to autofit the comment window to the comment text</param>
		/// <param name="ssIsRichText">Is the comment rich text</param>
		public void MssComment_Add(object ssWorksheet, int ssRowNumber, int ssColumnNumber, string ssText, string ssAuthor, bool ssAutofit, bool ssIsRichText) {
			// TODO: Write implementation for action
		} // MssComment_Add

		/// <summary>
		/// Inserts a new column into the spreadsheet.  Existing columns to the right of the insert index will be shifted right.  All formula are updated to take account of the new column.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssInsertAt">Column number where to insert new column.</param>
		/// <param name="ssNumberOfColumns">The number of columns to insert.</param>
		/// <param name="ssCopyStylesFrom">Copy Styles from this column. Applied to all inserted columns. 0 (default) will not copy any styles</param>
		public void MssColumn_Insert(object ssWorksheet, int ssInsertAt, int ssNumberOfColumns, int ssCopyStylesFrom) {
			// TODO: Write implementation for action
		} // MssColumn_Insert

		/// <summary>
		/// Delete a specified Conditional Formatting rule on a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRuleToDeleteIndex">The index of the rule to be deleted.</param>
		public void MssConditionalFormatting_DeleteRule(object ssWorksheet, int ssRuleToDeleteIndex) {
			// TODO: Write implementation for action
		} // MssConditionalFormatting_DeleteRule

		/// <summary>
		/// Delete ALL Conditional Formatting rules for a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		public void MssConditionalFormatting_DeleteAllRules(object ssWorksheet) {
			// TODO: Write implementation for action
		} // MssConditionalFormatting_DeleteAllRules

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
		/// <param name="ssMarginTop"> Offset in pixels	</param>
		/// <param name="ssMarginLeft"> Offset in pixels</param>
		public void MssImage_Insert(object ssWorksheet, byte[] ssImageFile, string ssImageType, string ssImageName, int ssRowNumber, int ssColumnNumber, string ssCellName, int ssImageWidth, int ssImageHeight, int ssMarginTop, int ssMarginLeft) {
			// TODO: Write implementation for action
		} // MssImage_Insert

		/// <summary>
		/// Apply the column autofit action to the specified range of cells specified in the given worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		public void MssWorksheet_AutofitColumns(object ssWorksheet) {
			// TODO: Write implementation for action
		} // MssWorksheet_AutofitColumns

		/// <summary>
		/// Add the automatic filter option of Excel to the specified range of cells.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRangeToFilter">The range where to add the filter. If not supplied, the dimension of the worksheet will be used.</param>
		public void MssWorksheet_AddAutoFilter(object ssWorksheet, RCRangeRecord ssRangeToFilter) {
			// TODO: Write implementation for action
		} // MssWorksheet_AddAutoFilter

		/// <summary>
		/// Set protection on the workbook level
		/// </summary>
		/// <param name="ssWorkbook">The workbook to work with</param>
		/// <param name="ssPassword">The password to set for the workbook. This does not encrypt the workbook.</param>
		/// <param name="ssLockStructure">Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.</param>
		/// <param name="ssLockWindows">Locks the position of the workbook window.</param>
		/// <param name="ssLockRevision">Lock the workbook for revision</param>
		public void MssWorkbook_Protect(object ssWorkbook, string ssPassword, bool ssLockStructure, bool ssLockWindows, bool ssLockRevision) {
			// TODO: Write implementation for action
		} // MssWorkbook_Protect

		/// <summary>
		/// Calculates the formula of a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		public void MssCell_CalculateByIndex(object ssWorksheet, int ssRow, int ssColumn) {
			// TODO: Write implementation for action
		} // MssCell_CalculateByIndex

		/// <summary>
		/// Calculates the formula of a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		public void MssCell_CalculateByName(object ssWorksheet, string ssCellName) {
			// TODO: Write implementation for action
		} // MssCell_CalculateByName

		/// <summary>
		/// Apply format to a range of cells.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssRowStart">Start row (integer)</param>
		/// <param name="ssColumnStart">Start column (integer)</param>
		/// <param name="ssRowEnd">End row (integer)</param>
		/// <param name="ssColumnEnd">End column (integer)</param>
		/// <param name="ssCellFormat">CellFormat for the target cells</param>
		public void MssCell_FormatRange(object ssWorksheet, int ssRowStart, int ssColumnStart, int ssRowEnd, int ssColumnEnd, RCCellFormatRecord ssCellFormat) {
			// TODO: Write implementation for action
		} // MssCell_FormatRange

		/// <summary>
		/// Reads the value of a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">row number</param>
		/// <param name="ssColumn">column number</param>
		/// <param name="ssReadText">If true always reads the cell value as text</param>
		/// <param name="ssCellValue">text-value</param>
		public void MssCell_ReadByIndex(object ssWorksheet, int ssRow, int ssColumn, bool ssReadText, out string ssCellValue) {
			ssCellValue = "";
			// TODO: Write implementation for action
		} // MssCell_ReadByIndex

		/// <summary>
		/// Reads the value of a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssReadText">If true always reads the cell value as text</param>
		/// <param name="ssCellValue">text-value</param>
		public void MssCell_ReadByName(object ssWorksheet, string ssCellName, bool ssReadText, out string ssCellValue) {
			ssCellValue = "";
			// TODO: Write implementation for action
		} // MssCell_ReadByName

		/// <summary>
		/// Write a formula to a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">rownumber</param>
		/// <param name="ssColumn">columnnumber</param>
		/// <param name="ssFormula">Formula</param>
		public void MssCell_SetFormulaByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssFormula) {
			// TODO: Write implementation for action
		} // MssCell_SetFormulaByIndex

		/// <summary>
		/// Write a formula to a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssFormula">Formula</param>
		public void MssCell_SetFormulaByName(object ssWorksheet, string ssCellName, string ssFormula) {
			// TODO: Write implementation for action
		} // MssCell_SetFormulaByName

		/// <summary>
		/// Adds a copy of a worksheet
		/// </summary>
		/// <param name="ssWorkbook">The workbook in which the worksheet is to be copied
		/// </param>
		/// <param name="ssWorksheetName">The name of the spreadsheet to create</param>
		/// <param name="ssWorksheetToCopy">The worksheet to be copied</param>
		/// <param name="ssWorksheet">The copied worksheet</param>
		public void MssWorkbook_AddCopyWorksheet(object ssWorkbook, string ssWorksheetName, object ssWorksheetToCopy, out object ssWorksheet) {
			ssWorksheet = null;
			// TODO: Write implementation for action
		} // MssWorkbook_AddCopyWorksheet

		/// <summary>
		/// Get all images in a worksheet
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssImages"></param>
		public void MssWorksheet_GetImages(object ssWorksheet, out RLImageRecordList ssImages) {
			ssImages = new RLImageRecordList();
			// TODO: Write implementation for action
		} // MssWorksheet_GetImages

		/// <summary>
		/// Select a worksheet by its index
		/// </summary>
		/// <param name="ssWorkbook">The worksheet to work with</param>
		/// <param name="ssWorksheetNumber">The index of the spreadsheet to select, starting at 1</param>
		/// <param name="ssWorksheet">The selected worksheet</param>
		public void MssWorksheet_SelectByIndex(object ssWorkbook, int ssWorksheetNumber, out object ssWorksheet) {
			ssWorksheet = null;
			// TODO: Write implementation for action
		} // MssWorksheet_SelectByIndex

		/// <summary>
		/// Select a worksheet to work on by its name
		/// </summary>
		/// <param name="ssWorkbook">The workbook to work with</param>
		/// <param name="ssWorksheetName">The name of the spreadsheet to select</param>
		/// <param name="ssWorksheet">The selected worksheet</param>
		public void MssWorksheet_SelectByName(object ssWorkbook, string ssWorksheetName, out object ssWorksheet) {
			ssWorksheet = null;
			// TODO: Write implementation for action
		} // MssWorksheet_SelectByName

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ssWorkbook"></param>
		/// <param name="ssNameToDelete"></param>
		public void MssWorksheet_DeleteByName(object ssWorkbook, string ssNameToDelete) {
			// TODO: Write implementation for action
		} // MssWorksheet_DeleteByName

		/// <summary>
		/// 
		/// </summary>
		/// <param name="ssWorkbook"></param>
		/// <param name="ssIndexToDelete"></param>
		public void MssWorksheet_DeleteByIndex(object ssWorkbook, int ssIndexToDelete) {
			// TODO: Write implementation for action
		} // MssWorksheet_DeleteByIndex

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
		public void MssWorksheet_Chart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos) {
			// TODO: Write implementation for action
		} // MssWorksheet_Chart_Create

		/// <summary>
		/// Create a defined &quot;Name&quot; (a word or string of characters in Excel that represents a cell, range of cells, formula, or constant value) in excel, starting in the RowStart / ColumnStart cell.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssName">&quot;Name&quot;</param>
		/// <param name="ssDataSet">Values to assigned the name</param>
		/// <param name="ssRowStart">Start row number</param>
		/// <param name="ssColumnStart">Start column number</param>
		public void MssWorksheet_AddName(object ssWorksheet, string ssName, object ssDataSet, int ssRowStart, int ssColumnStart) {
			// TODO: Write implementation for action
		} // MssWorksheet_AddName

		/// <summary>
		/// Opens an existing workbook for editing and keeps it in memory
		/// </summary>
		/// <param name="ssBinaryData"></param>
		/// <param name="ssWorkbook"></param>
		public void MssWorkbook_Open_BinaryData(byte[] ssBinaryData, out object ssWorkbook) {
			ssWorkbook = null;
			// TODO: Write implementation for action
		} // MssWorkbook_Open_BinaryData

		/// <summary>
		/// Set the pixel width of a column on a specific worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssColumnNumber">The column number, starting at 1</param>
		/// <param name="ssDesiredWidth">The pixel width you desire for the column.</param>
		public void MssColumn_SetWidth(object ssWorksheet, int ssColumnNumber, decimal ssDesiredWidth) {
			// TODO: Write implementation for action
		} // MssColumn_SetWidth

		/// <summary>
		/// Set the pixel height for a specific row in a worksheet
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with
		/// </param>
		/// <param name="ssRowNumber">The number of the row to set the height for</param>
		/// <param name="ssDesiredHeight">The desired pixel height for the row</param>
		public void MssRow_SetHeight(object ssWorksheet, int ssRowNumber, decimal ssDesiredHeight) {
			// TODO: Write implementation for action
		} // MssRow_SetHeight

		/// <summary>
		/// Write a converted value to a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		/// <param name="ssCellValue">Text Value</param>
		/// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
		public void MssCell_WriteByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssCellValue, string ssCellType) {
			// TODO: Write implementation for action
		} // MssCell_WriteByIndex

		/// <summary>
		/// Write a converted value to a cell, defined by its index.
		/// Accepts format for the target cell
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">Row Number</param>
		/// <param name="ssColumn">Column Number</param>
		/// <param name="ssCellValue">Text Value</param>
		/// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
		/// <param name="ssCellFormat">CellFormat for the target cell</param>
		public void MssCell_WriteByIndexWithFormat(object ssWorksheet, int ssRow, int ssColumn, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat) {
			// TODO: Write implementation for action
		} // MssCell_WriteByIndexWithFormat

		/// <summary>
		/// Write a converted value to a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet in which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssCellValue">Value to write</param>
		/// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
		public void MssCell_WriteByName(object ssWorksheet, string ssCellName, string ssCellValue, string ssCellType) {
			// TODO: Write implementation for action
		} // MssCell_WriteByName

		/// <summary>
		/// Write a converted value to a cell, defined by its name.
		/// Accepts format for the target cell
		/// </summary>
		/// <param name="ssWorksheet">Worksheet in which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssCellValue">Value to write</param>
		/// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
		/// <param name="ssCellFormat">CellFormat for the target cell</param>
		public void MssCell_WriteByNameWithFormat(object ssWorksheet, string ssCellName, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat) {
			// TODO: Write implementation for action
		} // MssCell_WriteByNameWithFormat

		/// <summary>
		/// Write a dataset to a range of column cells
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssRow"></param>
		/// <param name="ssColumnStart"></param>
		/// <param name="ssValueList"></param>
		/// <param name="ssCellType"></param>
		public void MssCell_WriteColumnRange(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType) {
			// TODO: Write implementation for action
		} // MssCell_WriteColumnRange

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
		public void MssCell_WriteColumnRangeWithFormat(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType, RCCellFormatRecord ssCellFormat) {
			// TODO: Write implementation for action
		} // MssCell_WriteColumnRangeWithFormat

		/// <summary>
		/// Write a image on a cell, defined by its index.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssRow">row number</param>
		/// <param name="ssColumn">column number</param>
		/// <param name="ssImageName">The image name</param>
		/// <param name="ssImage">The image to write.</param>
		public void MssCell_WriteImageByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssImageName, byte[] ssImage) {
			// TODO: Write implementation for action
		} // MssCell_WriteImageByIndex

		/// <summary>
		/// Write a image on a cell, defined by its name.
		/// </summary>
		/// <param name="ssWorksheet">Worksheet on which the cell resides</param>
		/// <param name="ssCellName">Cell-name (eg A4)</param>
		/// <param name="ssImageName">The image name</param>
		/// <param name="ssImage">The image to write.</param>
		public void MssCell_WriteImageByName(object ssWorksheet, string ssCellName, string ssImageName, byte[] ssImage) {
			// TODO: Write implementation for action
		} // MssCell_WriteImageByName

		/// <summary>
		/// Write a dataset to a range of cells.
		/// Accepts format for the target cells
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to write to</param>
		/// <param name="ssRowStart">Start row (integer)</param>
		/// <param name="ssColumnStart">Start column (integer)</param>
		/// <param name="ssDataSet">Data to write</param>
		/// <param name="ssCellFormat">CellFormat for the target cells</param>
		public void MssCell_WriteRangeWithFormat(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord ssCellFormat) {
			// TODO: Write implementation for action
		} // MssCell_WriteRangeWithFormat

		/// <summary>
		/// Add a worksheet to work on by its name
		/// </summary>
		/// <param name="ssWorkbook">Workbook where the sheet is to be added</param>
		/// <param name="ssWorksheetName">The name of the spreadsheet to create</param>
		/// <param name="ssWorksheet">The newly added worksheet</param>
		public void MssWorkbook_AddName(object ssWorkbook, string ssWorksheetName, out object ssWorksheet) {
			ssWorksheet = null;
			// TODO: Write implementation for action
		} // MssWorkbook_AddName

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
		/// <param name="ssLeftSection">The content for the left section</param>
		/// <param name="ssCenterSection">The content for the center section</param>
		/// <param name="ssRightSection">The content for the right section</param>
		public void MssWorksheet_SetHeader(object ssWorksheet, string ssLeftSection, string ssCenterSection, string ssRightSection) {
			// TODO: Write implementation for action
		} // MssWorksheet_SetHeader

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
		/// <param name="ssCenterSection">The content for the center section</param>
		/// <param name="ssRightSection">The content for the right section</param>
		public void MssWorksheet_SetFooter(object ssWorksheet, string ssLeftSection, string ssCenterSection, string ssRightSection) {
			// TODO: Write implementation for action
		} // MssWorksheet_SetFooter

		/// <summary>
		/// Get the left, center and right sections for the odd or even page header of the specified worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet from which to retrieve the header</param>
		/// <param name="ssIsEven">If True, retrieves the even page header, otherwise the odd page header.</param>
		/// <param name="ssLeftSection">The left section of the header.</param>
		/// <param name="ssCenterSection">The center section of the header.</param>
		/// <param name="ssRightSection">The right section of the header.</param>
		public void MssWorksheet_GetHeader(object ssWorksheet, bool ssIsEven, out string ssLeftSection, out string ssCenterSection, out string ssRightSection) {
			ssLeftSection = "";
			ssCenterSection = "";
			ssRightSection = "";
			// TODO: Write implementation for action
		} // MssWorksheet_GetHeader

		/// <summary>
		/// Get the left, center and right sections for the odd or even page footer of the specified worksheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet for which to get the footer.</param>
		/// <param name="ssIsEven">If True, retrieves the even page footer, otherwise the odd page footer.</param>
		/// <param name="ssLeftSection">The left section of the footer.</param>
		/// <param name="ssCenterSection">The center section of the footer.</param>
		/// <param name="ssRightSection">The right section of the footer.</param>
		public void MssWorksheet_GetFooter(object ssWorksheet, bool ssIsEven, out string ssLeftSection, out string ssCenterSection, out string ssRightSection) {
			ssLeftSection = "";
			ssCenterSection = "";
			ssRightSection = "";
			// TODO: Write implementation for action
		} // MssWorksheet_GetFooter

	} // CssAdvanced_Excel

} // OutSystems.NssAdvanced_Excel

