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
		/// <param name="ssWorkbook">The newly created workbook</param>
		/// <param name="ssFirstSheetName">Specify the name of the initial sheet in the workbook. Default = &quot;Sheet1&quot;</param>
		void MssWorkbook_Create(out object ssWorkbook, string ssFirstSheetName);

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
		/// 
		/// </summary>
		/// <param name="ssWorksheet">Worksheet to protect</param>
		/// <param name="sspassword">Password to the protected spreadsheet</param>
		void MssWorksheet_Protect(object ssWorksheet, string sspassword);

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
		/// 
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssWorksheetName"></param>
		void MssWorksheet_GetName(object ssWorksheet, out string ssWorksheetName);

		/// <summary>
		/// Get all properties of the workbook
		/// </summary>
		/// <param name="ssWorkbook">The workbook</param>
		/// <param name="ssProperties"></param>
		void MssWorkbook_GetProperties(object ssWorkbook, out RCWorkbookRecord ssProperties);

		/// <summary>
		/// 
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
		/// 
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

	} // IssAdvanced_Excel

} // OutSystems.NssAdvanced_Excel
