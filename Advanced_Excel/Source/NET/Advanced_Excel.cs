using System;
using System.Data;
using System.IO;
using System.Net;
using System.Drawing;
using OfficeOpenXml;
using OutSystems.HubEdition.RuntimePlatform;
using OutSystems.HubEdition.RuntimePlatform.Db;
using OutSystems.RuntimeCommon;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Drawing;

namespace OutSystems.NssAdvanced_Excel
{

    public class CssAdvanced_Excel : IssAdvanced_Excel
    {

		/// <summary>
		/// Define the cell range that will be printed for this worksheet (the print area saved in the file, used by Excel when printing). Pass an empty range to clear it and print the whole sheet.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRange">Print area in A1 notation, e.g. A1:H50. Empty clears it.</param>
		public void MssWorksheet_SetPrintArea(object ssWorksheet, string ssRange) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			if (string.IsNullOrEmpty(ssRange))
				ws.PrinterSettings.PrintArea = null;
			else
				ws.PrinterSettings.PrintArea = ws.Cells[ssRange];
		} // MssWorksheet_SetPrintArea

		/// <summary>
		/// Configure how the worksheet prints when opened in Excel: page orientation, paper size, and optional fit-to-page scaling (fit to a number of pages wide/tall).
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssOrientation">Portrait or Landscape.</param>
		/// <param name="ssPaperSize">Paper-size code; 0 = leave unchanged. Common: 9 = A4, 1 = Letter, 8 = A3, 5 = Legal.</param>
		/// <param name="ssFitToPage">Scale the sheet to fit the page(s).</param>
		/// <param name="ssFitToWidth">Pages wide (when FitToPage).</param>
		/// <param name="ssFitToHeight">Pages tall (0 = automatic).</param>
		public void MssWorksheet_SetPageLayout(object ssWorksheet, string ssOrientation, int ssPaperSize, bool ssFitToPage, int ssFitToWidth, int ssFitToHeight) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			ws.PrinterSettings.Orientation = ParseEnum(ssOrientation, eOrientation.Portrait);
			if (ssPaperSize > 0)
				ws.PrinterSettings.PaperSize = (ePaperSize)ssPaperSize;
			ws.PrinterSettings.FitToPage = ssFitToPage;
			if (ssFitToPage)
			{
				if (ssFitToWidth > 0) ws.PrinterSettings.FitToWidth = ssFitToWidth;
				ws.PrinterSettings.FitToHeight = ssFitToHeight;
			}
		} // MssWorksheet_SetPageLayout

		/// <summary>
		/// Set the rows and/or columns that repeat on every printed page (e.g. a header row that should appear at the top of each page).
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRepeatRows">Rows repeated on every printed page, e.g. 1:1 or 1:2. Empty = none.</param>
		/// <param name="ssRepeatColumns">Columns repeated on every page, e.g. A:A. Empty = none.</param>
		public void MssWorksheet_SetPrintTitles(object ssWorksheet, string ssRepeatRows, string ssRepeatColumns) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			if (!string.IsNullOrEmpty(ssRepeatRows))
				ws.PrinterSettings.RepeatRows = new ExcelAddress(ssRepeatRows);
			if (!string.IsNullOrEmpty(ssRepeatColumns))
				ws.PrinterSettings.RepeatColumns = new ExcelAddress(ssRepeatColumns);
		} // MssWorksheet_SetPrintTitles

		/// <summary>
		/// Set the page margins (top, bottom, left, right, header, footer) in inches, used by Excel when the worksheet is printed.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssTop">Margin between the top edge of the page and the start of the content, in inches.</param>
		/// <param name="ssBottom">Margin between the bottom edge of the page and the end of the content, in inches.</param>
		/// <param name="ssLeft">Margin between the left edge of the page and the content, in inches.</param>
		/// <param name="ssRight">Margin between the right edge of the page and the content, in inches.</param>
		/// <param name="ssHeader">Distance from the top edge of the page to the page header, in inches. Should be smaller than the Top margin so the header sits above the content.</param>
		/// <param name="ssFooter">Distance from the bottom edge of the page to the page footer, in inches. Should be smaller than the Bottom margin so the footer sits below the content.</param>
		public void MssWorksheet_SetMargins(object ssWorksheet, decimal ssTop, decimal ssBottom, decimal ssLeft, decimal ssRight, decimal ssHeader, decimal ssFooter) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			ws.PrinterSettings.TopMargin = ssTop;
			ws.PrinterSettings.BottomMargin = ssBottom;
			ws.PrinterSettings.LeftMargin = ssLeft;
			ws.PrinterSettings.RightMargin = ssRight;
			ws.PrinterSettings.HeaderMargin = ssHeader;
			ws.PrinterSettings.FooterMargin = ssFooter;
		} // MssWorksheet_SetMargins

		/// <summary>
		/// Turn a cell range into a native Excel Table (ListObject) with a built-in style, banded rows, header, and auto-filter — giving structured, filterable data instead of plain cells.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRange">Range incl. header row, e.g. A1:D20.</param>
		/// <param name="ssTableName">Unique table name.</param>
		/// <param name="ssTableStyle">e.g. None, Light1, Medium9, Dark1.</param>
		/// <param name="ssShowHeader">Show the header row.</param>
		/// <param name="ssShowFilter">Show auto-filter dropdowns.</param>
		/// <param name="ssShowTotal">Show the totals row.</param>
		public void MssWorksheet_AddTable(object ssWorksheet, string ssRange, string ssTableName, string ssTableStyle, bool ssShowHeader, bool ssShowFilter, bool ssShowTotal) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			var table = ws.Tables.Add(ws.Cells[ssRange], ssTableName);
			table.TableStyle = ParseEnum(ssTableStyle, OfficeOpenXml.Table.TableStyles.Medium2);
			table.ShowHeader = ssShowHeader;
			table.ShowFilter = ssShowFilter;
			table.ShowTotal = ssShowTotal;
		} // MssWorksheet_AddTable

		/// <summary>
		/// Add a clickable hyperlink to a cell, pointing to a URL, with optional display text shown in place of the raw link.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssCellName">Cell, e.g. A1.</param>
		/// <param name="ssUrl">Full URL incl. scheme, e.g. https://example.com.</param>
		/// <param name="ssDisplayText">Text shown in the cell; defaults to the URL.</param>
		public void MssCell_AddHyperlink(object ssWorksheet, string ssCellName, string ssUrl, string ssDisplayText) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			string display = string.IsNullOrEmpty(ssDisplayText) ? ssUrl : ssDisplayText;
			ExcelRange cell = ws.Cells[ssCellName];
			cell.Hyperlink = new ExcelHyperLink(ssUrl) { Display = display };
			cell.Value = display;
			cell.Style.Font.UnderLine = true;
			cell.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
		} // MssCell_AddHyperlink

		/// <summary>
		/// Set the color of the worksheet&apos;s sheet tab using a hex color code (e.g. #FF0000).
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssHexColor">Hex color, e.g. #FF0000.</param>
		public void MssWorksheet_SetTabColor(object ssWorksheet, string ssHexColor) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			ws.TabColor = Util.ConvertFromColorCode(ssHexColor);
		} // MssWorksheet_SetTabColor

		/// <summary>
		/// Copy a range of rows
		/// </summary>
		/// <param name="ssWorksheet"></param>
		/// <param name="ssRangeStart">Example: A1:B5</param>
		/// <param name="ssRangeEnd">Example: G1:H5</param>
		public void MssWorksheet_CopyRange(object ssWorksheet, string ssRangeStart, string ssRangeEnd) {
			var ws = AsWorksheet(ssWorksheet);
			ws.Cells[ssRangeStart].Copy(ws.Cells[ssRangeEnd]);
		} // MssWorksheet_CopyRange

        /// <summary>
        /// Get all merged cell ranges in the selected workbook.
        /// E.g., Worksheet: Sheet1
        /// A1:B2; D4:E4
        /// Worksheet: Sheet2
        /// C3:D5
        /// </summary>
        /// <param name="ssWorkbook">The workbook to work with</param>
        /// <param name="ssMergedRange">Ranges of the merged cells</param>
        public void MssWorkbook_GetMergedCellRanges(object ssWorkbook, out string ssMergedRange)
        {
            ssMergedRange = "";
            bool anyMergedCells = false;
            ExcelPackage p = ssWorkbook as ExcelPackage;

            foreach (var worksheet in p.Workbook.Worksheets)
            {
                if (worksheet.MergedCells.Count > 0)
                {
                    anyMergedCells = true;
                    ssMergedRange += $"Worksheet: {worksheet.Name}\n";

                    foreach (var mergedRange in worksheet.MergedCells)
                    {
                        ssMergedRange += $"{mergedRange}; ";
                    }

                    ssMergedRange = ssMergedRange.TrimEnd(new char[] { ';', ' ' });
                    ssMergedRange += "\n\n";
                }
            }

            if (!anyMergedCells)
            {
                ssMergedRange = "No merged cells found in any worksheet.";
            }
        } // MssWorkbook_GetMergedCellRanges

        /// <summary>
        /// Get all merged cell ranges in the selected worksheet. E.g., A1:A3; B1:C2.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssMergedRange">Ranges of the merged cells</param>
        public void MssWorksheet_GetMergedCellRanges(object ssWorksheet, out string ssMergedRange)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ssMergedRange = "";

            if (ws.MergedCells.Count > 0)
            {
                foreach (var mergedRange in ws.MergedCells)
                {
                    ssMergedRange += mergedRange + "; ";
                }

                ssMergedRange = ssMergedRange.TrimEnd(new char[] { ';', ' ' });
            }
            else
            {
                ssMergedRange = "No merged cells found.";
            }
        } // MssWorksheet_GetMergedCellRanges

        /// <summary>
        /// Gets the binary data of the workbook, setting worksheets to right-to-left view.
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssBinaryData"></param>
        public void MssWorkbook_SaveRightToLeft(object ssWorkbook, out byte[] ssBinaryData)
        {
            ExcelPackage p = ssWorkbook as ExcelPackage;

            if (p == null)
            {
                ssBinaryData = null;
                return;
            }

            foreach (var worksheet in p.Workbook.Worksheets)
            {
                worksheet.View.RightToLeft = true;
            }

            Util.PreserveVisibleRowsForZeroHeightSheets(p);
            ssBinaryData = p.GetAsByteArray();
            // GetAsByteArray closes the package; reload so the workbook stays usable.
            p.Load(new System.IO.MemoryStream(ssBinaryData));
        } // MssWorkbook_SaveRightToLeft

        /// <summary>
        /// Freeze cells, defined by row and column numbers.
        /// Example:
        /// - Choosing row = 2, column = 1 will freeze the first row.
        /// - Choosing row = 1, column = 2 will freeze the first column.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">Row Number</param>
        /// <param name="ssColumn">Column Number</param>
        public void MssCell_Freeze(object ssWorksheet, int ssRow, int ssColumn)
        {

            // Select the worksheet
            ExcelWorksheet ws;
            ws = AsWorksheet(ssWorksheet);

            // Freeze selected rows and columns
            ws.View.FreezePanes(ssRow, ssColumn);
        } // MssCell_Freeze

        /// <summary>
        /// Action to convert Hex code of color to RGB value
        /// </summary>
        /// <param name="ssHexCode">Color hex code (eg. #FFFFFF)</param>
        /// <param name="ssRGB">Color RGB value (eg. RGB(255, 255, 255))</param>
        public void MssUtil_ConvertHexCodeToRGB(string ssHexCode, out string ssRGB)
        {
            // Convert the hexadecimal color code to RGB format and assign it to output RGB
            if (!string.IsNullOrEmpty(ssHexCode) && ssHexCode != "No Fill Color")
            {
                Color rgbColor = ColorTranslator.FromHtml(ssHexCode);
                ssRGB = $"RGB({rgbColor.R}, {rgbColor.G}, {rgbColor.B})";
            }
            else
            {
                ssRGB = "No Fill Color";
            }
        } // MssUtil_ConvertHexCodeToRGB

        /// <summary>
        /// Get fill color of a cell, defined by its index.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">Row number</param>
        /// <param name="ssColumn">Column number</param>
        /// <param name="ssFillColor">Fill color of the cell</param>
        public void MssCell_GetFillColorByIndex(object ssWorksheet, int ssRow, int ssColumn, out string ssFillColor)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ssFillColor = ReadFillColorHex(ws.Cells[ssRow, ssColumn]);
        } // MssCell_GetFillColorByIndex

        /// <summary>
        /// Read a cell's fill color as a #RRGGBB hex string, or "No Fill Color" when unset.
        /// </summary>
        private static string ReadFillColorHex(ExcelRange cell)
        {
            var fillColor = cell.Style.Fill.BackgroundColor.Rgb;
            if (string.IsNullOrEmpty(fillColor))
            {
                return "No Fill Color";
            }
            return "#" + (fillColor.Length >= 6 ? fillColor.Substring(fillColor.Length - 6) : fillColor.PadLeft(6, '0'));
        }

        /// <summary>
        /// Get fill color of a cell, defined by its name.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell name (eg. A1)</param>
        /// <param name="ssFillColor">Fill color of the cell</param>
        public void MssCell_GetFillColorByName(object ssWorksheet, string ssCellName, out string ssFillColor)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ssFillColor = ReadFillColorHex(ws.Cells[ssCellName]);
        } // MssCell_GetFillColorByName

        /// <summary>
        /// Get the Microsoft Office properties of the Excel document.
        /// </summary>
        /// <param name="ssWorkbook">The workbook</param>
        /// <param name="ssProperties">The Microsoft Office properties of the Excel document.</param>
        public void MssExcel_GetProperties(object ssWorkbook, out RCOfficePropertiesRecord ssProperties)
        {
            var wb = ssWorkbook as ExcelWorkbook ?? (ssWorkbook as ExcelPackage)?.Workbook;
            var props = wb.Properties;
            ssProperties = new RCOfficePropertiesRecord(null);
            ssProperties.ssSTOfficeProperties.ssAuthor = props.Author;
            ssProperties.ssSTOfficeProperties.ssCategory = props.Category;
            ssProperties.ssSTOfficeProperties.ssComments = props.Comments;
            ssProperties.ssSTOfficeProperties.ssCompany = props.Company;
            ssProperties.ssSTOfficeProperties.ssKeywords = props.Keywords;
            ssProperties.ssSTOfficeProperties.ssLastModifiedBy = props.LastModifiedBy;
            ssProperties.ssSTOfficeProperties.ssManager = props.Manager;
            ssProperties.ssSTOfficeProperties.ssStatus = props.Status;
            ssProperties.ssSTOfficeProperties.ssSubject = props.Subject;
            ssProperties.ssSTOfficeProperties.ssTitle = props.Title;
        } // MssExcel_GetProperties

        /// <summary>
        /// Set the Microsoft Office properties of the Excel document.
        /// </summary>
        /// <param name="ssWorkbook">The workbook</param>
        /// <param name="ssProperties">The Microsoft Office properties of the Excel document.</param>
        /// <param name="ssIgnoreBlank">If True, any blank properties in the Properties structure provided will be left with their existing values. If False, any blank properties in the Properties structure provided will be set to blank.</param>
        public void MssExcel_SetProperties(object ssWorkbook, RCOfficePropertiesRecord ssProperties, bool ssIgnoreBlank)
        {
            var wb = ssWorkbook as ExcelWorkbook ?? (ssWorkbook as ExcelPackage)?.Workbook;
            var props = wb.Properties;
            var inProps = ssProperties.ssSTOfficeProperties;
            if (!string.IsNullOrEmpty(inProps.ssAuthor)) { props.Author = inProps.ssAuthor; }
            if (!string.IsNullOrEmpty(inProps.ssCategory)) { props.Category = inProps.ssCategory; }
            if (!string.IsNullOrEmpty(inProps.ssComments)) { props.Comments = inProps.ssComments; }
            if (!string.IsNullOrEmpty(inProps.ssCompany)) { props.Company = inProps.ssCompany; }
            if (!string.IsNullOrEmpty(inProps.ssKeywords)) { props.Keywords = inProps.ssKeywords; }
            if (!string.IsNullOrEmpty(inProps.ssLastModifiedBy)) { props.LastModifiedBy = inProps.ssLastModifiedBy; }
            if (!string.IsNullOrEmpty(inProps.ssManager)) { props.Manager = inProps.ssManager; }
            if (!string.IsNullOrEmpty(inProps.ssStatus)) { props.Status = inProps.ssStatus; }
            if (!string.IsNullOrEmpty(inProps.ssSubject)) { props.Subject = inProps.ssSubject; }
            if (!string.IsNullOrEmpty(inProps.ssTitle)) { props.Title = inProps.ssTitle; }
        } // MssExcel_SetProperties

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
        public void MssExcel_ClearProperties(object ssWorkbook, bool ssClearTitle, bool ssClearSubject, bool ssClearAuthor, bool ssClearComments, bool ssClearKeywords, bool ssClearLastModifiedBy, bool ssClearCategory, bool ssClearStatus, bool ssClearCompany, bool ssClearManager)
        {
            var wb = ssWorkbook as ExcelWorkbook ?? (ssWorkbook as ExcelPackage)?.Workbook;
            var props = wb.Properties;
            if (ssClearAuthor) { props.Author = string.Empty; }
            if (ssClearCategory) { props.Category = string.Empty; }
            if (ssClearComments) { props.Comments = string.Empty; }
            if (ssClearCompany) { props.Company = string.Empty; }
            if (ssClearKeywords) { props.Keywords = string.Empty; }
            if (ssClearLastModifiedBy) { props.LastModifiedBy = string.Empty; }
            if (ssClearManager) { props.Manager = string.Empty; }
            if (ssClearStatus) { props.Status = string.Empty; }
            if (ssClearSubject) { props.Subject = string.Empty; }
            if (ssClearTitle) { props.Title = string.Empty; }
        } // MssExcel_ClearProperties


        /// <summary>
        /// Input text address and get back the Row/Col values
        /// </summary>
        /// <param name="ssAddress">Text address, e.g. AB47 or A11:AB47</param>
        /// <param name="ssRowStart">Address row or range start row</param>
        /// <param name="ssColStart">Address col or range start column</param>
        /// <param name="ssRowEnd">Range end row</param>
        /// <param name="ssColEnd">Range end column</param>
        public void MssAddress_From_Text(string ssAddress, out int ssRowStart, out int ssColStart, out int ssRowEnd, out int ssColEnd)
        {
            ssRowStart = 0;
            ssColStart = 0;
            ssRowEnd = 0;
            ssColEnd = 0;

            if (string.IsNullOrEmpty(ssAddress))
            {
                throw new ArgumentException("Address cannot be empty");
            }
            ExcelAddress address = new ExcelAddress(ssAddress);
            ssRowStart = address.Start.Row;
            ssColStart = address.Start.Column;
            ssRowEnd = address.End.Row;
            ssColEnd = address.End.Column;
        } // MssAddress_From_Text

        /// <summary>
        /// Input Row/Col values and get the text address
        /// </summary>
        /// <param name="ssRowStart">Start row of the address</param>
        /// <param name="ssColStart">Start column of the address</param>
        /// <param name="ssRowEnd">End row of the address or zero</param>
        /// <param name="ssColEnd">End column of the address or zero</param>
        /// <param name="ssAddress">Text address, e.g. AB47 or C11:AB47</param>
        public void MssAddress_From_RowCol(int ssRowStart, int ssColStart, int ssRowEnd, int ssColEnd, out string ssAddress)
        {
            ssAddress = "";
            if (ssRowStart < 1 || ssColStart < 1)
            {
                throw new ArgumentException("RowStart and ColStart need to be greater than 0");
            }
            if (ssRowEnd <= 0 || ssColEnd <= 0)
            {
                ssRowEnd = ssRowStart;
                ssColEnd = ssColStart;
            }
            ssAddress = (new ExcelAddress(ssRowStart, ssColStart, ssRowEnd, ssColEnd)).Address;
        } // MssAddress_From_RowCol

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
        public void MssWorksheet_AddDropdown(object ssWorksheet, RLItemsRecordList ssItemsList, string ssItemsAddress, string ssCellRange, string ssTitleMessage, string ssPromptMessage, bool ssShowError, string ssCustomErrorMessage, string ssCustomErrorTitle)
        {
            /* 
             * Miguel 'Kelter' Antunes
             * Code from Advanced_Excel_Dropdowns component
             * https://www.outsystems.com/forge/component-overview/10562/advanced-excel-dropdowns
             * 
             */

            var ws = AsWorksheet(ssWorksheet);
            var unitMeasure = ws.DataValidations.AddListValidation(ssCellRange);

            if (String.IsNullOrEmpty(ssItemsAddress)) //Check if address string is empty, proceed to fill list with items.
            {
                for (int i = 0; i < ssItemsList.Count; i++)
                {
                    unitMeasure.Formula.Values.Add(ssItemsList[i].ssSTItems.ssItemText);
                }
            }
            else
            {
                //TODO: Validation of Formula?
                //TODO: Input sheet as Object instead of within string.
                unitMeasure.Formula.ExcelFormula = ssItemsAddress;
            }

            unitMeasure.ShowInputMessage = true;
            unitMeasure.PromptTitle = ssTitleMessage;
            unitMeasure.Prompt = ssPromptMessage;
            unitMeasure.Error = ssCustomErrorMessage;
            unitMeasure.ErrorTitle = ssCustomErrorTitle;
            unitMeasure.ShowErrorMessage = ssShowError;
            unitMeasure.AllowBlank = true;




        } // MssWorksheet_AddDropdown

        /// <summary>
        /// Set the active sheet
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssWorksheetName"></param>
        /// <param name="ssWorksheetIndex"></param>
        public void MssWorksheet_SetActive(object ssWorkbook, string ssWorksheetName, int ssWorksheetIndex)
        {
            ExcelPackage ee = ssWorkbook as ExcelPackage;
            ExcelWorksheet target = null;

            if (!string.IsNullOrEmpty(ssWorksheetName))
            {
                target = ee.Workbook.Worksheets[ssWorksheetName];
            }
            if (ssWorksheetIndex > 0)
            {
                target = ee.Workbook.Worksheets[ssWorksheetIndex];
            }

            if (target == null) return;

            foreach (var sheet in ee.Workbook.Worksheets)
            {
                sheet.View.TabSelected = false;
            }
            target.View.TabSelected = true;
            ee.Workbook.View.ActiveTab = target.Index;
        } // MssWorksheet_SetActive

        /// <summary>
        /// Write a converted value to a cell, defined by its index.
        /// Input is a worksheet-object
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
        public void MssCell_WriteByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssCellValue, string ssCellType)
        {
            RCCellFormatRecord format = new RCCellFormatRecord();
            MssCell_WriteByIndexWithFormat(ssWorksheet, ssRow, ssColumn, ssCellValue, ssCellType, format);
        } // MssCell_WriteByIndex

        /// <summary>
        /// Write a converted value to a cell, defined by its index.
        /// Input is a worksheet-object.
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
        public void MssCell_WriteByIndexWithFormat(object ssWorksheet, int ssRow, int ssColumn, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat)
        {
            MssCell_Write(ssWorksheet, null, ssRow, ssColumn, ssCellValue, ssCellType, false, ssCellFormat);
        } // MssCell_WriteByIndexWithFormat

        /// <summary>
        /// Write a converted value to a cell, defined by its name.
        /// Input is a worksheet-object
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
        public void MssCell_WriteByName(object ssWorksheet, string ssCellName, string ssCellValue, string ssCellType)
        {
            RCCellFormatRecord format = new RCCellFormatRecord();
            MssCell_WriteByNameWithFormat(ssWorksheet, ssCellName, ssCellValue, ssCellType, format);
        } // MssCell_WriteByName

        /// <summary>
        /// Write a converted value to a cell, defined by its name.
        /// Input is a worksheet-object.
        /// Accepts format for the target cell
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
        /// <param name="ssCellFormat">CellFormat for the target cell</param>
        public void MssCell_WriteByNameWithFormat(object ssWorksheet, string ssCellName, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat)
        {
            MssCell_Write(ssWorksheet, ssCellName, 0, 0, ssCellValue, ssCellType, false, ssCellFormat);
        } // MssCell_WriteByNameWithFormat

        /// <summary>
        /// Write a dataset to a range of column cells
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssRow"></param>
        /// <param name="ssColumnStart"></param>
        /// <param name="ssValueList"></param>
        /// <param name="ssCellType">Type can be:
        /// general (default if empty)
        /// text,
        /// datetime,
        /// integer,
        /// decimal,
        /// boolean,
        /// formula</param>
        public void MssCell_WriteColumnRange(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType)
        {
            RCCellFormatRecord format = new RCCellFormatRecord();
            MssCell_WriteColumnRangeWithFormat(ssWorksheet, ssRow, ssColumnStart, ssValueList, ssCellType, format);
        } // MssCell_WriteColumnRange

        /// <summary>
        /// Write a dataset to a range of column cells
        /// Accepts format for the target cells
        /// </summary>
        /// <param name="ssWorksheet">Worksheet to write to</param>
        /// <param name="ssRow">rownumber</param>
        /// <param name="ssColumnStart">Start column (integer)</param>
        /// <param name="ssValueList">Values to write to columns</param>
        /// <param name="ssCellType">Type can be:
        /// general (default if empty)
        /// text,
        /// datetime,
        /// integer,
        /// decimal,
        /// boolean,
        /// formula</param>
        /// <param name="ssCellFormat">CellFormat for the target cells</param>
        public void MssCell_WriteColumnRangeWithFormat(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType, RCCellFormatRecord ssCellFormat)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            DataTable dt = Util.RecordListToDataTable((RecordList)ssValueList);

            if (dt != null)
            {
                ExcelRange range = (ExcelRange)ws.Cells[ssRow, ssColumnStart].LoadFromDataTable(Util.Transpose(dt, ssCellType), false);

                Util.ApplyFormatToRange(range, ssCellFormat);
            }
        } // MssCell_WriteColumnRangeWithFormat

        /// <summary>
        /// Write a image on a cell, defined by its index.
        /// Input are a worksheet-object and a image-object.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">row number</param>
        /// <param name="ssColumn">column number</param>
        /// <param name="ssImageName">The image name</param>
        /// <param name="ssImage">The image to write.</param>
        public void MssCell_WriteImageByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssImageName, byte[] ssImage)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            if (ssImage == null || ssImage.Length == 0)
            {
                throw new ArgumentException("No image data supplied.", nameof(ssImage));
            }

            using (MemoryStream ms = new MemoryStream(ssImage))
            using (Image i = Image.FromStream(ms))
            {
                ExcelPicture pic = ws.Drawings.AddPicture(ssImageName, i);
                pic.SetPosition(ssRow - 1, 0, ssColumn - 1, 0);
            }
        } // MssCell_WriteImageByIndex

        /// <summary>
        /// Write a image on a cell, defined by its name.
        /// Input are a worksheet-object and a image-object.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell-name (eg A4)</param>
        /// <param name="ssImageName">The image name</param>
        /// <param name="ssImage">The image to write.</param>
        public void MssCell_WriteImageByName(object ssWorksheet, string ssCellName, string ssImageName, byte[] ssImage)
        {
            ExcelCellAddress addr = new ExcelCellAddress(ssCellName);
            MssCell_WriteImageByIndex(ssWorksheet, addr.Row, addr.Column, ssImageName, ssImage);
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
        public void MssCell_WriteRangeWithFormat(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord ssCellFormat)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            DataTable dt = Util.RecordListToDataTable((RecordList)ssDataSet);

            if (dt != null)
            {
                ExcelRange range = (ExcelRange)ws.Cells[ssRowStart, ssColumnStart].LoadFromDataTable(dt, false);

                Util.ApplyFormatToRange(range, ssCellFormat);
            }
        } // MssCell_WriteRangeWithFormat

        /// <summary>
        /// Add a worksheet to work on by its name
        /// </summary>
        /// <param name="ssWorkbook">Workbook where the sheet is to be added</param>
        /// <param name="ssWorksheetName">The name of the spreadsheet to create</param>
        /// <param name="ssWorksheet">The newly added worksheet</param>
        public void MssWorkbook_AddName(object ssWorkbook, string ssWorksheetName, out object ssWorksheet)
        {
            ExcelPackage p = (ExcelPackage)ssWorkbook;
            ExcelWorksheet ws;
            ws = p.Workbook.Worksheets.Add(ssWorksheetName);
            ssWorksheet = ws;
        } // MssWorkbook_AddName

        /// <summary>
        /// Set the pixel width of a column on a specific worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssColumnNumber">The column number, starting at 1</param>
        /// <param name="ssDesiredWidth">The pixel width you desire for the column.</param>
        public void MssColumn_SetWidth(object ssWorksheet, int ssColumnNumber, decimal ssDesiredWidth)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Column(ssColumnNumber).Width = (double)ssDesiredWidth;
        } // MssColumn_SetWidth

        /// <summary>
        /// Set the pixel height for a specific row in a worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with
        /// </param>
        /// <param name="ssRowNumber">The number of the row to set the height for</param>
        /// <param name="ssDesiredHeight">The desired pixel height for the row</param>
        public void MssRow_SetHeight(object ssWorksheet, int ssRowNumber, decimal ssDesiredHeight)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Row(ssRowNumber).Height = (double)ssDesiredHeight;
        } // MssRow_SetHeight

        /// <summary>
        /// Calculates the formula of a cell, defined by its name.
        /// Input is a worksheet-object
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell-name (eg A4)</param>
        public void MssCell_CalculateByName(object ssWorksheet, string ssCellName)
        {
            ExcelCellAddress addr = new ExcelCellAddress(ssCellName);
            MssCell_CalculateByIndex(ssWorksheet, addr.Row, addr.Column);
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
        public void MssCell_FormatRange(object ssWorksheet, int ssRowStart, int ssColumnStart, int ssRowEnd, int ssColumnEnd, RCCellFormatRecord ssCellFormat)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            Util.ApplyFormatToRange(ws.Cells[ssRowStart, ssColumnStart, ssRowEnd, ssColumnEnd], ssCellFormat);
        } // MssCell_FormatRange

        /// <summary>
        /// Reads the value of a cell, defined by its index.
        /// Input is a worksheet-object
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">row number</param>
        /// <param name="ssColumn">column number</param>
        /// <param name="ssReadText">If true always reads the cell value as text</param>
        /// <param name="ssCellValue">text-value</param>
        public void MssCell_ReadByIndex(object ssWorksheet, int ssRow, int ssColumn, bool ssReadText, out string ssCellValue)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            try
            {
                if (ssReadText)
                    ssCellValue = ws.Cells[ssRow, ssColumn].Text;
                else
                    ssCellValue = Convert.ToString(ws.GetValue(ssRow, ssColumn));
            }
            catch (Exception ex)
            {
                Util.LogMessage("MssCell_ReadByIndex failed at [" + ssRow + "," + ssColumn + "]: " + ex.Message);
                ssCellValue = String.Empty;
            }
        } // MssCell_ReadByIndex

        /// <summary>
        /// Reads the value of a cell, defined by its name.
        /// Input is a worksheet-object
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell-name (eg A4)</param>
        /// <param name="ssReadText">If true always reads the cell value as text</param>
        /// <param name="ssCellValue">text-value</param>
        public void MssCell_ReadByName(object ssWorksheet, string ssCellName, bool ssReadText, out string ssCellValue)
        {
            ExcelCellAddress addr = new ExcelCellAddress(ssCellName);
            MssCell_ReadByIndex(ssWorksheet, addr.Row, addr.Column, ssReadText, out ssCellValue);
        } // MssCell_ReadByName

        /// <summary>
        /// Write a formula to a cell, defined by its index.
        /// Input is a worksheet-object
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">rownumber</param>
        /// <param name="ssColumn">columnnumber</param>
        /// <param name="ssFormula">Formula</param>
        public void MssCell_SetFormulaByIndex(object ssWorksheet, int ssRow, int ssColumn, string ssFormula)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Cells[ssRow, ssColumn].Formula = ssFormula;
        } // MssCell_SetFormulaByIndex

        /// <summary>
        /// Write a formula to a cell, defined by its name.
        /// Input is a worksheet-object
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell-name (eg A4)</param>
        /// <param name="ssFormula">Formula</param>
        public void MssCell_SetFormulaByName(object ssWorksheet, string ssCellName, string ssFormula)
        {
            ExcelCellAddress addr = new ExcelCellAddress(ssCellName);
            MssCell_SetFormulaByIndex(ssWorksheet, addr.Row, addr.Column, ssFormula);
        } // MssCell_SetFormulaByName

        /// <summary>
        /// Adds a copy of a worksheet
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssWorksheetName">The name of the spreadsheet to create</param>
        /// <param name="ssWorksheetToCopy">The worksheet to be copied</param>
        /// <param name="ssWorksheet"></param>
        public void MssWorkbook_AddCopyWorksheet(object ssWorkbook, string ssWorksheetName, object ssWorksheetToCopy, out object ssWorksheet)
        {
            ExcelPackage p = (ExcelPackage)ssWorkbook;
            ExcelWorksheet wsToCopy = AsWorksheet(ssWorksheetToCopy);
            ExcelWorksheet ws;
            ws = p.Workbook.Worksheets.Add(ssWorksheetName, wsToCopy);
            ssWorksheet = ws;
        } // MssWorkbook_AddCopyWorksheet

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssImages"></param>
        public void MssWorksheet_GetImages(object ssWorksheet, out RLImageRecordList ssImages)
        {
            ssImages = new RLImageRecordList();

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            var pics = ws.Drawings;
            for (int i = 0; i < pics.Count; i++)
            {
                ExcelPicture picture = pics[i] as ExcelPicture;
                if (picture == null)
                {
                    continue;
                }

                RCImageRecord Img = new RCImageRecord();
                Img.ssSTImage.ssName = picture.Name;
                Img.ssSTImage.ssContent = Util.ImageToByteArray(picture.Image);
                Img.ssSTImage.ssColumn = picture.From.Column;
                Img.ssSTImage.ssRow = picture.From.Row;
                ssImages.Append(Img);
            }
        } // MssWorksheet_GetImages

        /// <summary>
        /// 
        /// </summary>
        public void MssWorksheet_SelectByIndex(object ssWorkbook, int ssWorksheetNumber, out object ssWorksheet)
        {
            ExcelWorkbook wb;
            ExcelWorksheet ws;
            ExcelPackage p;
            p = (ExcelPackage)ssWorkbook;
            wb = p.Workbook;
            ws = wb.Worksheets[ssWorksheetNumber];

            ssWorksheet = ws;
        } // MssWorksheet_SelectByIndex

        /// <summary>
        /// Select a worksheet to work on by its name
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssWorksheetName">The name of the spreadsheet to select</param>
        /// <param name="ssWorksheet"></param>
        public void MssWorksheet_SelectByName(object ssWorkbook, string ssWorksheetName, out object ssWorksheet)
        {
            ExcelPackage p = (ExcelPackage)ssWorkbook;
            ExcelWorksheet ws;
            ws = p.Workbook.Worksheets[ssWorksheetName];
            ssWorksheet = ws;
        } // MssWorksheet_SelectByName

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssNameToDelete"></param>
        public void MssWorksheet_DeleteByName(object ssWorkbook, string ssNameToDelete)
        {
            ExcelPackage ee = ssWorkbook as ExcelPackage;
            ee.Workbook.Worksheets.Delete(ssNameToDelete);
        } // MssWorksheet_DeleteByName

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssIndexToDelete"></param>
        public void MssWorksheet_DeleteByIndex(object ssWorkbook, int ssIndexToDelete)
        {
            ExcelPackage ee = ssWorkbook as ExcelPackage;
            ee.Workbook.Worksheets.Delete(ssIndexToDelete);
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
        public void MssWorksheet_Chart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos)
        {
            MssChart_Create(ssWorksheet, ssChartType, ssChartName, ssDataSeries_List, ssHeight, ssWidth, ssRowPos, ssColPos);
        } // MssWorksheet_Chart_Create

        /// <summary>
        /// Create a defined &quot;Name&quot; (a word or string of characters in Excel that represents a cell, range of cells, formula, or constant value) in excel, starting in the RowStart / ColumnStart cell.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet to write to</param>
        /// <param name="ssName">&quot;Name&quot;</param>
        /// <param name="ssDataSet">Values to assigned the name</param>
        /// <param name="ssRowStart">Start row number</param>
        /// <param name="ssColumnStart">Start column number</param>
        public void MssWorksheet_AddName(object ssWorksheet, string ssName, object ssDataSet, int ssRowStart, int ssColumnStart)
        {
            ExcelWorksheet ws;
            ws = AsWorksheet(ssWorksheet);

            if (!ws.Names.ContainsKey(ssName))
            {
                MssCell_WriteRange(ssWorksheet, ssRowStart, ssColumnStart, ssDataSet, new RCCellFormatRecord(), false);
                ws.Workbook.Names.Add(ssName, ws.Cells[ssRowStart, ssColumnStart, ssRowStart + ((RecordList)ssDataSet).Length - 1, ssColumnStart]);
            }
        } // MssWorksheet_AddName

        /// <summary>
        /// Opens an existing workbook for editing and keeps it in memory
        /// </summary>
        /// <param name="ssBinaryData"></param>
        /// <param name="ssWorkbook"></param>
        public void MssWorkbook_Open_BinaryData(byte[] ssBinaryData, out object ssWorkbook)
        {
            MssWorkbook_Open("", ssBinaryData, out ssWorkbook);
        } // MssWorkbook_Open_BinaryData

        /// <summary>
        /// Creates a new excel workbook, optionally specifying the name of the fiirst sheet.
        /// </summary>
        /// <param name="ssWorkbook">The newly created workbook</param>
        /// <param name="ssFirstSheetName">Specify the name of the initial sheet in the workbook. Default = &quot;Sheet1&quot;</param>
        /// <param name="ssNumberOfSheets">The number of sheets to add. Sheet names will be auto generated, i.e. Sheet1, Sheet2.</param>
        /// <param name="ssSheetNames">List of new sheets to add, with at least a name specified. The index, if specified, will be used to add sheets in that order.
        /// FirstSheetName and NrSheets are ignored if SheetNames is populated</param>
        public void MssWorkbook_Create(int ssNumberOfSheets, string ssFirstSheetName, RLNewSheetRecordList ssSheetNames, out object ssWorkbook)
        {
            ssWorkbook = null;

            ExcelPackage p = new ExcelPackage();
            ExcelWorkbook wb = p.Workbook;

            if (ssSheetNames == null || ssSheetNames.Count == 0 || (ssSheetNames.Count == 1 && ssSheetNames[0].ssSTNewSheet.ssName == "" && ssSheetNames[0].ssSTNewSheet.ssIndex == 0))
            {
                if (string.IsNullOrEmpty(ssFirstSheetName))
                {
                    ssFirstSheetName = "Sheet1";
                }

                wb.Worksheets.Add(ssFirstSheetName);
                if (ssNumberOfSheets > 1)
                {
                    // Strip trailing "1" so "Sheet1" produces Sheet1/Sheet2/Sheet3, not Sheet1/Sheet12/Sheet13.
                    string baseName = (ssFirstSheetName.Length > 1 && ssFirstSheetName.EndsWith("1"))
                        ? ssFirstSheetName.Substring(0, ssFirstSheetName.Length - 1)
                        : ssFirstSheetName;
                    for (int i = 2; i <= ssNumberOfSheets; i++)
                    {
                        wb.Worksheets.Add(baseName + i);
                    }
                }
            }
            else
            {
                ssSheetNames.Sort(s => s.ssSTNewSheet.ssIndex, true);
                foreach (RCNewSheetRecord item in ssSheetNames)
                {
                    wb.Worksheets.Add(item.ssSTNewSheet.ssName);
                }
            }
            ssWorkbook = p;
        } // MssWorkbook_Create

        /// <summary>
        /// Set protection on the workbook level
        /// </summary>
        /// <param name="ssWorkbook">The workbook to work with</param>
        /// <param name="ssPassword">The password to set for the workbook. This does not encrypt the workbook.</param>
        /// <param name="ssLockStructure">Locks the structure,which prevents users from adding or deleting worksheets or from displaying hidden worksheets.</param>
        /// <param name="ssLockWindows">Locks the position of the workbook window.</param>
        /// <param name="ssLockRevision">Lock the workbook for revision</param>
        public void MssWorkbook_Protect(object ssWorkbook, string ssPassword, bool ssLockStructure, bool ssLockWindows, bool ssLockRevision)
        {
            ExcelPackage p = ssWorkbook as ExcelPackage;
            ExcelWorkbook wb = p.Workbook;

            if (!string.IsNullOrEmpty(ssPassword))
            {
                wb.Protection.SetPassword(ssPassword);
            }

            wb.Protection.LockRevision = ssLockRevision;
            wb.Protection.LockStructure = ssLockStructure;
            wb.Protection.LockWindows = ssLockWindows;

        } // MssWorkbook_Protect

        /// <summary>
        /// Set protection on an Excel Worksheet
        /// </summary>
        /// <param name="ssWorksheet">Worksheet to protect</param>
        /// <param name="ssPassword">Password to protect the worksheet with.</param>
        /// <param name="ssProtectionOptions"></param>
        public void MssWorksheet_Protect(object ssWorksheet, string ssPassword, RCProtectionRecord ssProtectionOptions)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ws.Protection.IsProtected = ssProtectionOptions.ssSTProtection.ssIsProtected;

            ws.Protection.AllowAutoFilter = ssProtectionOptions.ssSTProtection.ssAllowAutoFilter;
            ws.Protection.AllowDeleteColumns = ssProtectionOptions.ssSTProtection.ssAllowDeleteColumns;
            ws.Protection.AllowDeleteRows = ssProtectionOptions.ssSTProtection.ssAllowDeleteRows;
            ws.Protection.AllowEditObject = ssProtectionOptions.ssSTProtection.ssAllowEditObject;
            ws.Protection.AllowEditScenarios = ssProtectionOptions.ssSTProtection.ssAllowEditScenarios;
            ws.Protection.AllowFormatCells = ssProtectionOptions.ssSTProtection.ssAllowFormatCells;
            ws.Protection.AllowFormatColumns = ssProtectionOptions.ssSTProtection.ssAllowFormatColumns;
            ws.Protection.AllowFormatRows = ssProtectionOptions.ssSTProtection.ssAllowFormatRows;
            ws.Protection.AllowInsertColumns = ssProtectionOptions.ssSTProtection.ssAllowInsertColumns;
            ws.Protection.AllowInsertHyperlinks = ssProtectionOptions.ssSTProtection.ssAllowInsertHyperlinks;
            ws.Protection.AllowInsertRows = ssProtectionOptions.ssSTProtection.ssAllowInsertRows;
            ws.Protection.AllowPivotTables = ssProtectionOptions.ssSTProtection.ssAllowPivotTables;
            ws.Protection.AllowSelectLockedCells = ssProtectionOptions.ssSTProtection.ssAllowSelectLockedCells;
            ws.Protection.AllowSelectUnlockedCells = ssProtectionOptions.ssSTProtection.ssAllowSelectUnlockedCells;
            ws.Protection.AllowSort = ssProtectionOptions.ssSTProtection.ssAllowSort;

            if (!string.IsNullOrEmpty(ssPassword))
            {
                ws.Protection.SetPassword(ssPassword);
            }
            else if (!string.IsNullOrEmpty(ssProtectionOptions.ssSTProtection.ssPassword))
            {
                ws.Protection.SetPassword(ssProtectionOptions.ssSTProtection.ssPassword);
            }
        } // MssWorksheet_Protect

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
        /// <param name="ssMarginLeft"></param>
        /// <param name="ssMarginTop"></param>
        public void MssImage_Insert(object ssWorksheet, byte[] ssImageFile, string ssImageType, string ssImageName, int ssRowNumber, int ssColumnNumber, string ssCellName, int ssImageWidth, int ssImageHeight, int ssMarginTop, int ssMarginLeft)
        {
            if (string.IsNullOrEmpty(ssCellName) && ssColumnNumber <= 0 && ssRowNumber <= 0)
            {
                throw new ArgumentException("You need to specify a valid cell name (i.e. A4) or cell index (row/column combination)");
            }

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            if (ssImageFile == null || ssImageFile.Length == 0)
            {
                throw new ArgumentException("No image data supplied.", nameof(ssImageFile));
            }

            ExcelRange range = ws.Cells["A1"];

            if (!string.IsNullOrEmpty(ssCellName))
            {
                range = ws.Cells[ssCellName];
            }
            else if (ssRowNumber > 0 && ssColumnNumber > 0)
            {
                range = ws.Cells[ssRowNumber, ssColumnNumber];
            }

            using (MemoryStream ms = new MemoryStream(ssImageFile))
            using (Bitmap bitmap = new Bitmap(ms))
            using (ExcelPicture picture = ws.Drawings.AddPicture(ssImageName, bitmap))
            {
                picture.SetPosition(range.Start.Row - 1, ssMarginTop, range.Start.Column - 1, ssMarginLeft);
                picture.SetSize(ssImageWidth, ssImageHeight);
            }
            range.Dispose();
        } // MssImage_Insert

        /// <summary>
        /// Add the automatic filter option of Excel to the specified range of cells.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with.</param>
        /// <param name="ssRangeToFilter">The range where to add the filter. If not supplied, the dimension of the worksheet will be used.</param>
        public void MssWorksheet_AddAutoFilter(object ssWorksheet, RCRangeRecord ssRangeToFilter)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            int startRow, startCol, endRow, endCol;

            if (ssRangeToFilter.ssSTRange.ssStartRow == 0 && ssRangeToFilter.ssSTRange.ssStartCol == 0 && ssRangeToFilter.ssSTRange.ssEndRow == 0 && ssRangeToFilter.ssSTRange.ssEndCol == 0)
            {
                if (ws.Dimension == null)
                {
                    return;
                }
                startRow = ws.Dimension.Start.Row;
                startCol = ws.Dimension.Start.Column;
                endRow = ws.Dimension.End.Row;
                endCol = ws.Dimension.End.Column;
            }
            else
            {
                startRow = ssRangeToFilter.ssSTRange.ssStartRow;
                startCol = ssRangeToFilter.ssSTRange.ssStartCol;
                endRow = ssRangeToFilter.ssSTRange.ssEndRow;
                endCol = ssRangeToFilter.ssSTRange.ssEndCol;
            }

            using (var range = ws.Cells[startRow, startCol, endRow, endCol])
            {
                range.AutoFilter = true;
            }
        } // MssWorksheet_AddAutoFilter

        /// <summary>
        /// Apply the column autofit action to the specified range of cells specified in the given worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        public void MssWorksheet_AutofitColumns(object ssWorksheet)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Cells.AutoFitColumns();
        } // MssWorksheet_AutofitColumns

        /// <summary>
        /// Delete a specified Conditional Formatting rule on a worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with.</param>
        /// <param name="ssRuleToDeleteIndex">The index of the rule to be deleted.</param>
        public void MssConditionalFormatting_DeleteRule(object ssWorksheet, int ssRuleToDeleteIndex)
        {
            if (ssRuleToDeleteIndex <= 0)
            {
                throw new IndexOutOfRangeException("Index needs to be >= 1.");
            }

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.ConditionalFormatting.RemoveAt(ssRuleToDeleteIndex - 1);
        } // MssConditionalFormatting_DeleteRule

        /// <summary>
        /// Delete ALL Conditional Formatting rules for a worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with.</param>
        public void MssConditionalFormatting_DeleteAllRules(object ssWorksheet)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.ConditionalFormatting.RemoveAll();
        } // MssConditionalFormatting_DeleteAllRules

        /// <summary>
        /// Add a comment to a cell
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with.</param>
        /// <param name="ssRowNumber">The row number of the cell to add the comment to.</param>
        /// <param name="ssColumnNumber">The column number of the cell to add the comment to.</param>
        /// <param name="ssText">The comment.</param>
        /// <param name="ssAuthor">The author of the comment.</param>
        /// <param name="ssAutofit"></param>
        public void MssComment_Add(object ssWorksheet, int ssRowNumber, int ssColumnNumber, string ssText, string ssAuthor, bool ssAutofit)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ExcelComment comment = ws.Comments.Add(ws.Cells[ssRowNumber, ssColumnNumber], ssText, ssAuthor);
            comment.AutoFit = ssAutofit;
        } // MssComment_Add

        /// <summary>
        /// Delete column(s) from a worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssStartColumnNumber">Column number where to start deleting columns.</param>
        /// <param name="ssNumberOfColumns">The number of rows to delete. Default = 1.</param>
        public void MssColumn_Delete(object ssWorksheet, int ssStartColumnNumber, int ssNumberOfColumns)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            // Delete all comments from cells in column(s) before deleting the column(s).
            // Considers the rows in the dimension of the worksheet to prevent unnecessary processing.
            int nrRows = ws.Dimension?.Rows ?? 0;
            RemoveCommentsInRange(ws, 1, nrRows, ssStartColumnNumber, ssStartColumnNumber + ssNumberOfColumns - 1);

            ws.DeleteColumn(ssStartColumnNumber, ssNumberOfColumns);
        } // MssColumn_Delete

        /// <summary>
        /// Delete comment(s) in a specified range
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with.</param>
        /// <param name="ssRange">Range to delete comments from.</param>
        public void MssComment_Delete(object ssWorksheet, RCRangeRecord ssRange)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            RemoveCommentsInRange(ws, ssRange.ssSTRange.ssStartRow, ssRange.ssSTRange.ssEndRow,
                                      ssRange.ssSTRange.ssStartCol, ssRange.ssSTRange.ssEndCol);
        } // MssComment_Delete

        /// <summary>
        /// Remove every comment within the given 1-based cell rectangle. A zero/empty range
        /// (start &gt; end) simply iterates nothing.
        /// </summary>
        private static void RemoveCommentsInRange(ExcelWorksheet ws, int startRow, int endRow, int startCol, int endCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    if (ws.Cells[row, col].Comment == null)
                    {
                        continue;
                    }
                    ws.Comments.Remove(ws.Cells[row, col].Comment);
                }
            }
        }

        /// <summary>
        /// Inserts a new column into the spreadsheet.  Existing columns to the right of the insert index will be shifted right.  All formula are updated to take account of the new column.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with.</param>
        /// <param name="ssInsertAt">Column number where to insert new column.</param>
        /// <param name="ssNumberOfColumns">The number of columns to insert.</param>
        /// <param name="ssCopyStylesFrom">Copy Styles from this column. Applied to all inserted columns. 0 (default) will not copy any styles</param>
        public void MssColumn_Insert(object ssWorksheet, int ssInsertAt, int ssNumberOfColumns, int ssCopyStylesFrom)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.InsertColumn(ssInsertAt, ssNumberOfColumns, ssCopyStylesFrom);
        } // MssColumn_Insert

        /// <summary>
        /// Delete row(s) from a worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssStartRowNumber">Row number where to start deleting rows.</param>
        /// <param name="ssNumberOfRows">The number of rows to delete. Default = 1.</param>
        public void MssRow_Delete(object ssWorksheet, int ssStartRowNumber, int ssNumberOfRows)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            // Delete all comments from cells in row(s) before deleting the row(s).
            // Considers the columns in the dimension of the worksheet to prevent unnecessary processing.
            int nrColumns = ws.Dimension?.Columns ?? 0;
            RemoveCommentsInRange(ws, ssStartRowNumber, ssStartRowNumber + ssNumberOfRows - 1, 1, nrColumns);

            ws.DeleteRow(ssStartRowNumber, ssNumberOfRows);
        } // MssRow_Delete

        /// <summary>
        /// Un-Merge cells in the range provided
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssRangeToUnmerge">The range of cell to un-merge</param>
        public void MssCell_UnMerge(object ssWorksheet, RCRangeRecord ssRangeToUnmerge)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ws.Cells[ssRangeToUnmerge.ssSTRange.ssStartRow, ssRangeToUnmerge.ssSTRange.ssStartCol, ssRangeToUnmerge.ssSTRange.ssEndRow, ssRangeToUnmerge.ssSTRange.ssEndCol].Merge = false;
        } // MssCell_UnMerge

        /// <summary>
        /// Merge cells in the range provided
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssRangeToMerge">The range of the cells to merge</param>
        public void MssCell_Merge(object ssWorksheet, RCRangeRecord ssRangeToMerge)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ws.Cells[ssRangeToMerge.ssSTRange.ssStartRow, ssRangeToMerge.ssSTRange.ssStartCol, ssRangeToMerge.ssSTRange.ssEndRow, ssRangeToMerge.ssSTRange.ssEndCol].Merge = true;
        } // MssCell_Merge

        /// <summary>
        /// Find all cells that contain the specified value in the given worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet in which to search</param>
        /// <param name="ssValueToFind">The value to search for</param>
        /// <param name="ssListOfCells">List of cells (ranges) where the value has been found</param>
        public void MssCells_FindByValue(object ssWorksheet, string ssValueToFind, out RLRangeRecordList ssListOfCells)
        {
            if (string.IsNullOrEmpty(ssValueToFind))
            {
                throw new ArgumentException("Cannot search for an undefined value!");
            }

            ssListOfCells = new RLRangeRecordList();

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            List<ExcelRangeBase> result = ws.Cells.Where(c => c.Value?.ToString() == ssValueToFind).ToList();

            foreach (ExcelRangeBase item in result)
            {
                RCRangeRecord rc = new RCRangeRecord();
                rc.ssSTRange.ssStartRow = item.Start.Row;
                rc.ssSTRange.ssStartCol = item.Start.Column;

                ssListOfCells.Add(rc);
            }
        } // MssCells_FindByValue

        /// <summary>
        /// Get a list of all the conditional formatting rules in a worksheet.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssListOfRule">List of conditional formatting rules</param>
        public void MssConditionalFormatting_GetAllRules(object ssWorksheet, out RLConditionalFormatItemRecordList ssListOfRule)
        {
            ssListOfRule = new RLConditionalFormatItemRecordList();

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            foreach (var item in ws.ConditionalFormatting)
            {
                RCConditionalFormatItemRecord newItem = new RCConditionalFormatItemRecord();
                newItem.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssAddress = item.Address.Address;
                newItem.ssSTConditionalFormatItem.ssPriority = item.Priority;
                newItem.ssSTConditionalFormatItem.ssStopIfTrue = item.StopIfTrue;
                switch (item.Type)
                {
                    case eExcelConditionalFormattingRuleType.AboveAverage:
                        var aboveAverage = item as IExcelConditionalFormattingAverageGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)aboveAverage.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                        var aboveOrEqualAverage = item as IExcelConditionalFormattingAverageGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)aboveOrEqualAverage.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.BelowAverage:
                        var belowAverage = item as IExcelConditionalFormattingAverageGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)belowAverage.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                        var belowOrEqualAverage = item as IExcelConditionalFormattingAverageGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)belowOrEqualAverage.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.AboveStdDev:
                        var aboveStdDev = item as IExcelConditionalFormattingStdDevGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)aboveStdDev.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.BelowStdDev:
                        var belowStdDev = item as IExcelConditionalFormattingStdDevGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)belowStdDev.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.Bottom:
                        var bottom = item as IExcelConditionalFormattingRule;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)bottom.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.BottomPercent:
                        var bottomPercent = item as IExcelConditionalFormattingRule;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)bottomPercent.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.Top:
                        var top = item as IExcelConditionalFormattingRule;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)top.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.TopPercent:
                        var topPercent = item as IExcelConditionalFormattingRule;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)topPercent.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.Last7Days:
                        var last7Days = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)last7Days.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.LastMonth:
                        var lastMonth = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)lastMonth.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.LastWeek:
                        var lastWeek = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)lastWeek.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.NextMonth:
                        var nextMonth = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)nextMonth.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.NextWeek:
                        var nextWeek = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)nextWeek.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.ThisMonth:
                        var thisMonth = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)thisMonth.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.ThisWeek:
                        var thisWeek = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)thisWeek.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.Today:
                        var today = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)today.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.Tomorrow:
                        var tomorrow = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)tomorrow.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.Yesterday:
                        var yesterday = item as IExcelConditionalFormattingTimePeriodGroup;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)yesterday.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.BeginsWith:
                        var beginsWith = item as IExcelConditionalFormattingBeginsWith;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)beginsWith.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = beginsWith.Text;
                        break;
                    case eExcelConditionalFormattingRuleType.Between:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var rBetween = item as IExcelConditionalFormattingBetween;
                        if (rBetween != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssFormula = rBetween.Formula;
                            newItem.ssSTConditionalFormatItem.ssFormula2 = rBetween.Formula2;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.ContainsBlanks:
                        var containsBlanks = item as IExcelConditionalFormattingContainsBlanks;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)containsBlanks.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.ContainsErrors:
                        var containsErrors = item as IExcelConditionalFormattingContainsErrors;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)containsErrors.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.ContainsText:
                        var containsText = item as IExcelConditionalFormattingContainsText;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)containsText.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = containsText.Text;
                        break;
                    case eExcelConditionalFormattingRuleType.DuplicateValues:
                        var duplicateValues = item as IExcelConditionalFormattingDuplicateValues;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)duplicateValues.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.EndsWith:
                        var endsWith = item as IExcelConditionalFormattingEndsWith;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)endsWith.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = endsWith.Text;
                        break;
                    case eExcelConditionalFormattingRuleType.Equal:
                        var equal = item as IExcelConditionalFormattingEqual;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)equal.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = equal.Formula;
                        break;
                    case eExcelConditionalFormattingRuleType.Expression:
                        var expression = item as IExcelConditionalFormattingExpression;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)expression.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = expression.Formula;
                        break;
                    case eExcelConditionalFormattingRuleType.GreaterThan:
                        var gt = item as IExcelConditionalFormattingGreaterThan;
                        newItem.ssSTConditionalFormatItem.ssFormula = gt.Formula;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)gt.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                        var gte = item as IExcelConditionalFormattingGreaterThanOrEqual;
                        newItem.ssSTConditionalFormatItem.ssFormula = gte.Formula;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)gte.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.LessThan:
                        var lt = item as IExcelConditionalFormattingLessThan;
                        newItem.ssSTConditionalFormatItem.ssFormula = lt.Formula;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)lt.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                        var lte = item as IExcelConditionalFormattingLessThanOrEqual;
                        newItem.ssSTConditionalFormatItem.ssFormula = lte.Formula;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)lte.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.NotBetween:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var rNotBetween = item as IExcelConditionalFormattingBetween;
                        if (rNotBetween != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssFormula = rNotBetween.Formula;
                            newItem.ssSTConditionalFormatItem.ssFormula2 = rNotBetween.Formula2;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.NotContains:
                        break;
                    case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                        var notContainsBlanks = item as IExcelConditionalFormattingNotContainsBlanks;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)notContainsBlanks.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.NotContainsErrors:
                        var notContainsErrors = item as IExcelConditionalFormattingNotContainsErrors;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)notContainsErrors.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.NotContainsText:
                        var notContainsText = item as IExcelConditionalFormattingNotContainsText;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)notContainsText.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = notContainsText.Text;
                        break;
                    case eExcelConditionalFormattingRuleType.NotEqual:
                        var notEqual = item as IExcelConditionalFormattingNotEqual;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)notEqual.Type;
                        newItem.ssSTConditionalFormatItem.ssFormula = notEqual.Formula;
                        break;
                    case eExcelConditionalFormattingRuleType.UniqueValues:
                        var uniqueValues = item as IExcelConditionalFormattingUniqueValues;
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)uniqueValues.Type;
                        break;
                    case eExcelConditionalFormattingRuleType.ThreeColorScale:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var rcs3 = item as IExcelConditionalFormattingThreeColorScale;
                        if (rcs3 != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssLowColor = ToHexColor(rcs3.LowValue.Color);
                            newItem.ssSTConditionalFormatItem.ssMidColor = ToHexColor(rcs3.MiddleValue.Color);
                            newItem.ssSTConditionalFormatItem.ssHighColor = ToHexColor(rcs3.HighValue.Color);
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.TwoColorScale:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var rcs2 = item as IExcelConditionalFormattingTwoColorScale;
                        if (rcs2 != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssLowColor = ToHexColor(rcs2.LowValue.Color);
                            newItem.ssSTConditionalFormatItem.ssHighColor = ToHexColor(rcs2.HighValue.Color);
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.ThreeIconSet:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var icon3 = item as IExcelConditionalFormattingIconSetGroup<eExcelconditionalFormatting3IconsSetType>;
                        if (icon3 != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssIconSetStyle = icon3.IconSet.ToString();
                            newItem.ssSTConditionalFormatItem.ssReverse = icon3.Reverse;
                            newItem.ssSTConditionalFormatItem.ssShowValue = icon3.ShowValue;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.FourIconSet:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var icon4 = item as IExcelConditionalFormattingIconSetGroup<eExcelconditionalFormatting4IconsSetType>;
                        if (icon4 != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssIconSetStyle = icon4.IconSet.ToString();
                            newItem.ssSTConditionalFormatItem.ssReverse = icon4.Reverse;
                            newItem.ssSTConditionalFormatItem.ssShowValue = icon4.ShowValue;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.FiveIconSet:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        var icon5 = item as IExcelConditionalFormattingIconSetGroup<eExcelconditionalFormatting5IconsSetType>;
                        if (icon5 != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssIconSetStyle = icon5.IconSet.ToString();
                            newItem.ssSTConditionalFormatItem.ssReverse = icon5.Reverse;
                            newItem.ssSTConditionalFormatItem.ssShowValue = icon5.ShowValue;
                        }
                        break;
                    case eExcelConditionalFormattingRuleType.DataBar:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        break;
                    default:
                        break;
                }

                ssListOfRule.Add(newItem);
            }

        } // MssConditionalFormatting_GetAllRules

        /// <summary>
        /// Add a rule for conditionally formatting a range of cells.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssConditionalFormatRecord">The conditional formatting to apply to the Address Range</param>        
        public void MssConditionalFormatting_AddRule(object ssWorksheet, RCConditionalFormatItemRecord ssConditionalFormatRecord)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ExcelAddress address;
            if (string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssAddress))
            {
                address = new ExcelAddress(
                    ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssRow,
                    ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssColumn,
                    ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssRow,
                    ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssColumn);
            }
            else
            {
                address = new ExcelAddress(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssAddress);
            }
            eExcelConditionalFormattingRuleType ruleType = (eExcelConditionalFormattingRuleType)ssConditionalFormatRecord.ssSTConditionalFormatItem.ssRuleType;

            switch (ruleType)
            {
                case eExcelConditionalFormattingRuleType.AboveAverage:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddAboveAverage(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddAboveOrEqualAverage(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddBelowAverage(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddBelowOrEqualAverage(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddAboveStdDev(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddBelowStdDev(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Bottom:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddBottom(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddBottomPercent(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Top:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddTop(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.TopPercent:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddTopPercent(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Last7Days:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddLast7Days(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.LastMonth:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddLastMonth(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.LastWeek:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddLastWeek(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NextMonth:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddNextMonth(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NextWeek:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddNextWeek(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.ThisMonth:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddThisMonth(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.ThisWeek:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddThisWeek(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Today:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddToday(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Tomorrow:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddTomorrow(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Yesterday:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddYesterday(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.BeginsWith:
                    var beginsWith = ws.ConditionalFormatting.AddBeginsWith(address);
                    beginsWith.Text = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(beginsWith, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Between:
                    var between = ws.ConditionalFormatting.AddBetween(address);
                    between.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    between.Formula2 = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula2;
                    ApplyRuleCommon(between, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddContainsBlanks(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddContainsErrors(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.ContainsText:
                    var containsText = ws.ConditionalFormatting.AddContainsText(address);
                    containsText.Text = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(containsText, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddDuplicateValues(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.EndsWith:
                    var endsWith = ws.ConditionalFormatting.AddEndsWith(address);
                    endsWith.Text = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(endsWith, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Equal:
                    var equal = ws.ConditionalFormatting.AddEqual(address);
                    equal.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(equal, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.Expression:
                    var expression = ws.ConditionalFormatting.AddExpression(address);
                    expression.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(expression, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.GreaterThan:
                    var gt = ws.ConditionalFormatting.AddGreaterThan(address);
                    gt.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(gt, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                    var gte = ws.ConditionalFormatting.AddGreaterThanOrEqual(address);
                    gte.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(gte, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.LessThan:
                    var lt = ws.ConditionalFormatting.AddLessThan(address);
                    lt.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(lt, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                    var lte = ws.ConditionalFormatting.AddLessThanOrEqual(address);
                    lte.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(lte, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NotBetween:
                    var notBetween = ws.ConditionalFormatting.AddNotBetween(address);
                    notBetween.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    notBetween.Formula2 = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula2;
                    ApplyRuleCommon(notBetween, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NotContains:
                    throw new NotSupportedException("ConditionalFormatting rule type 'NotContains' is not yet supported by this extension.");
                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddNotContainsBlanks(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddNotContainsErrors(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NotContainsText:
                    var notContainsText = ws.ConditionalFormatting.AddNotContainsText(address);
                    notContainsText.Text = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(notContainsText, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.NotEqual:
                    var notEqual = ws.ConditionalFormatting.AddNotEqual(address);
                    notEqual.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    ApplyRuleCommon(notEqual, ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.UniqueValues:
                    ApplyRuleCommon(ws.ConditionalFormatting.AddUniqueValues(address), ssConditionalFormatRecord);
                    break;
                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    var cs3 = ws.ConditionalFormatting.AddThreeColorScale(address);
                    if (!string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssLowColor))
                        cs3.LowValue.Color = Util.ConvertFromColorCode(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssLowColor);
                    if (!string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssMidColor))
                        cs3.MiddleValue.Color = Util.ConvertFromColorCode(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssMidColor);
                    if (!string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssHighColor))
                        cs3.HighValue.Color = Util.ConvertFromColorCode(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssHighColor);
                    cs3.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    cs3.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    break;
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    var cs2 = ws.ConditionalFormatting.AddTwoColorScale(address);
                    if (!string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssLowColor))
                        cs2.LowValue.Color = Util.ConvertFromColorCode(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssLowColor);
                    if (!string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssHighColor))
                        cs2.HighValue.Color = Util.ConvertFromColorCode(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssHighColor);
                    cs2.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    cs2.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    break;
                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    var i3 = ws.ConditionalFormatting.AddThreeIconSet(address,
                        ParseEnum(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssIconSetStyle,
                                  eExcelconditionalFormatting3IconsSetType.TrafficLights1));
                    i3.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    i3.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    i3.Reverse = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssReverse;
                    i3.ShowValue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssShowValue;
                    break;
                case eExcelConditionalFormattingRuleType.FourIconSet:
                    var i4 = ws.ConditionalFormatting.AddFourIconSet(address,
                        ParseEnum(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssIconSetStyle,
                                  eExcelconditionalFormatting4IconsSetType.Arrows));
                    i4.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    i4.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    i4.Reverse = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssReverse;
                    i4.ShowValue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssShowValue;
                    break;
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    var i5 = ws.ConditionalFormatting.AddFiveIconSet(address,
                        ParseEnum(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssIconSetStyle,
                                  eExcelconditionalFormatting5IconsSetType.Arrows));
                    i5.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    i5.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    i5.Reverse = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssReverse;
                    i5.ShowValue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssShowValue;
                    break;
                case eExcelConditionalFormattingRuleType.DataBar:
                    System.Drawing.Color dataBarColor = string.IsNullOrEmpty(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssDataBarColor)
                        ? System.Drawing.ColorTranslator.FromHtml("#638EC6")
                        : Util.ConvertFromColorCode(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssDataBarColor);
                    var dataBar = ws.ConditionalFormatting.AddDatabar(address, dataBarColor);
                    dataBar.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    dataBar.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    dataBar.ShowValue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssShowValue;
                    break;
                default:
                    throw new ArgumentException("Invalid Rule Type: " + ssConditionalFormatRecord.ssSTConditionalFormatItem.ssRuleType);
            }

        } // MssConditionalFormatting_AddRule

        /// <summary>
        /// Inserts a new row into the spreadsheet.  Existing rows below the position are shifted down.  All formula are updated to take account of the new row.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to insert the row(s) into</param>
        /// <param name="ssInsertAt">The position of the new row
        /// </param>
        /// <param name="ssNrRows">Number of rows to insert</param>
        /// <param name="ssCopyStyleFromRow">Copy Styles from this row. Applied to all inserted rows. 0 will not copy any styles</param>
        public void MssRow_Insert(object ssWorksheet, int ssInsertAt, int ssNrRows, int ssCopyStyleFromRow)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ws.InsertRow(ssInsertAt, ssNrRows, ssCopyStyleFromRow);
        } // MssRow_Insert

        /// <summary>
        /// Apply a specified cell format to the range specified for the given worksheet
        /// </summary>
        /// <param name="ssWorksheet">Worksheet object where formatting is to be applied</param>
        /// <param name="ssCellFormat">CellFormat to apply</param>
        /// <param name="ssRange">Range that CellFormat is to be applied to</param>
        public void MssCellFormat_ApplyToRange(object ssWorksheet, RCCellFormatRecord ssCellFormat, RCRangeRecord ssRange)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ExcelRange er = ws.Cells[ssRange.ssSTRange.ssStartRow, ssRange.ssSTRange.ssStartCol, ssRange.ssSTRange.ssEndRow, ssRange.ssSTRange.ssEndCol];

            Util.ApplyFormatToRange(er, ssCellFormat);
        } // MssCellFormat_ApplyToRange

        /// <summary>
        /// Hide / Show a worksheet
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssHidden">Visible = 0 - The worksheet is visible
        /// Hidden = 1 - The worksheet is hidden but can be shown by the user via the user interface
        /// VeryHidden = 2 - The worksheet is hidden and cannot be shown by the user via the user interface</param>
        public void MssWorksheet_Hide_Show(object ssWorksheet, int ssHidden)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Hidden = (eWorkSheetHidden)ssHidden;
        } // MssWorksheet_Hide_Show

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
        public void MssContainInRange(object ssWorksheet, string ssRange, string ssValue, string ssParameter1, out bool ssFound, out int ssRowIndex, out int ssColumnIndex)
        {
            ssFound = false;
            ssRowIndex = 0;
            ssColumnIndex = 0;
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            foreach (var item in ws.Cells[ssRange])
            {
                if (item.Value?.ToString() == ssValue)
                {
                    ssRowIndex = item.Start.Row;
                    ssColumnIndex = item.Start.Column;
                    ssFound = true;
                    break;
                }
            }

        } // MssContainInRange

        /// <summary>
        /// Hides / Shows Row passed by index
        /// </summary>
        /// <param name="ssWorksheet">Worksheet to work with</param>
        /// <param name="ssRowIndex">Index of the Row to show/hide</param>
        /// <param name="ssHidden">True = Hide, False = Show</param>
        public void MssRow_Hide_Show(object ssWorksheet, int ssRowIndex, bool ssHidden)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Row(ssRowIndex).Hidden = ssHidden;
        } // MssRow_Hide_Show

        /// <summary>
        /// Calculate all formulae for the entire workbook provided.
        /// </summary>
        /// <param name="ssWorkbook">The workbook to work with</param>
        public void MssWorkbook_Calculate(object ssWorkbook)
        {
            ExcelPackage p = ssWorkbook as ExcelPackage;
            ExcelWorkbook wb = p.Workbook;

            wb.Calculate();
        } // MssWorkbook_Calculate

        /// <summary>
        /// Calculate all formulae on the provided worksheet.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        public void MssWorksheet_Calculate(object ssWorksheet)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Calculate();
        } // MssWorksheet_Calculate

        /// <summary>
        /// Hides / Shows Column passed by index
        /// </summary>
        /// <param name="ssWorksheet">The worksheet you want to work with.</param>
        /// <param name="ssColumn">The index of the column within the worksheet that you want to hide/show.</param>
        /// <param name="ssHidden">A Boolean value, set to True to hide the column, and to False to show the column.</param>
        public void MssColumn_Hide_Show(object ssWorksheet, int ssColumn, bool ssHidden)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ws.Column(ssColumn).Hidden = ssHidden;

            if (!ssHidden)
            {
                ws.Column(ssColumn).Width = ws.DefaultColWidth;
            }
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
        public void MssCell_Read(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, out string ssCellValue, bool ssReadText)
        {
            ssCellValue = "";

            if (string.IsNullOrEmpty(ssCellName) && ssCellColumn <= 0 && ssCellRow <= 0)
            {
                throw new ArgumentException("You need to specify a valid cell name (i.e. A4) or cell index (row/column combination)");
            }

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            try
            {
                ExcelRange cell;

                if (string.IsNullOrEmpty(ssCellName))
                {
                    cell = ws.Cells[ssCellRow, ssCellColumn];
                }
                else
                {
                    cell = ws.Cells[ssCellName];
                }

                ssCellValue = ssReadText ? cell.Text : Convert.ToString(cell.Value);
            }
            catch (Exception ex)
            {
                Util.LogMessage("MssCell_Read failed (name='" + ssCellName + "', row=" + ssCellRow + ", col=" + ssCellColumn + "): " + ex.Message);
                ssCellValue = String.Empty;
            }
        } // MssCell_Read

        /// <summary>
        /// Write a converted value to a cell, defined by its index.
        /// Input is a worksheet-object.
        /// Accepts format for the target cell
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides </param>
        /// <param name="ssCellName">Name of the cell to write to, i.e. A4</param>
        /// <param name="ssCellRow">Row number of the cell to write to</param>
        /// <param name="ssCellColumn">Column number of the cell to write to</param>
        /// <param name="ssCellValue">The value to write to the cell</param>
        /// <param name="ssCellType">Type can be:
        /// general (default if empty)
        /// text,
        /// datetime,
        /// integer,
        /// decimal,
        /// boolean,
        /// formula</param>
        /// <param name="ssPreserveFormat">Default value is False. If set to True, the CellFormat parameter is ignored.</param>
        /// <param name="ssCellFormat">CellFormat for the target cell</param>
        public void MssCell_Write(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, string ssCellValue, string ssCellType, bool ssPreserveFormat, RCCellFormatRecord ssCellFormat)
        {
            if (string.IsNullOrEmpty(ssCellName) && ssCellRow < 1 && ssCellColumn < 1)
            {
                throw new ArgumentException("You need to specify a valid cell name (i.e. A4) or cell index (row/column combination)");
            }

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ExcelAddress address = null;

            if (!string.IsNullOrEmpty(ssCellName))
            {
                address = new ExcelAddress(ssCellName);
            }
            if (ssCellColumn >= 1 && ssCellRow >= 1)
            {
                address = new ExcelAddress(ssCellRow, ssCellColumn, ssCellRow, ssCellColumn);
            }

            if (address == null)
            {
                throw new ArgumentException("Invalid address");
            }

            switch (ssCellType.ToLower())
            {
                case "integer": ws.SetValue(address.Address, Convert.ToInt32(ssCellValue, System.Globalization.CultureInfo.InvariantCulture)); break;
                case "datetime": ws.SetValue(address.Address, Convert.ToDateTime(ssCellValue, System.Globalization.CultureInfo.InvariantCulture)); break;
                case "decimal": ws.SetValue(address.Address, Convert.ToDecimal(ssCellValue, System.Globalization.CultureInfo.InvariantCulture)); break;
                case "boolean": ws.SetValue(address.Address, Convert.ToBoolean(ssCellValue, System.Globalization.CultureInfo.InvariantCulture)); break;
                case "formula": ws.Cells[address.Address].Formula = ssCellValue.TrimStart('='); break;
                case "text":
                    ssCellFormat.ssSTCellFormat.ssNumberFormat = "@"; /// Formats the cell as text. Ref: https://stackoverflow.com/a/30095442
                    ws.SetValue(address.Address, ssCellValue);
                    if (ssPreserveFormat)
                    {
                        // ApplyFormatToRange is skipped below when preserving formatting,
                        // so apply the text number format directly to the cell.
                        ws.Cells[address.Address].Style.Numberformat.Format = "@";
                    }
                    break;
                default: ws.SetValue(address.Address, ssCellValue); break;
            }

            if (!ssPreserveFormat)
            {
                Util.ApplyFormatToRange(ws.Cells[address.Address], ssCellFormat);
            }

        } // MssCell_Write

        /// <summary>
        /// Change the index of a worksheet in the document
        /// </summary>
        /// <param name="ssWorkbook">The workbook in which the change is to be made.</param>
        /// <param name="ssCurrentIndex">The current index(position) of the sheet in question</param>
        /// <param name="ssNewIndex">The new index for the sheet</param>
        public void MssWorkbook_ChangeSheetIndex(object ssWorkbook, int ssCurrentIndex, int ssNewIndex)
        {
            if (ssCurrentIndex <= 0 || ssNewIndex <= 0)
            {
                throw new ArgumentException("Current and New index values must be >= 1.");
            }

            ExcelPackage ee = ssWorkbook as ExcelPackage;

            if (ssNewIndex > ee.Workbook.Worksheets.Count)
            {
                ee.Workbook.Worksheets.MoveToEnd(ssCurrentIndex);
                return;
            }
            ee.Workbook.Worksheets.MoveBefore(ssCurrentIndex, ssNewIndex);
        } // MssWorkbook_ChangeSheetIndex

        /// <summary>
        /// Select a worksheet by its index or by its name
        /// </summary>
        /// <param name="ssWorkbook">The workbook wherein the worksheet exists</param>
        /// <param name="ssWorksheetIndex">The index of the worksheet to find. Indexes start at 1</param>
        /// <param name="ssWorksheetName">The name of the worksheet to find</param>
        /// <param name="ssWorksheet">This is the worksheet object that you have been looking for,</param>
        public void MssWorksheet_Select(object ssWorkbook, int ssWorksheetIndex, string ssWorksheetName, out object ssWorksheet)
        {
            if (ssWorksheetIndex <= 0 && string.IsNullOrEmpty(ssWorksheetName))
            {
                throw new ArgumentException("You need to specify at least one of WorksheetIndex or WorksheetName");
            }

            ssWorksheet = null;

            ExcelPackage p = ssWorkbook as ExcelPackage;
            ExcelWorkbook wb = p.Workbook;
            ExcelWorksheet ws;

            if (ssWorksheetIndex > 0)
            {
                ws = p.Workbook.Worksheets[ssWorksheetIndex];
                ssWorksheet = ws;
                return;
            }

            ws = p.Workbook.Worksheets[ssWorksheetName];
            ssWorksheet = ws;

        } // MssWorksheet_Select

        /// <summary>
        /// Delete a worksheet in a workbook by specifying either the index, or the name of the worksheet.
        /// </summary>
        /// <param name="ssWorkbook">The workbook from which you want to delete the worksheet</param>
        /// <param name="ssIndexToDelete">The index of the worksheet to delete. Set to 0 if using the worksheet name to delete</param>
        /// <param name="ssNameToDelete">The name of the worksheet to delete. Set to empty string "" if using the index to delete.</param>
        public void MssWorksheet_Delete(object ssWorkbook, int ssIndexToDelete, string ssNameToDelete)
        {
            if (ssIndexToDelete <= 0 && string.IsNullOrEmpty(ssNameToDelete))
            {
                throw new ArgumentException("You need to specify at least one of WorksheetIndex or WorksheetName");
            }

            ExcelPackage ee = ssWorkbook as ExcelPackage;
            if (ssIndexToDelete == 0)
            {
                ee.Workbook.Worksheets.Delete(ssNameToDelete);
            }
            if (ssIndexToDelete > 0)
            {
                ee.Workbook.Worksheets.Delete(ssIndexToDelete);
            }
        }

        /// <summary>
        /// Add a worksheet to an existing workbook, optionally at the index specified. Specifying only a name will create a blank sheet. 
        /// Specifying  a name with binary data, will add the sheet from the existing binary data, and then rename to the newly provided name
        /// </summary>
        /// <param name="ssWorkbook">The workbook that you want to add the sheet to</param>
        /// <param name="ssWorksheetName">The name of the worksheet you want to add. If binary data is nullbinary(), an empty sheet will be added</param>
        /// <param name="ssWorksheet">The worksheet object that you want to add. Set to nullbinary() if adding a new sheet by name</param>
        /// <param name="ssIndexWhereToAdd">The index where to add the new sheet. Default will be highest sheet index plus 1</param>
        public void MssWorkBook_AddSheet(object ssWorkbook, string ssWorksheetName, object ssWorksheet, int ssIndexWhereToAdd)
        {
            ExcelPackage ee = ssWorkbook as ExcelPackage;
            ExcelWorksheet ws;
            ExcelWorksheet newSheet;

            if (ssWorksheet != null)
            {
                ws = AsWorksheet(ssWorksheet);
                newSheet = ee.Workbook.Worksheets.Add(string.IsNullOrEmpty(ssWorksheetName) ? "Copy_" + ws.Name : ssWorksheetName, ws);
                if (ssIndexWhereToAdd > 0)
                {
                    MssWorkbook_ChangeSheetIndex((object)ee, newSheet.Index, ssIndexWhereToAdd);
                }
                return;
            }

            newSheet = ee.Workbook.Worksheets.Add(string.IsNullOrEmpty(ssWorksheetName) ? "Sheet1" : ssWorksheetName);
            if (ssIndexWhereToAdd > 0)
            {
                MssWorkbook_ChangeSheetIndex((object)ee, newSheet.Index, ssIndexWhereToAdd);
            }
        } // MssWorkBook_AddSheet

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssProperties"></param>
        public void MssWorkbook_GetProperties(object ssWorkbook, out RCWorkbookRecord ssProperties)
        {
            ssProperties = new RCWorkbookRecord();
            ssProperties.ssSTWorkbook.ssWorksheets = new RLWorksheetRecordList();

            ExcelPackage p = ssWorkbook as ExcelPackage;
            ExcelWorkbook wb = p.Workbook;

            foreach (var sheet in wb.Worksheets)
            {
                RCWorksheetRecord newSheet;
                MssWorksheet_GetProperties(sheet, out newSheet);

                ssProperties.ssSTWorkbook.ssWorksheets.Add(newSheet);
            }
        } // MssWorkbook_GetProperties

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssProperties"></param>
        public void MssWorksheet_GetProperties(object ssWorksheet, out RCWorksheetRecord ssProperties)
        {
            ssProperties = new RCWorksheetRecord();

            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            ssProperties.ssSTWorksheet.ssIndex = ws.Index;
            ssProperties.ssSTWorksheet.ssName = ws.Name;

            ssProperties.ssSTWorksheet.ssDimension = Util.CastDimension(ws.Dimension);

            Color c = ws.TabColor;
            RCColorRecord rc = new RCColorRecord();

            rc.ssSTColor.ssA = c.A;
            rc.ssSTColor.ssB = c.B;
            rc.ssSTColor.ssG = c.G;
            rc.ssSTColor.ssR = c.R;
            rc.ssSTColor.ssIsKnownColor = c.IsKnownColor;
            rc.ssSTColor.ssIsNamedColor = c.IsNamedColor;
            rc.ssSTColor.ssIsSystemColor = c.IsSystemColor;
            rc.ssSTColor.ssName = c.Name;

            ssProperties.ssSTWorksheet.ssTabColor = rc;
        } // MssWorksheet_GetProperties

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssWorksheetName"></param>
        public void MssWorksheet_GetName(object ssWorksheet, out string ssWorksheetName)
        {
            ssWorksheetName = (AsWorksheet(ssWorksheet)).Name;
        } // MssWorksheet_GetName

        /// <summary>
        /// Write a dataset to a range of cells.
        /// Accepts format for the taget cells
        /// </summary>
        /// <param name="ssWorksheet">Worksheet to write to</param>
        /// <param name="ssRowStart">Start row (integer)</param>
        /// <param name="ssColumnStart">Start column (integer)</param>
        /// <param name="ssDataSet">Data to write</param>
        /// <param name="ssCellFormat">CellFormat for the target cell</param>
        /// <param name="ssExportHeaders">True if headers should be included in export</param>
        public void MssCell_WriteRange(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord ssCellFormat, bool ssExportHeaders)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            DataTable dt = Util.RecordListToDataTable((RecordList)ssDataSet);

            if (dt != null)
            {
                // Strip the OutSystems "ss" attribute-name prefix so exported headers read cleanly.
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (dt.Columns[i].ColumnName.StartsWith("ss", StringComparison.CurrentCulture))
                    {
                        dt.Columns[i].ColumnName = dt.Columns[i].ColumnName.Substring(2);
                    }
                }

                ExcelRange range = (ExcelRange)ws.Cells[ssRowStart, ssColumnStart].LoadFromDataTable(dt, ssExportHeaders);

                Util.ApplyFormatToRange(range, ssCellFormat);
            }
        } // MssCell_WriteRangeWithFormat

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        public void MssWorkbook_Close(object ssWorkbook)
        {
            ExcelPackage p = ssWorkbook as ExcelPackage;
            p.Dispose();
        } // MssWorkbook_Close

        /// <summary>
        /// Rename a worksheet
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssName"></param>
        public void MssWorksheet_Rename(object ssWorksheet, string ssName)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Name = ssName;
        } // MssWorksheet_Rename

        /// <summary>
        /// Get the in-memory binary data of the specified workbook
        /// </summary>
        /// <param name="ssWorkbook">The workbook you want the binary data for</param>
        /// <param name="ssBinaryData">The binary data of the file</param>
        public void MssWorkbook_GetBinaryData(object ssWorkbook, out byte[] ssBinaryData)
        {
            ExcelPackage p = ssWorkbook as ExcelPackage;
            Util.PreserveVisibleRowsForZeroHeightSheets(p);
            ssBinaryData = p.GetAsByteArray();
            // GetAsByteArray closes the package; reload so the workbook stays usable.
            p.Load(new System.IO.MemoryStream(ssBinaryData));
        } // MssWorkbook_GetBinaryData

        /// <summary>
        /// Opens an existing workbook for editing by either specifying a name or the binary data.
        /// </summary>
        /// <param name="ssFileName">Location of the file that you want to open. Set to empty string "" when using binary data</param>
        /// <param name="ssBinary_Data">Binary data of the file that you want to open. Set to nullbinary() if using FileName</param>
        /// <param name="ssWorkbook">The workbook that you want to work with.</param>
        public void MssWorkbook_Open(string ssFileName, byte[] ssBinary_Data, out object ssWorkbook)
        {
            bool hasBinaryData = ssBinary_Data != null && ssBinary_Data.LongLength > 0;

            if (!hasBinaryData && string.IsNullOrEmpty(ssFileName))
            {
                throw new ArgumentException("You need to specify at least one of FileName or Binary_Data");
            }

            ExcelPackage p = new ExcelPackage();
            if (ssFileName.ToLower().StartsWith("http:") || ssFileName.ToLower().StartsWith("https:"))
            {
                const int timeoutMs = 30000;                 // fail fast instead of hanging on a slow/dead host
                const long maxDownloadBytes = 100L * 1024 * 1024; // cap the download to guard against runaway responses

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ssFileName);
                request.Timeout = timeoutMs;
                request.ReadWriteTimeout = timeoutMs;

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    if (response.StatusCode != HttpStatusCode.OK)
                    {
                        throw new IOException("Failed to download workbook from '" + ssFileName + "': HTTP " + (int)response.StatusCode + " " + response.StatusCode + ".");
                    }

                    using (Stream responseStream = response.GetResponseStream())
                    using (MemoryStream buffer = new MemoryStream())
                    {
                        byte[] chunk = new byte[81920];
                        long total = 0;
                        int read;
                        while ((read = responseStream.Read(chunk, 0, chunk.Length)) > 0)
                        {
                            total += read;
                            if (total > maxDownloadBytes)
                            {
                                throw new IOException("The downloaded workbook exceeds the maximum allowed size (100 MB).");
                            }
                            buffer.Write(chunk, 0, read);
                        }

                        buffer.Position = 0;
                        p.Load(buffer);
                    }
                }
            }
            else if (!string.IsNullOrEmpty(ssFileName))
            {
                using (FileStream fs = System.IO.File.Open(ssFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read))
                {
                    p.Load(fs);
                }
            }
            else if (hasBinaryData)
            {
                Stream s = new MemoryStream(ssBinary_Data);
                p.Load(s);
            }
            else
            {
                throw new FileNotFoundException("Could not open a file with the given information. Please verify your filename/binary data and try again.");
            }

            ssWorkbook = p;
        } // MssWorkbook_Open

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssChartType">Receives the graph type in text, possible types:
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
        public void MssChart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            var chart = ws.Drawings.AddChart(ssChartName, Util.StringToChartType(ssChartType));
            chart.SetPosition(ssRowPos, 0, ssColPos, 0);
            chart.SetSize(ssWidth, ssHeight);

            for (int i = 0; i < ssDataSeries_List.Count; i++)
            {
                RCDataSeriesRecord dataSeries = ssDataSeries_List[i];
                STRangeStructure valuerange = dataSeries.ssSTDataSeries.ssValueRange.ssSTRange;
                STRangeStructure labelrange = dataSeries.ssSTDataSeries.ssLabelRange.ssSTRange;

                int val_startRow = valuerange.ssStartRow;
                int val_startCol = valuerange.ssStartCol;
                int val_endRow = valuerange.ssEndRow;
                int val_endCol = valuerange.ssEndCol;
                int lbl_startRow = labelrange.ssStartRow;
                int lbl_startCol = labelrange.ssStartCol;
                int lbl_endRow = labelrange.ssEndRow;
                int lbl_endCol = labelrange.ssEndCol;

                var series = chart.Series.Add(ExcelRange.GetAddress(val_startRow, val_startCol, val_endRow, val_endCol),
                    ExcelRange.GetAddress(lbl_startRow, lbl_startCol, lbl_endRow, lbl_endCol));
                series.Header = dataSeries.ssSTDataSeries.ssName;
            }
        } // MssChart_Create

        /// <summary>
        /// Calculates the formula of a cell, defined by its index.
        /// Input is a worksheet-object
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">row number</param>
        /// <param name="ssColumn">column number</param>
        public void MssCell_CalculateByIndex(object ssWorksheet, int ssRow, int ssColumn)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Cells[ssRow, ssColumn].Calculate();
        } // MssCell_CalculateByIndex 

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
        public void MssWorksheet_SetFooter(object ssWorksheet, string ssLeftSection, string ssCenterSection, string ssRightSection)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.HeaderFooter.OddFooter.LeftAlignedText = ssLeftSection;
            ws.HeaderFooter.OddFooter.CenteredText = ssCenterSection;
            ws.HeaderFooter.OddFooter.RightAlignedText = ssRightSection;
            ws.HeaderFooter.EvenFooter.LeftAlignedText = ssLeftSection;
            ws.HeaderFooter.EvenFooter.CenteredText = ssCenterSection;
            ws.HeaderFooter.EvenFooter.RightAlignedText = ssRightSection;
        } // MssWorksheet_SetFooter

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
        public void MssWorksheet_SetHeader(object ssWorksheet, string ssLeftSection, string ssCenterSection, string ssRightSection)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.HeaderFooter.OddHeader.LeftAlignedText = ssLeftSection;
            ws.HeaderFooter.OddHeader.CenteredText = ssCenterSection;
            ws.HeaderFooter.OddHeader.RightAlignedText = ssRightSection;
            ws.HeaderFooter.EvenHeader.LeftAlignedText = ssLeftSection;
            ws.HeaderFooter.EvenHeader.CenteredText = ssCenterSection;
            ws.HeaderFooter.EvenHeader.RightAlignedText = ssRightSection;
        } // MssWorksheet_SetHeader

        /// <summary>
        /// Get the left, center and right sections for the odd or even page header of the specified worksheet.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet from which to retrieve the header.</param>
        /// <param name="ssIsEven">If True, retrieves the even page header, otherwise the odd page header.</param>
        /// <param name="ssLeftSection">The left section of the header.</param>
        /// <param name="ssCenterSection">The center section of the header.</param>
        /// <param name="ssRightSection">The right section of the header.</param>
        public void MssWorksheet_GetHeader(object ssWorksheet, bool ssIsEven, out string ssLeftSection, out string ssCenterSection, out string ssRightSection)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ssLeftSection = (ssIsEven ? ws.HeaderFooter.EvenHeader.LeftAlignedText : ws.HeaderFooter.OddHeader.LeftAlignedText);
            ssCenterSection = (ssIsEven ? ws.HeaderFooter.EvenHeader.CenteredText : ws.HeaderFooter.OddHeader.CenteredText);
            ssRightSection = (ssIsEven ? ws.HeaderFooter.EvenHeader.RightAlignedText : ws.HeaderFooter.OddHeader.RightAlignedText);
        } // MssWorksheet_GetHeader

        /// <summary>
        /// Get the left, center and right sections for the odd or even page footer of the specified worksheet.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet for which to get the footer.</param>
        /// <param name="ssIsEven">If True, retrieves the even page footer, otherwise the odd page footer.</param>
        /// <param name="ssLeftSection">The left section of the footer.</param>
        /// <param name="ssCenterSection">The center section of the footer.</param>
        /// <param name="ssRightSection">The right section of the footer.</param>
        public void MssWorksheet_GetFooter(object ssWorksheet, bool ssIsEven, out string ssLeftSection, out string ssCenterSection, out string ssRightSection)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ssLeftSection = (ssIsEven ? ws.HeaderFooter.EvenFooter.LeftAlignedText : ws.HeaderFooter.OddFooter.LeftAlignedText);
            ssCenterSection = (ssIsEven ? ws.HeaderFooter.EvenFooter.CenteredText : ws.HeaderFooter.OddFooter.CenteredText);
            ssRightSection = (ssIsEven ? ws.HeaderFooter.EvenFooter.RightAlignedText : ws.HeaderFooter.OddFooter.RightAlignedText);
        } // MssWorksheet_GetFooter

        /// <summary>
        /// Clear value of a cell, defined by its index.
        /// Option to specify whether the cell is part of a merged group or not.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">Row Number</param>
        /// <param name="ssStartColumn">Column Number</param>
        /// <param name="ssEndColumn">Column Number. Mandatory if IsMerged is True</param>
        /// <param name="ssIsMerged">If True cells are merged</param>
        public void MssCell_ClearValueByIndex(object ssWorksheet, int ssRow, int ssStartColumn, int ssEndColumn, bool ssIsMerged)
        {
            // Select the worksheet
            ExcelWorksheet ws;
            ws = AsWorksheet(ssWorksheet);

            // Check if the specified cell is part of a merged range
            if (ssIsMerged)
            {
                if (ssEndColumn <= 0)
                {
                    throw new ArgumentException("You need to specify a valid cell column for End Column");
                }

                // Select the range of merged cells to clear
                ExcelRange mergedCells = ws.Cells[ssRow, ssStartColumn, ssRow, ssEndColumn];

                // Unmerge the cells in the range
                mergedCells.Merge = false;

                // Clear the values in the range
                mergedCells.Value = null;

                // Merge the cells again
                mergedCells.Merge = true;
            }
            else
            {
                // The specified cell is not part of a merged range
                ws.Cells[ssRow, ssStartColumn].Value = null;
            }
        } // MssCell_ClearValueByIndex

        /// <summary>
        /// Clear value clear the value of a specific cell by its name.
        /// Option to specify whether the cell is part of a merged group or not.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell-name (eg A1:B1, if cells are merged; eg A1, if single cell)</param>
        /// <param name="ssIsMerged">If True cells are merged</param>
        public void MssCell_ClearValueByName(object ssWorksheet, string ssCellName, bool ssIsMerged)
        {
            // Select the worksheet
            ExcelWorksheet ws;
            ws = AsWorksheet(ssWorksheet);

            if (ssIsMerged)
            {   // Select the range of merged cells to clear
                ExcelRange cell = ws.Cells[ssCellName];

                // Unmerge the cells in the range
                cell.Merge = false;

                // Clear the values in the range
                cell.Value = null;

                // Merge the cells again 
                cell.Merge = true;
            }

            else
            {
                ws.Cells[ssCellName].Value = null;
            }
        } // MssCell_ClearValueByName

        /// <summary>
        /// Reads formula of a cell, defined by its index.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">Row Number</param>
        /// <param name="ssColumn">Column Number</param>
        /// <param name="ssFormula">The value in the cell, as text</param>
        public void MssCell_ReadFormulaByIndex(object ssWorksheet, int ssRow, int ssColumn, out string ssFormula)
        {
            // Select the worksheet
            ExcelWorksheet ws;
            ws = AsWorksheet(ssWorksheet);

            // Get the cell
            ExcelRange cell = ws.Cells[ssRow, ssColumn];

            // Get the value of the cell containing the formula
            ssFormula = cell.Formula;
        } // MssCell_ReadFormulaByIndex

        /// <summary>
        /// Group a contiguous range of rows into a collapsible outline.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssStartRow">First row of the group (1-based)</param>
        /// <param name="ssEndRow">Last row of the group (1-based)</param>
        /// <param name="ssOutlineLevel">Outline level for the group. Defaults to 1 when &lt;= 0. Use higher values for nested groups.</param>
        /// <param name="ssCollapsed">If True, the group is created collapsed (its rows are hidden).</param>
        public void MssRow_Group(object ssWorksheet, int ssStartRow, int ssEndRow, int ssOutlineLevel, bool ssCollapsed)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            int level = ssOutlineLevel <= 0 ? 1 : ssOutlineLevel;

            for (int r = ssStartRow; r <= ssEndRow; r++)
            {
                ws.Row(r).OutlineLevel = level;
                ws.Row(r).Collapsed = ssCollapsed;
                ws.Row(r).Hidden = ssCollapsed;
            }
        } // MssRow_Group

        /// <summary>
        /// Remove outline grouping from a range of rows.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssStartRow">First row to ungroup (1-based)</param>
        /// <param name="ssEndRow">Last row to ungroup (1-based)</param>
        public void MssRow_Ungroup(object ssWorksheet, int ssStartRow, int ssEndRow)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            for (int r = ssStartRow; r <= ssEndRow; r++)
            {
                ws.Row(r).OutlineLevel = 0;
                ws.Row(r).Collapsed = false;
                ws.Row(r).Hidden = false;
            }
        } // MssRow_Ungroup

        /// <summary>
        /// Group a contiguous range of columns into a collapsible outline.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssStartColumn">First column of the group (1-based)</param>
        /// <param name="ssEndColumn">Last column of the group (1-based)</param>
        /// <param name="ssOutlineLevel">Outline level for the group. Defaults to 1 when &lt;= 0. Use higher values for nested groups.</param>
        /// <param name="ssCollapsed">If True, the group is created collapsed (its columns are hidden).</param>
        public void MssColumn_Group(object ssWorksheet, int ssStartColumn, int ssEndColumn, int ssOutlineLevel, bool ssCollapsed)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            int level = ssOutlineLevel <= 0 ? 1 : ssOutlineLevel;

            for (int c = ssStartColumn; c <= ssEndColumn; c++)
            {
                ws.Column(c).OutlineLevel = level;
                ws.Column(c).Collapsed = ssCollapsed;
                ws.Column(c).Hidden = ssCollapsed;
            }
        } // MssColumn_Group

        /// <summary>
        /// Remove outline grouping from a range of columns.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssStartColumn">First column to ungroup (1-based)</param>
        /// <param name="ssEndColumn">Last column to ungroup (1-based)</param>
        public void MssColumn_Ungroup(object ssWorksheet, int ssStartColumn, int ssEndColumn)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            for (int c = ssStartColumn; c <= ssEndColumn; c++)
            {
                ws.Column(c).OutlineLevel = 0;
                ws.Column(c).Collapsed = false;
                ws.Column(c).Hidden = false;
            }
        } // MssColumn_Ungroup

        /// <summary>
        /// Parse a string into an enum value (case-insensitive), returning a fallback when
        /// the value is empty or not a recognised member.
        /// </summary>
        private static T ParseEnum<T>(string value, T fallback) where T : struct
        {
            return (!string.IsNullOrEmpty(value) && Enum.TryParse<T>(value, true, out var parsed))
                ? parsed : fallback;
        }

        /// <summary>
        /// Cast the worksheet handle to an ExcelWorksheet, throwing a clear error instead of
        /// the cryptic NullReferenceException that a failed "as" cast would surface later.
        /// </summary>
        private static ExcelWorksheet AsWorksheet(object ssWorksheet)
        {
            if (ssWorksheet is ExcelWorksheet ws)
            {
                return ws;
            }
            throw new ArgumentException("Expected a Worksheet object (e.g. from Worksheet_Select, Worksheet_SelectByName/Index, or Workbook_AddName).", nameof(ssWorksheet));
        }

        /// <summary>
        /// Apply the priority, stop-if-true and dxf style shared by every style-based
        /// conditional-formatting rule, so each case in MssConditionalFormatting_AddRule
        /// doesn't repeat the same three lines.
        /// </summary>
        private static void ApplyRuleCommon(IExcelConditionalFormattingRule rule, RCConditionalFormatItemRecord rec)
        {
            rule.Priority = rec.ssSTConditionalFormatItem.ssPriority;
            rule.StopIfTrue = rec.ssSTConditionalFormatItem.ssStopIfTrue;
            Util.ApplyConditionalFormattingStyle(rule.Style, rec.ssSTConditionalFormatItem.ssStyle);
        }

        /// <summary>
        /// Format a colour as a #RRGGBB hex string (used when reading colour-scale stops back).
        /// </summary>
        private static string ToHexColor(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

    } // CssAdvanced_Excel

} // OutSystems.NssAdvanced_Excel

