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
using OfficeOpenXml.Style;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style.Dxf;

namespace OutSystems.NssAdvanced_Excel
{

    public class CssAdvanced_Excel : IssAdvanced_Excel
    {

        /// <summary>
        /// Un-Merge cells in the range provided
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssRangeToUnmerge">The range of cell to un-merge</param>
        public void MssCell_UnMerge(object ssWorksheet, RCRangeRecord ssRangeToUnmerge)
        {
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

            ws.Cells[ssRangeToUnmerge.ssSTRange.ssStartRow, ssRangeToUnmerge.ssSTRange.ssStartCol, ssRangeToUnmerge.ssSTRange.ssEndRow, ssRangeToUnmerge.ssSTRange.ssEndCol].Merge = false;
        } // MssCell_UnMerge

        /// <summary>
        /// Merge cells in the range provided
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssRangeToMerge">The range of the cells to merge</param>
        public void MssCell_Merge(object ssWorksheet, RCRangeRecord ssRangeToMerge)
        {
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

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
                throw new Exception("Cannot search for an undefined value!");
            }

            ssListOfCells = new RLRangeRecordList();

            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

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

            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

            foreach (var item in ws.ConditionalFormatting)
            {
                RCConditionalFormatItemRecord newItem = new RCConditionalFormatItemRecord();
                newItem.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssAddress = item.Address.Address;
                newItem.ssSTConditionalFormatItem.ssPriority = item.Priority;
                newItem.ssSTConditionalFormatItem.ssStopIfTrue = item.StopIfTrue;
                switch (item.Type)
                {
                    case eExcelConditionalFormattingRuleType.AboveAverage:
                        break;
                    case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                        break;
                    case eExcelConditionalFormattingRuleType.BelowAverage:
                        break;
                    case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                        break;
                    case eExcelConditionalFormattingRuleType.AboveStdDev:
                        break;
                    case eExcelConditionalFormattingRuleType.BelowStdDev:
                        break;
                    case eExcelConditionalFormattingRuleType.Bottom:
                        break;
                    case eExcelConditionalFormattingRuleType.BottomPercent:
                        break;
                    case eExcelConditionalFormattingRuleType.Top:
                        break;
                    case eExcelConditionalFormattingRuleType.TopPercent:
                        break;
                    case eExcelConditionalFormattingRuleType.Last7Days:
                        break;
                    case eExcelConditionalFormattingRuleType.LastMonth:
                        break;
                    case eExcelConditionalFormattingRuleType.LastWeek:
                        break;
                    case eExcelConditionalFormattingRuleType.NextMonth:
                        break;
                    case eExcelConditionalFormattingRuleType.NextWeek:
                        break;
                    case eExcelConditionalFormattingRuleType.ThisMonth:
                        break;
                    case eExcelConditionalFormattingRuleType.ThisWeek:
                        break;
                    case eExcelConditionalFormattingRuleType.Today:
                        break;
                    case eExcelConditionalFormattingRuleType.Tomorrow:
                        break;
                    case eExcelConditionalFormattingRuleType.Yesterday:
                        break;
                    case eExcelConditionalFormattingRuleType.BeginsWith:
                        break;
                    case eExcelConditionalFormattingRuleType.Between:
                        break;
                    case eExcelConditionalFormattingRuleType.ContainsBlanks:
                        break;
                    case eExcelConditionalFormattingRuleType.ContainsErrors:
                        break;
                    case eExcelConditionalFormattingRuleType.ContainsText:
                        break;
                    case eExcelConditionalFormattingRuleType.DuplicateValues:
                        break;
                    case eExcelConditionalFormattingRuleType.EndsWith:
                        break;
                    case eExcelConditionalFormattingRuleType.Equal:
                        break;
                    case eExcelConditionalFormattingRuleType.Expression:
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
                        break;
                    case eExcelConditionalFormattingRuleType.NotContains:
                        break;
                    case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                        break;
                    case eExcelConditionalFormattingRuleType.NotContainsErrors:
                        break;
                    case eExcelConditionalFormattingRuleType.NotContainsText:
                        break;
                    case eExcelConditionalFormattingRuleType.NotEqual:
                        break;
                    case eExcelConditionalFormattingRuleType.UniqueValues:
                        break;
                    case eExcelConditionalFormattingRuleType.ThreeColorScale:
                        break;
                    case eExcelConditionalFormattingRuleType.TwoColorScale:
                        break;
                    case eExcelConditionalFormattingRuleType.ThreeIconSet:
                        break;
                    case eExcelConditionalFormattingRuleType.FourIconSet:
                        break;
                    case eExcelConditionalFormattingRuleType.FiveIconSet:
                        break;
                    case eExcelConditionalFormattingRuleType.DataBar:
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
        /// </param>
        public void MssConditionalFormatting_AddRule(object ssWorksheet, RCConditionalFormatItemRecord ssConditionalFormatRecord)
        {
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
            ExcelAddress address = new ExcelAddress(ssConditionalFormatRecord.ssSTConditionalFormatItem.ssAddress.ssSTAddress.ssAddress);

            LogMessage("Rule type: " + ssConditionalFormatRecord.ssSTConditionalFormatItem.ssRuleType);

            eExcelConditionalFormattingRuleType ruleType = (eExcelConditionalFormattingRuleType)ssConditionalFormatRecord.ssSTConditionalFormatItem.ssRuleType;

            switch (ruleType)
            {
                case eExcelConditionalFormattingRuleType.AboveAverage:
                    break;
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                    break;
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    break;
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                    break;
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                    break;
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    break;
                case eExcelConditionalFormattingRuleType.Bottom:
                    break;
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    break;
                case eExcelConditionalFormattingRuleType.Top:
                    break;
                case eExcelConditionalFormattingRuleType.TopPercent:
                    break;
                case eExcelConditionalFormattingRuleType.Last7Days:
                    break;
                case eExcelConditionalFormattingRuleType.LastMonth:
                    break;
                case eExcelConditionalFormattingRuleType.LastWeek:
                    break;
                case eExcelConditionalFormattingRuleType.NextMonth:
                    break;
                case eExcelConditionalFormattingRuleType.NextWeek:
                    break;
                case eExcelConditionalFormattingRuleType.ThisMonth:
                    break;
                case eExcelConditionalFormattingRuleType.ThisWeek:
                    break;
                case eExcelConditionalFormattingRuleType.Today:
                    break;
                case eExcelConditionalFormattingRuleType.Tomorrow:
                    break;
                case eExcelConditionalFormattingRuleType.Yesterday:
                    break;
                case eExcelConditionalFormattingRuleType.BeginsWith:
                    break;
                case eExcelConditionalFormattingRuleType.Between:
                    break;
                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    break;
                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    break;
                case eExcelConditionalFormattingRuleType.ContainsText:
                    break;
                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    break;
                case eExcelConditionalFormattingRuleType.EndsWith:
                    break;
                case eExcelConditionalFormattingRuleType.Equal:
                    break;
                case eExcelConditionalFormattingRuleType.Expression:
                    break;
                case eExcelConditionalFormattingRuleType.GreaterThan:
                    var gt = ws.ConditionalFormatting.AddGreaterThan(address);
                    gt.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    gt.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    gt.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    ApplyStyle(gt.Style, ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStyle);
                    break;
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                    var gte = ws.ConditionalFormatting.AddGreaterThanOrEqual(address);
                    gte.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    gte.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    gte.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    ApplyStyle(gte.Style, ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStyle);
                    break;
                case eExcelConditionalFormattingRuleType.LessThan:
                    var lt = ws.ConditionalFormatting.AddLessThan(address);
                    lt.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    lt.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    lt.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    ApplyStyle(lt.Style, ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStyle);
                    break;
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                    var lte = ws.ConditionalFormatting.AddLessThanOrEqual(address);
                    lte.Formula = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssFormula;
                    lte.Priority = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssPriority;
                    lte.StopIfTrue = ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStopIfTrue;
                    ApplyStyle(lte.Style, ssConditionalFormatRecord.ssSTConditionalFormatItem.ssStyle);
                    break;
                case eExcelConditionalFormattingRuleType.NotBetween:
                    break;
                case eExcelConditionalFormattingRuleType.NotContains:
                    break;
                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    break;
                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    break;
                case eExcelConditionalFormattingRuleType.NotContainsText:
                    break;
                case eExcelConditionalFormattingRuleType.NotEqual:
                    break;
                case eExcelConditionalFormattingRuleType.UniqueValues:
                    break;
                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    break;
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    break;
                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    break;
                case eExcelConditionalFormattingRuleType.FourIconSet:
                    break;
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    break;
                case eExcelConditionalFormattingRuleType.DataBar:
                    break;
                default:
                    throw new Exception("Invalid Rule Type: " + ssConditionalFormatRecord.ssSTConditionalFormatItem.ssRuleType);
                    break;
            }

        } // MssConditionalFormatting_AddRule

        /// <summary>
        /// Apply conditional formatting style to rule style property
        /// </summary>
        /// <param name="style"></param>
        /// <param name="ssStyle"></param>
        private void ApplyStyle(ExcelDxfStyleConditionalFormatting style, RCConditionalFormatStyleRecord ssStyle)
        {
            ExcelUnderLineType underline = (ExcelUnderLineType)ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssUnderline;
            ExcelBorderStyle bTop = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssStyle;
            ExcelBorderStyle bBottom = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssStyle;
            ExcelBorderStyle bLeft = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssStyle;
            ExcelBorderStyle bRight = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssStyle;
            ExcelFillStyle patternType = (ExcelFillStyle)ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternType;

            style.Border.Bottom.Style = bBottom;
            style.Border.Bottom.Color.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssColor);

            style.Border.Left.Style = bLeft;
            style.Border.Left.Color.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssColor);

            style.Border.Right.Style = bRight;
            style.Border.Right.Color.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssColor);

            style.Border.Top.Style = bTop;
            style.Border.Top.Color.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssColor);

            style.Fill.BackgroundColor.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssBackgroundColor);
            style.Fill.PatternColor.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternColor);
            style.Fill.PatternType = patternType;

            style.Font.Bold = ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssBold;
            style.Font.Italic = ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssItalic;
            style.Font.Strike = ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssStrike;
            style.Font.Underline = underline;
            style.Font.Color.Color = Util.ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssColor);

            style.NumberFormat.Format = ssStyle.ssSTConditionalFormatStyle.ssNumberFormat;
        }

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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

            ExcelRange er = ws.Cells[ssRange.ssSTRange.ssStartRow, ssRange.ssSTRange.ssStartCol, ssRange.ssSTRange.ssEndRow, ssRange.ssSTRange.ssEndCol];

            ApplyFormatToRange(er, ssCellFormat);
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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
            ws.Row(ssRowIndex).Hidden = ssHidden;
        } // MssRow_Hide_Show

        /// <summary>
        /// Calculate all formulae for the entire workbook provided.
        /// </summary>
        /// <param name="ssWorkbook">The workbook to work with</param>
        public void MssWorkbook_Calculate(object ssWorkbook)
        {
            ExcelWorkbook wb = ssWorkbook as ExcelWorkbook;
            wb.Calculate();
        } // MssWorkbook_Calculate

        /// <summary>
        /// Calculate all formulae on the provided worksheet.
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        public void MssWorksheet_Calculate(object ssWorksheet)
        {
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet; ;

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

            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

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
            catch (Exception)
            {
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
        /// <param name="ssCellType">Type can by text (default), datetime, integer, decimal, boolean</param>
        /// <param name="ssCellFormat">CellFormat for the target cell</param>
        public void MssCell_Write(object ssWorksheet, string ssCellName, int ssCellRow, int ssCellColumn, string ssCellValue, string ssCellType, RCCellFormatRecord ssCellFormat)
        {
            if (string.IsNullOrEmpty(ssCellName) && ssCellRow < 1 && ssCellColumn < 1)
            {
                throw new Exception("You need to specify a valid cell name (i.e. A4) or cell index (row/column combination)");
            }

            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

            if (!string.IsNullOrEmpty(ssCellName))
            {
                switch (ssCellType.ToLower())
                {
                    case "integer": ws.SetValue(ssCellName, Convert.ToInt32(ssCellValue)); break;
                    case "datetime": ws.SetValue(ssCellName, Convert.ToDateTime(ssCellValue)); break;
                    case "decimal": ws.SetValue(ssCellName, Convert.ToDecimal(ssCellValue)); break;
                    case "boolean": ws.SetValue(ssCellName, Convert.ToBoolean(ssCellValue)); break;
                    case "formula": ws.Cells[ssCellName].Formula = ssCellValue.StartsWith("=") ? ssCellValue : string.Concat("=", ssCellValue); break;
                    default: ws.SetValue(ssCellName, ssCellValue); break;
                }

                ApplyFormatToRange(ws.Cells[ssCellName], ssCellFormat);
                return;
            }
            if (ssCellColumn >= 1 && ssCellRow >= 1)
            {
                switch (ssCellType.ToLower())
                {
                    case "integer": ws.SetValue(ssCellRow, ssCellColumn, Convert.ToInt32(ssCellValue)); break;
                    case "datetime": ws.SetValue(ssCellRow, ssCellColumn, Convert.ToDateTime(ssCellValue)); break;
                    case "decimal": ws.SetValue(ssCellRow, ssCellColumn, Convert.ToDecimal(ssCellValue)); break;
                    case "boolean": ws.SetValue(ssCellRow, ssCellColumn, Convert.ToBoolean(ssCellValue)); break;
                    case "formula": ws.Cells[ssCellRow, ssCellColumn].Formula = ssCellValue.StartsWith("=") ? ssCellValue : string.Concat("=", ssCellValue); break;
                    default: ws.SetValue(ssCellRow, ssCellColumn, ssCellValue); break;
                }

                ApplyFormatToRange(ws.Cells[ssCellRow, ssCellColumn], ssCellFormat);
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
        /// Log a message to the General Log
        /// </summary>
        /// <param name="message">What to log</param>
        void LogMessage(object message)
        {
            GenericExtendedActions.LogMessage(AppInfo.GetAppInfo().OsContext, message.ToString(), "AdvXL");
        }

        /// <summary>
        /// Delete a worksheet in a workbook by specifying either the index, or the name of the worksheet.
        /// </summary>
        /// <param name="ssWorkbook">The workbook from which you want to delete the worksheet</param>
        /// <param name="ssIndexToDelete">The index of the worksheet to delete. Set to 0 if using the worksheet name to delete</param>
        /// <param name="ssNameToDelete">The name of the worksheet to delete. Set to empty string "" if using the index to delete.</param>
        public void MssWorksheet_Delete(object ssWorkbook, int ssIndexToDelete, string ssNameToDelete)
        {
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
                ws = ssWorksheet as ExcelWorksheet;
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

            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

            ssProperties.ssSTWorksheet.ssIndex = ws.Index;
            ssProperties.ssSTWorksheet.ssName = ws.Name;

            ssProperties.ssSTWorksheet.ssDimension = CastDimension(ws.Dimension);

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
        /// <param name="dimension"></param>
        /// <returns></returns>
        private RCDimensionRecord CastDimension(ExcelAddressBase dimension)
        {
            RCDimensionRecord dim = new RCDimensionRecord();

            if (dimension == null)
            {
                return dim;
            }

            dim.ssSTDimension.ssAddress = dimension.Address;
            dim.ssSTDimension.ssColumns = dimension.Columns;
            dim.ssSTDimension.ssEnd = CastAddress(dimension.End);
            dim.ssSTDimension.ssRows = dimension.Rows;
            dim.ssSTDimension.ssStart = CastAddress(dimension.Start);

            return dim;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        private RCAddressRecord CastAddress(ExcelCellAddress address)
        {
            RCAddressRecord add = new RCAddressRecord();

            if (address == null)
            {
                return add;
            }

            add.ssSTAddress.ssAddress = address.Address;
            add.ssSTAddress.ssColumn = address.Column;
            add.ssSTAddress.ssIsRef = address.IsRef;
            add.ssSTAddress.ssRow = address.Row;

            return add;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssWorksheetName"></param>
        public void MssWorksheet_GetName(object ssWorksheet, out string ssWorksheetName)
        {
            ssWorksheetName = (ssWorksheet as ExcelWorksheet).Name;
        } // MssWorksheet_GetName

        /// <summary>
        /// Apply the specified format to a range of cells
        /// </summary>
        /// <param name="range">The range of cells to apply the formatting to</param>
        /// <param name="format">The format to apply to the range of cells</param>
        private void ApplyFormatToRange(ExcelRange range, RCCellFormatRecord format)
        {
            if (format == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssFontName))
            {
                range.Style.Font.Name = format.ssSTCellFormat.ssFontName;
            }

            if (format.ssSTCellFormat.ssFontSize != 0)
            {
                range.Style.Font.Size = format.ssSTCellFormat.ssFontSize;
            }

            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBackgroundColor))
            {
                Color color = Util.ConvertFromColorCode(format.ssSTCellFormat.ssBackgroundColor);
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(color);
            }

            if (format.ssSTCellFormat.ssBold)
            {
                range.Style.Font.Bold = true;
            }

            if (format.ssSTCellFormat.ssBorderStyle > 0)
            {
                Color borderColor = new Color();
                if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBorderColor))
                {
                    borderColor = Util.ConvertFromColorCode(format.ssSTCellFormat.ssBorderColor);
                }
                ExcelBorderStyle borderStyle = (ExcelBorderStyle)format.ssSTCellFormat.ssBorderStyle;
                range.Style.Border.BorderAround(borderStyle, borderColor);
            }

            if (format.ssSTCellFormat.ssAutofitColumn)
            {
                range.AutoFitColumns();
            }

            range.Style.Numberformat.Format = format.ssSTCellFormat.ssNumberFormat;
        }

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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
            DataTable dt;
            RecordList rl = (RecordList)ssDataSet;
            rl.Reset();

            if (rl.Data.Count > 0)
            {
                dt = Util.ConvertArrayListToDataTable(rl.Data);

                //exclude platform generated fields 
                if (dt.Columns.Contains("OptimizedAttributes")) dt.Columns.Remove("OptimizedAttributes");

                //if (dt.Columns.Contains("ChangedAttributes")) dt.Columns.Remove("ChangedAttributes");
                if (dt.Columns.Contains("OriginalKey")) dt.Columns.Remove("OriginalKey");

                ws.Cells[ssRowStart, ssColumnStart].LoadFromDataTable(dt, ssExportHeaders);
            }

            ApplyFormatToRange(ws.Cells[ssRowStart, ssColumnStart], ssCellFormat);
        } // MssCell_WriteRangeWithFormat

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="sspassword"></param>
        public void MssWorksheet_Protect(object ssWorksheet, string sspassword)
        {
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
            ws.Protection.IsProtected = true;
            ws.Protection.AllowEditObject = false;
            ws.Protection.SetPassword(sspassword);

        } // MssWorksheet_Protect

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        public void MssWorkbook_Close(object ssWorkbook)
        {
            ExcelPackage p = ssWorkbook as ExcelPackage;
            p.Dispose();
            p = null;
        } // MssWorkbook_Close


        /// <summary>
        /// Rename a worksheet
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssName"></param>
        public void MssWorksheet_Rename(object ssWorksheet, string ssName)
        {
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;
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
            ssBinaryData = p.GetAsByteArray();
        } // MssWorkbook_GetBinaryData

        /// <summary>
        /// Opens an existing workbook for editing by either specifying a name or the binary data.
        /// </summary>
        /// <param name="ssFileName">Location of the file that you want to open. Set to empty string "" when using binary data</param>
        /// <param name="ssBinary_Data">Binary data of the file that you want to open. Set to nullbinary() if using FileName</param>
        /// <param name="ssWorkbook">The workbook that you want to work with.</param>
        public void MssWorkbook_Open(string ssFileName, byte[] ssBinary_Data, out object ssWorkbook)
        {
            ExcelPackage p = new ExcelPackage();
            if (ssFileName.ToLower().StartsWith("http:") || ssFileName.ToLower().StartsWith("https:"))
            {
                System.Net.HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ssFileName);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                p.Load(response.GetResponseStream());
            }
            else if (!string.IsNullOrEmpty(ssFileName))
            {
                p.Load(System.IO.File.Open(ssFileName, System.IO.FileMode.OpenOrCreate));
            }
            else if (ssBinary_Data.LongLength > 0)
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
        ///  Creates a new excel workbook, optionally specifying the name of the fiirst sheet.
        /// </summary>
        /// <param name="ssWorkbook">The newly created workbook</param>
        /// <param name="ssFirstSheetName">Specify the name of the initial sheet in the workbook. Default = "Sheet1"</param>
        public void MssWorkbook_Create(out object ssWorkbook, string ssFirstSheetName)
        {
            ExcelPackage p = new ExcelPackage();
            p.Workbook.Worksheets.Add(string.IsNullOrEmpty(ssFirstSheetName) ? "Sheet1" : ssFirstSheetName);
            ssWorkbook = p;
        } // MssWorkbook_Create

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
            ExcelWorksheet ws = ssWorksheet as ExcelWorksheet;

            var chart = ws.Drawings.AddChart(ssChartName, stringToChartType(ssChartType));
            chart.SetPosition(ssRowPos, 0, ssColPos, 0);
            chart.SetSize(ssWidth, ssHeight);

            STDataSeriesStructure curr_series = ssDataSeries_List.CurrentRec;
            for (int i = 0; i < ssDataSeries_List.Count; i++)
            {
                STRangeStructure valuerange = ssDataSeries_List.CurrentRec.ssSTDataSeries.ssValueRange.ssSTRange;
                STRangeStructure labelrange = ssDataSeries_List.CurrentRec.ssSTDataSeries.ssLabelRange.ssSTRange;

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
                series.Header = ssDataSeries_List.CurrentRec.ssSTDataSeries.ssName;
                ssDataSeries_List.Advance();
            }
        } // MssChart_Create

        private OfficeOpenXml.Drawing.Chart.eChartType stringToChartType(string chartType)
        {
            switch (chartType)
            {
                case "Area3D": return OfficeOpenXml.Drawing.Chart.eChartType.Area3D;
                case "AreaStacked3D": return OfficeOpenXml.Drawing.Chart.eChartType.AreaStacked3D;
                case "AreaStacked1003D": return OfficeOpenXml.Drawing.Chart.eChartType.AreaStacked1003D;
                case "BarClustered3D": return OfficeOpenXml.Drawing.Chart.eChartType.BarClustered3D;
                case "BarStacked3D": return OfficeOpenXml.Drawing.Chart.eChartType.BarStacked3D;
                case "BarStacked1003D": return OfficeOpenXml.Drawing.Chart.eChartType.BarStacked1003D;
                case "Column3D": return OfficeOpenXml.Drawing.Chart.eChartType.Column3D;
                case "ColumnClustered3D": return OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered3D;
                case "ColumnStacked3D": return OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked3D;
                case "ColumnStacked1003D": return OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked1003D;
                case "Line3D": return OfficeOpenXml.Drawing.Chart.eChartType.Line3D;
                case "Pie3D": return OfficeOpenXml.Drawing.Chart.eChartType.Pie3D;
                case "PieExploded3D": return OfficeOpenXml.Drawing.Chart.eChartType.PieExploded3D;
                case "Area": return OfficeOpenXml.Drawing.Chart.eChartType.Area;
                case "AreaStacked": return OfficeOpenXml.Drawing.Chart.eChartType.AreaStacked;
                case "AreaStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.AreaStacked100;
                case "BarClustered": return OfficeOpenXml.Drawing.Chart.eChartType.BarClustered;
                case "BarOfPie": return OfficeOpenXml.Drawing.Chart.eChartType.BarOfPie;
                case "BarStacked": return OfficeOpenXml.Drawing.Chart.eChartType.BarStacked;
                case "BarStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.BarStacked100;
                case "Bubble": return OfficeOpenXml.Drawing.Chart.eChartType.Bubble;
                case "Bubble3DEffect": return OfficeOpenXml.Drawing.Chart.eChartType.Bubble3DEffect;
                case "ColumnClustered": return OfficeOpenXml.Drawing.Chart.eChartType.ColumnClustered;
                case "ColumnStacked": return OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked;
                case "ColumnStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.ColumnStacked100;
                case "ConeBarClustered": return OfficeOpenXml.Drawing.Chart.eChartType.ConeBarClustered;
                case "ConeBarStacked": return OfficeOpenXml.Drawing.Chart.eChartType.ConeBarStacked;
                case "ConeBarStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.ConeBarStacked100;
                case "ConeCol": return OfficeOpenXml.Drawing.Chart.eChartType.ConeCol;
                case "ConeColClustered": return OfficeOpenXml.Drawing.Chart.eChartType.ConeColClustered;
                case "ConeColStacked": return OfficeOpenXml.Drawing.Chart.eChartType.ConeColStacked;
                case "ConeColStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.ConeColStacked100;
                case "CylinderBarClustered": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderBarClustered;
                case "CylinderBarStacked": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderBarStacked;
                case "CylinderBarStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderBarStacked100;
                case "CylinderCol": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderCol;
                case "CylinderColClustered": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderColClustered;
                case "CylinderColStacked": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderColStacked;
                case "CylinderColStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.CylinderColStacked100;
                case "Doughnut": return OfficeOpenXml.Drawing.Chart.eChartType.Doughnut;
                case "DoughnutExploded": return OfficeOpenXml.Drawing.Chart.eChartType.DoughnutExploded;
                case "Line": return OfficeOpenXml.Drawing.Chart.eChartType.Line;
                case "LineMarkers": return OfficeOpenXml.Drawing.Chart.eChartType.LineMarkers;
                case "LineMarkersStacked": return OfficeOpenXml.Drawing.Chart.eChartType.LineMarkersStacked;
                case "LineMarkersStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.LineMarkersStacked100;
                case "LineStacked": return OfficeOpenXml.Drawing.Chart.eChartType.LineStacked;
                case "LineStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.LineStacked100;
                case "Pie": return OfficeOpenXml.Drawing.Chart.eChartType.Pie;
                case "PieExploded": return OfficeOpenXml.Drawing.Chart.eChartType.PieExploded;
                case "PieOfPie": return OfficeOpenXml.Drawing.Chart.eChartType.PieOfPie;
                case "PyramidBarClustered": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidBarClustered;
                case "PyramidBarStacked": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidBarStacked;
                case "PyramidBarStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidBarStacked100;
                case "PyramidCol": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidCol;
                case "PyramidColClustered": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidColClustered;
                case "PyramidColStacked": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidColStacked;
                case "PyramidColStacked100": return OfficeOpenXml.Drawing.Chart.eChartType.PyramidColStacked100;
                case "Radar": return OfficeOpenXml.Drawing.Chart.eChartType.Radar;
                case "RadarFilled": return OfficeOpenXml.Drawing.Chart.eChartType.RadarFilled;
                case "RadarMarkers": return OfficeOpenXml.Drawing.Chart.eChartType.RadarMarkers;
                case "StockHLC": return OfficeOpenXml.Drawing.Chart.eChartType.StockHLC;
                case "StockOHLC": return OfficeOpenXml.Drawing.Chart.eChartType.StockOHLC;
                case "StockVHLC": return OfficeOpenXml.Drawing.Chart.eChartType.StockVHLC;
                case "StockVOHLC": return OfficeOpenXml.Drawing.Chart.eChartType.StockVOHLC;
                case "Surface": return OfficeOpenXml.Drawing.Chart.eChartType.Surface;
                case "SurfaceTopView": return OfficeOpenXml.Drawing.Chart.eChartType.SurfaceTopView;
                case "SurfaceTopViewWireframe": return OfficeOpenXml.Drawing.Chart.eChartType.SurfaceTopViewWireframe;
                case "SurfaceWireframe": return OfficeOpenXml.Drawing.Chart.eChartType.SurfaceWireframe;
                case "XYScatter": return OfficeOpenXml.Drawing.Chart.eChartType.XYScatter;
                case "XYScatterLines": return OfficeOpenXml.Drawing.Chart.eChartType.XYScatterLines;
                case "XYScatterLinesNoMarkers": return OfficeOpenXml.Drawing.Chart.eChartType.XYScatterLinesNoMarkers;
                case "XYScatterSmooth": return OfficeOpenXml.Drawing.Chart.eChartType.XYScatterSmooth;
                case "XYScatterSmoothNoMarkers": return OfficeOpenXml.Drawing.Chart.eChartType.XYScatterSmoothNoMarkers;
            }
            return OfficeOpenXml.Drawing.Chart.eChartType.Column3D;
        }

    } // CssAdvanced_Excel

} // OutSystems.NssAdvanced_Excel

