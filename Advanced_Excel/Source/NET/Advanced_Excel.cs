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
		/// Get the visibility of a worksheet. Returns the same values that Worksheet_Hide_Show accepts, so the result can be passed back to it.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssHidden">Visible = 0 - The worksheet is visible
		/// Hidden = 1 - The worksheet is hidden but can be shown by the user via the user interface
		/// VeryHidden = 2 - The worksheet is hidden and cannot be shown by the user via the user interface</param>
		public void MssWorksheet_GetVisibility(object ssWorksheet, out int ssHidden) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			ssHidden = (int)ws.Hidden;
		} // MssWorksheet_GetVisibility

		/// <summary>
		/// Places an image, such as a company logo, in the worksheet&apos;s page header or footer. The image appears when the sheet is printed or in print preview. The image occupies the chosen section, so setting text in that same section with Worksheet_SetHeader or Worksheet_SetFooter afterwards replaces it — use different sections for text and image.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssImage">The image content (PNG or JPEG). PNG transparency is preserved</param>
		/// <param name="ssInFooter">Places the image in the page footer instead of the header</param>
		/// <param name="ssAlignment">Section of the header or footer to place the image in: Left, Center, or Right</param>
		public void MssWorksheet_AddHeaderFooterImage(object ssWorksheet, byte[] ssImage, bool ssInFooter, string ssAlignment) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);

			if (ssImage == null || ssImage.Length == 0)
			{
				throw new ArgumentException("Image is required.", nameof(ssImage));
			}

			PictureAlignment alignment;
			switch ((ssAlignment ?? "").Trim().ToLowerInvariant())
			{
				case "":
				case "left":
					alignment = PictureAlignment.Left;
					break;
				case "center":
				case "centered":
					alignment = PictureAlignment.Centered;
					break;
				case "right":
					alignment = PictureAlignment.Right;
					break;
				default:
					throw new ArgumentException("Invalid Alignment '" + ssAlignment + "'. Use Left, Center, or Right.", nameof(ssAlignment));
			}

			ExcelHeaderFooter headerFooter = HeaderFooterOf(ws);
			ExcelHeaderFooterText target = ssInFooter ? headerFooter.OddFooter : headerFooter.OddHeader;

			// InsertPicture copies the picture into the package, so the image can be disposed here.
			using (Image img = Util.ImageFromBytes(ssImage, nameof(ssImage)))
			{
				target.InsertPicture(img, alignment);
			}
		} // MssWorksheet_AddHeaderFooterImage

		/// <summary>
		/// Sets how the worksheet is displayed on screen: gridlines, row and column headers, and zoom.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with</param>
		/// <param name="ssShowGridLines">Show the cell gridlines. Turn off for report or dashboard sheets</param>
		/// <param name="ssShowHeaders">Show the row numbers and column letters</param>
		/// <param name="ssZoomScale">Zoom percentage, from 10 to 400</param>
		public void MssWorksheet_SetView(object ssWorksheet, bool ssShowGridLines, bool ssShowHeaders, int ssZoomScale) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);

			if (ssZoomScale < 10 || ssZoomScale > 400)
			{
				throw new ArgumentException("ZoomScale must be between 10 and 400.", nameof(ssZoomScale));
			}

			ws.View.ShowGridLines = ssShowGridLines;
			ws.View.ShowHeaders = ssShowHeaders;
			ws.View.ZoomScale = ssZoomScale;
		} // MssWorksheet_SetView

		/// <summary>
		/// Sorts the rows of a range by one or more columns. Cell values and formatting move together with their row.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRange">A1 range, e.g. A1:D100.</param>
		/// <param name="ssHasHeader">Keeps the first row in place instead of sorting it into the data.</param>
		/// <param name="ssSortFields">Applied in order (first field is the primary sort).</param>
		public void MssWorksheet_SortRange(object ssWorksheet, string ssRange, bool ssHasHeader, RLSortFieldRecordList ssSortFields) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);

			if (string.IsNullOrEmpty(ssRange))
			{
				throw new ArgumentException("Range is required.", nameof(ssRange));
			}
			if (ssSortFields == null || ssSortFields.Count == 0)
			{
				throw new ArgumentException("At least one sort field is required.", nameof(ssSortFields));
			}

			ExcelAddress address = new ExcelAddress(ssRange);
			int startRow = ssHasHeader ? address.Start.Row + 1 : address.Start.Row;
			if (startRow > address.End.Row)
			{
				return;
			}

			int width = address.End.Column - address.Start.Column + 1;
			var columns = new int[ssSortFields.Count];
			var descending = new bool[ssSortFields.Count];

			for (int i = 0; i < ssSortFields.Count; i++)
			{
				var field = ssSortFields[i].ssSTSortField;
				if (field.ssColumn < 1 || field.ssColumn > width)
				{
					throw new ArgumentException("Sort column " + field.ssColumn + " is outside the range " + ssRange + ", which is " + width + " column(s) wide. Columns are counted from the left of the range starting at 1.", nameof(ssSortFields));
				}
				// the extension counts range columns from 1; EPPlus expects a 0-based offset
				columns[i] = field.ssColumn - 1;
				descending[i] = field.ssDescending;
			}

			// sorting relocates cells without adjusting merges or formula references, so both
			// would be silently corrupted; refuse them like Excel does. A merge confined to the
			// excluded header row is fine.
			foreach (var mergedAddress in ws.MergedCells)
			{
				if (string.IsNullOrEmpty(mergedAddress))
				{
					continue;
				}
				ExcelAddress merged = new ExcelAddress(mergedAddress);
				bool intersects = merged.End.Row >= startRow && merged.Start.Row <= address.End.Row
				               && merged.End.Column >= address.Start.Column && merged.Start.Column <= address.End.Column;
				if (intersects)
				{
					throw new ArgumentException("The range contains merged cells (" + mergedAddress + "). Sorting would separate the merged cells from their values; unmerge them before sorting.", nameof(ssRange));
				}
			}

			foreach (var cell in ws.Cells[startRow, address.Start.Column, address.End.Row, address.End.Column])
			{
				if (!string.IsNullOrEmpty(cell.Formula))
				{
					throw new ArgumentException("The range contains a formula in " + cell.Address + ". Sorting moves formulas without adjusting their references, which produces incorrect results; sort values only.", nameof(ssRange));
				}
			}

			ws.Cells[startRow, address.Start.Column, address.End.Row, address.End.Column]
			  .Sort(columns, descending, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.CompareOptions.None);
		} // MssWorksheet_SortRange

		/// <summary>
		/// Saves the workbook as an AES-encrypted, password-protected .xlsx. The password is required to open the file in Excel.
		/// </summary>
		/// <param name="ssWorkbook">The workbook to save.</param>
		/// <param name="ssPassword">Password required to open the resulting file.</param>
		/// <param name="ssBinaryData">The encrypted .xlsx file</param>
		public void MssWorkbook_SaveWithPassword(object ssWorkbook, string ssPassword, out byte[] ssBinaryData) {
			ExcelPackage p = AsWorkbook(ssWorkbook);
			if (string.IsNullOrEmpty(ssPassword))
			{
				throw new ArgumentException("Password is required to save an encrypted workbook.", nameof(ssPassword));
			}

			Util.PreserveVisibleRowsForZeroHeightSheets(p);

			p.Encryption.IsEncrypted = true;
			p.Encryption.Password = ssPassword;
			ssBinaryData = SaveAndReload(p, ssPassword);
			// clear the flag so a later plain save isn't encrypted too
			p.Encryption.IsEncrypted = false;
		} // MssWorkbook_SaveWithPassword

		/// <summary>
		/// Creates a pivot table from a source range, placing row/column/value/filter fields as specified. The pivot recalculates when the file is opened in Excel.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet where the pivot table is placed</param>
		/// <param name="ssLocation">Top-left anchor cell in Worksheet, e.g. A1</param>
		/// <param name="ssSourceWorksheet">The worksheet holding the source data (same workbook)</param>
		/// <param name="ssSourceRange">Data range including the header row, e.g. A1:D100</param>
		/// <param name="ssPivotTableName">Unique name for the pivot table</param>
		/// <param name="ssFields">Fields to place on the pivot</param>
		/// <param name="ssRowGrandTotals">Show grand totals for rows</param>
		/// <param name="ssColumnGrandTotals">Show grand totals for columns</param>
		public void MssWorksheet_AddPivotTable(object ssWorksheet, string ssLocation, object ssSourceWorksheet, string ssSourceRange, string ssPivotTableName, RLPivotFieldRecordList ssFields, bool ssRowGrandTotals, bool ssColumnGrandTotals) {
			ExcelWorksheet target = AsWorksheet(ssWorksheet);
			ExcelWorksheet source = AsWorksheet(ssSourceWorksheet);

			if (!ReferenceEquals(target.Workbook, source.Workbook))
			{
				throw new ArgumentException("SourceWorksheet must belong to the same workbook as Worksheet.", nameof(ssSourceWorksheet));
			}

			if (string.IsNullOrEmpty(ssLocation))
			{
				throw new ArgumentException("Location (the anchor cell for the pivot table) is required.", nameof(ssLocation));
			}
			if (string.IsNullOrEmpty(ssSourceRange))
			{
				throw new ArgumentException("SourceRange is required.", nameof(ssSourceRange));
			}

			var pivot = target.PivotTables.Add(target.Cells[ssLocation], source.Cells[ssSourceRange], ssPivotTableName);
			pivot.RowGrandTotals = ssRowGrandTotals;
			pivot.ColumnGrandTotals = ssColumnGrandTotals;

			if (ssFields != null)
			{
				for (int i = 0; i < ssFields.Count; i++)
				{
					var f = ssFields[i].ssSTPivotField;
					if (string.IsNullOrEmpty(f.ssFieldName))
					{
						continue;
					}

					OfficeOpenXml.Table.PivotTable.ExcelPivotTableField field = null;
					try
					{
						field = pivot.Fields[f.ssFieldName];
					}
					catch
					{
						// fall through to the shared not-found error below
					}
					if (field == null)
					{
						throw new ArgumentException("Pivot field '" + f.ssFieldName + "' was not found in the source data's header row (" + ssSourceRange + ").", nameof(ssFields));
					}

					switch ((f.ssAxis ?? "").Trim().ToLowerInvariant())
					{
						case "row":
							pivot.RowFields.Add(field);
							break;
						case "column":
							pivot.ColumnFields.Add(field);
							break;
						case "filter":
							pivot.PageFields.Add(field);
							break;
						case "data":
							var dataField = pivot.DataFields.Add(field);
							dataField.Function = ParseEnum(
								string.IsNullOrEmpty(f.ssFunction) ? "Sum" : f.ssFunction,
								OfficeOpenXml.Table.PivotTable.DataFieldFunctions.Sum);
							break;
						default:
							throw new ArgumentException("Pivot field '" + f.ssFieldName + "' has an invalid Axis '" + f.ssAxis + "'. Use Row, Column, Data, or Filter.", nameof(ssFields));
					}
				}
			}
		} // MssWorksheet_AddPivotTable

		/// <summary>
		/// Returns the bounds of the populated area of a worksheet, so you can read a file whose size you don&apos;t know in advance.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssHasData">False if the worksheet is empty</param>
		/// <param name="ssStartRow">First used row (1-based); 0 if empty</param>
		/// <param name="ssStartColumn">First used column; 0 if empty</param>
		/// <param name="ssEndRow">Last used row; 0 if empty</param>
		/// <param name="ssEndColumn">Last used column; 0 if empty</param>
		/// <param name="ssAddress">A1-style range, e.g. A1:F20 (empty if no data)</param>
		public void MssWorksheet_GetUsedRange(object ssWorksheet, out bool ssHasData, out int ssStartRow, out int ssStartColumn, out int ssEndRow, out int ssEndColumn, out string ssAddress) {
			ssHasData = false;
			ssStartRow = 0;
			ssStartColumn = 0;
			ssEndRow = 0;
			ssEndColumn = 0;
			ssAddress = "";

			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			if (ws.Dimension == null)
			{
				return;
			}

			ssHasData = true;
			ssStartRow = ws.Dimension.Start.Row;
			ssStartColumn = ws.Dimension.Start.Column;
			ssEndRow = ws.Dimension.End.Row;
			ssEndColumn = ws.Dimension.End.Column;
			ssAddress = ws.Dimension.Address;
		} // MssWorksheet_GetUsedRange

		/// <summary>
		/// Writes many cells at once from a list of row/column/value records — the counterpart to Worksheet_ReadRange.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssCells">Cells to write (Row, Column, Value)</param>
		/// <param name="ssParseValues">When True, values that look like numbers/dates are written as typed values; when False, everything is written as text</param>
		/// <param name="ssCellsWritten">Number of cells written</param>
		public void MssWorksheet_WriteCells(object ssWorksheet, RLCellDataRecordList ssCells, bool ssParseValues, out int ssCellsWritten) {
			ssCellsWritten = 0;

			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			if (ssCells == null)
			{
				return;
			}

			for (int i = 0; i < ssCells.Count; i++)
			{
				var cd = ssCells[i].ssSTCellData;
				if (cd.ssRow < 1 || cd.ssColumn < 1)
				{
					continue; // skip records with invalid (1-based) coordinates
				}

				ws.Cells[cd.ssRow, cd.ssColumn].Value = ssParseValues ? ParseCellValue(cd.ssValue) : (object)(cd.ssValue ?? "");
				ssCellsWritten++;
			}
		} // MssWorksheet_WriteCells

		/// <summary>
		/// Auto-detect a written value's type from its text: whole number, decimal, boolean, date,
		/// otherwise the text itself. Parsing is invariant-culture so it matches how the other write
		/// actions interpret typed values.
		/// </summary>
		private static object ParseCellValue(string value)
		{
			if (string.IsNullOrEmpty(value))
			{
				return value ?? "";
			}

			var ci = System.Globalization.CultureInfo.InvariantCulture;

			long l;
			if (long.TryParse(value, System.Globalization.NumberStyles.Integer, ci, out l))
			{
				return l;
			}

			double d;
			if (double.TryParse(value, System.Globalization.NumberStyles.Float, ci, out d))
			{
				return d;
			}

			bool b;
			if (bool.TryParse(value, out b))
			{
				return b;
			}

			DateTime dt;
			if (DateTime.TryParse(value, ci, System.Globalization.DateTimeStyles.None, out dt))
			{
				return dt;
			}

			return value;
		}

		/// <summary>
		/// Finds and replaces text across a range (or the whole used range), returning how many cells changed.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssRange">Range to search; leave all fields 0 to search the whole used range</param>
		/// <param name="ssFindText">Text to find</param>
		/// <param name="ssReplaceText">Replacement text (empty removes the found text)</param>
		/// <param name="ssMatchCase">Case-sensitive match</param>
		/// <param name="ssMatchEntireCell">When True, only replace when the whole cell equals FindText</param>
		/// <param name="ssReplacedCount">Number of cells changed</param>
		public void MssCells_Replace(object ssWorksheet, RCRangeRecord ssRange, string ssFindText, string ssReplaceText, bool ssMatchCase, bool ssMatchEntireCell, out int ssReplacedCount) {
			ssReplacedCount = 0;

			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			if (string.IsNullOrEmpty(ssFindText))
			{
				throw new ArgumentException("FindText is required.", nameof(ssFindText));
			}

			ExcelRange range;
			var r = ssRange.ssSTRange;
			if (r.ssStartRow == 0 && r.ssStartCol == 0 && r.ssEndRow == 0 && r.ssEndCol == 0)
			{
				if (ws.Dimension == null)
				{
					return; // empty sheet, nothing to search
				}
				range = ws.Cells[ws.Dimension.Address];
			}
			else
			{
				range = ws.Cells[r.ssStartRow, r.ssStartCol, r.ssEndRow, r.ssEndCol];
			}

			string replaceWith = ssReplaceText ?? "";
			StringComparison comparison = ssMatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

			foreach (var cell in range)
			{
				// only replace within text cells, so numeric/date values are never corrupted
				string text = cell.Value as string;
				if (text == null)
				{
					continue;
				}

				string newText;
				if (ssMatchEntireCell)
				{
					newText = string.Equals(text, ssFindText, comparison) ? replaceWith : text;
				}
				else
				{
					newText = ReplaceAll(text, ssFindText, replaceWith, comparison);
				}

				if (newText != text)
				{
					cell.Value = newText;
					ssReplacedCount++;
				}
			}
		} // MssCells_Replace

		/// <summary>
		/// Replace every occurrence of <paramref name="find"/> in <paramref name="source"/> using the
		/// given comparison (.NET Framework's string.Replace has no culture/case-insensitive overload).
		/// </summary>
		private static string ReplaceAll(string source, string find, string replace, StringComparison comparison)
		{
			if (string.IsNullOrEmpty(find))
			{
				return source;
			}

			var sb = new System.Text.StringBuilder();
			int pos = 0;
			while (true)
			{
				int idx = source.IndexOf(find, pos, comparison);
				if (idx < 0)
				{
					sb.Append(source, pos, source.Length - pos);
					break;
				}
				sb.Append(source, pos, idx - pos);
				sb.Append(replace);
				pos = idx + find.Length;
			}
			return sb.ToString();
		}

		/// <summary>
		/// Add a data-validation rule to a cell range using a DataValidation configuration (whole-number, decimal, date, text-length, or custom formula), with optional input prompt and error message.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssCellRange">Range to validate, e.g. B2:B100.</param>
		/// <param name="ssValidation">The validation rule configuration.</param>
		public void MssWorksheet_AddDataValidation(object ssWorksheet, string ssCellRange, RCDataValidationRecord ssValidation) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			RequireAddress(ssCellRange, nameof(ssCellRange));
			STDataValidationStructure v = ssValidation.ssSTDataValidation;
			string type = (v.ssValidationType ?? "").Trim().ToLowerInvariant();
			OfficeOpenXml.DataValidation.ExcelDataValidationOperator op =
				ParseEnum(v.ssOperator, OfficeOpenXml.DataValidation.ExcelDataValidationOperator.between);

			switch (type)
			{
				case "integer":
					var iv = ws.DataValidations.AddIntegerValidation(ssCellRange);
					iv.Operator = op;
					if (!string.IsNullOrEmpty(v.ssFormula1)) iv.Formula.Value = int.Parse(v.ssFormula1, System.Globalization.CultureInfo.InvariantCulture);
					if (!string.IsNullOrEmpty(v.ssFormula2)) iv.Formula2.Value = int.Parse(v.ssFormula2, System.Globalization.CultureInfo.InvariantCulture);
					ApplyValidationCommon(iv, v);
					break;
				case "decimal":
					var dv = ws.DataValidations.AddDecimalValidation(ssCellRange);
					dv.Operator = op;
					if (!string.IsNullOrEmpty(v.ssFormula1)) dv.Formula.Value = double.Parse(v.ssFormula1, System.Globalization.CultureInfo.InvariantCulture);
					if (!string.IsNullOrEmpty(v.ssFormula2)) dv.Formula2.Value = double.Parse(v.ssFormula2, System.Globalization.CultureInfo.InvariantCulture);
					ApplyValidationCommon(dv, v);
					break;
				case "datetime":
					var dtv = ws.DataValidations.AddDateTimeValidation(ssCellRange);
					dtv.Operator = op;
					if (!string.IsNullOrEmpty(v.ssFormula1)) dtv.Formula.Value = DateTime.Parse(v.ssFormula1, System.Globalization.CultureInfo.InvariantCulture);
					if (!string.IsNullOrEmpty(v.ssFormula2)) dtv.Formula2.Value = DateTime.Parse(v.ssFormula2, System.Globalization.CultureInfo.InvariantCulture);
					ApplyValidationCommon(dtv, v);
					break;
				case "textlength":
					var tv = ws.DataValidations.AddTextLengthValidation(ssCellRange);
					tv.Operator = op;
					if (!string.IsNullOrEmpty(v.ssFormula1)) tv.Formula.Value = int.Parse(v.ssFormula1, System.Globalization.CultureInfo.InvariantCulture);
					if (!string.IsNullOrEmpty(v.ssFormula2)) tv.Formula2.Value = int.Parse(v.ssFormula2, System.Globalization.CultureInfo.InvariantCulture);
					ApplyValidationCommon(tv, v);
					break;
				case "custom":
					var cv = ws.DataValidations.AddCustomValidation(ssCellRange);
					// Excel ignores the operator of a custom rule, but EPPlus defaults it to "between" and then
					// refuses to save a rule without Formula2.
					cv.Operator = OfficeOpenXml.DataValidation.ExcelDataValidationOperator.equal;
					if (!string.IsNullOrEmpty(v.ssFormula1)) cv.Formula.ExcelFormula = v.ssFormula1;
					ApplyValidationCommon(cv, v);
					break;
				default:
					throw new ArgumentException("Invalid or missing ValidationType. Use Integer, Decimal, DateTime, TextLength, or Custom.", "ssValidation");
			}
		} // MssWorksheet_AddDataValidation

		/// <summary>
		/// Read every populated cell in a range into a list of {Row, Column, Value}, so the data can be processed in your logic.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to read from.</param>
		/// <param name="ssCellRange">Range to read, e.g. A1:E100.</param>
		/// <param name="ssCells">One record per populated cell.</param>
		public void MssWorksheet_ReadRange(object ssWorksheet, string ssCellRange, out RLCellDataRecordList ssCells) {
			ssCells = new RLCellDataRecordList();
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			if (string.IsNullOrEmpty(ssCellRange))
			{
				return;
			}

			foreach (var cell in ws.Cells[ssCellRange])
			{
				if (cell.Value == null)
				{
					continue;
				}
				RCCellDataRecord rec = new RCCellDataRecord();
				rec.ssSTCellData.ssRow = cell.Start.Row;
				rec.ssSTCellData.ssColumn = cell.Start.Column;
				// invariant culture so dates/decimals read back the same on any server locale
				rec.ssSTCellData.ssValue = Convert.ToString(cell.Value, System.Globalization.CultureInfo.InvariantCulture);
				ssCells.Add(rec);
			}
		} // MssWorksheet_ReadRange

		/// <summary>
		/// Add in-cell sparkline mini-charts (line, column, or stacked) for a data range, placed in a location range.
		/// </summary>
		/// <param name="ssWorksheet">The worksheet to work with.</param>
		/// <param name="ssLocationRange">Where the sparklines go (one per cell), e.g. F2:F10.</param>
		/// <param name="ssDataRange">The source data, e.g. A2:E10 (one row of data per sparkline).</param>
		/// <param name="ssSparklineType">Line, Column, or Stacked.</param>
		/// <param name="ssColor">Series color as hex, e.g. #638EC6. Blank = EPPlus default.</param>
		public void MssWorksheet_AddSparkline(object ssWorksheet, string ssLocationRange, string ssDataRange, string ssSparklineType, string ssColor) {
			ExcelWorksheet ws = AsWorksheet(ssWorksheet);
			OfficeOpenXml.Sparkline.eSparklineType type =
				ParseEnum(ssSparklineType, OfficeOpenXml.Sparkline.eSparklineType.Line);
			var group = ws.SparklineGroups.Add(type, new ExcelAddress(ssLocationRange), new ExcelAddress(ssDataRange));
			if (!string.IsNullOrEmpty(ssColor))
			{
				group.ColorSeries.SetColor(Util.ConvertFromColorCode(ssColor));
			}
		} // MssWorksheet_AddSparkline

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
			{
				// Excel allows several print areas in one value ("A1:B5,D1:E5"). EPPlus accepts any text
				// here and saves it as a broken name ("'Data'!#REF!"), so every part is checked first.
				foreach (string part in ssRange.Split(','))
				{
					RequireAddress(part, nameof(ssRange));
					string trimmed = part.Trim();
					int bang = trimmed.LastIndexOf('!');
					if (bang >= 0)
					{
						string sheet = trimmed.Substring(0, bang).Trim();
						if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
						{
							sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
						}
						if (!string.Equals(sheet, ws.Name, StringComparison.OrdinalIgnoreCase))
						{
							throw new ArgumentException("The print area must be on this worksheet ('" + ws.Name + "').", nameof(ssRange));
						}
					}
				}
				ws.PrinterSettings.PrintArea = ws.Cells[ssRange];
			}
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
			// Each call sets both parts; an empty part means none. EPPlus merges into an existing
			// Print_Titles name (and drops the rows part when columns are set again), so start from none.
			if (ws.Names.ContainsKey("_xlnm.Print_Titles"))
			{
				ws.Names.Remove("_xlnm.Print_Titles");
			}
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
			if (!string.IsNullOrEmpty(ssRange) && ssRange.IndexOf(',') >= 0)
			{
				throw new ArgumentException("A table needs one rectangular range, e.g. A1:D10.", nameof(ssRange));
			}
			RequireAddress(ssRange, nameof(ssRange));
			var newAddr = new ExcelAddressBase(ssRange);
			if (ssShowTotal && newAddr.End.Row >= ExcelPackage.MaxRows)
			{
				throw new ArgumentException("The totals row would be below the last row of the worksheet; end the range at least one row higher.", nameof(ssRange));
			}
			int startRow = newAddr.Start.Row, startCol = newAddr.Start.Column;
			int endRow = newAddr.End.Row + (ssShowTotal ? 1 : 0), endCol = newAddr.End.Column;
			foreach (var existing in ws.Tables)
			{
				var ex = existing.Address;
				if (startRow <= ex.End.Row && endRow >= ex.Start.Row && startCol <= ex.End.Column && endCol >= ex.Start.Column)
				{
					throw new ArgumentException("The range " + ssRange + " overlaps the table '" + existing.Name + "' (" + ex.Address + "). Tables can't overlap.", nameof(ssRange));
				}
			}
			var table = ws.Tables.Add(ws.Cells[ssRange], ssTableName);
			table.TableStyle = ParseEnum(ssTableStyle, OfficeOpenXml.Table.TableStyles.Medium2);
			table.ShowHeader = ssShowHeader;
			table.ShowTotal = ssShowTotal;
			table.ShowFilter = ssShowHeader && ssShowFilter;
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
			RequireAddress(ssCellName, nameof(ssCellName));
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
			RequireAddress(ssRangeStart, nameof(ssRangeStart));
			RequireAddress(ssRangeEnd, nameof(ssRangeEnd));
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
            ExcelPackage p = AsWorkbook(ssWorkbook);

            foreach (var worksheet in p.Workbook.Worksheets)
            {
                // EPPlus keeps a null placeholder in MergedCells for every unmerged range; skip them.
                string sheetRanges = "";
                foreach (var mergedRange in worksheet.MergedCells)
                {
                    if (!string.IsNullOrEmpty(mergedRange))
                    {
                        sheetRanges += $"{mergedRange}; ";
                    }
                }

                if (sheetRanges.Length > 0)
                {
                    anyMergedCells = true;
                    ssMergedRange += $"Worksheet: {worksheet.Name}\n";
                    ssMergedRange += sheetRanges.TrimEnd(new char[] { ';', ' ' });
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

            // EPPlus keeps a null placeholder in MergedCells for every unmerged range; skip them.
            foreach (var mergedRange in ws.MergedCells)
            {
                if (!string.IsNullOrEmpty(mergedRange))
                {
                    ssMergedRange += mergedRange + "; ";
                }
            }

            ssMergedRange = ssMergedRange.TrimEnd(new char[] { ';', ' ' });

            if (ssMergedRange.Length == 0)
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
            ExcelPackage p = AsWorkbook(ssWorkbook);

            foreach (var worksheet in p.Workbook.Worksheets)
            {
                worksheet.View.RightToLeft = true;
            }

            Util.PreserveVisibleRowsForZeroHeightSheets(p);
            ssBinaryData = SaveAndReload(p);
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
            if (wb == null)
            {
                throw new ArgumentException("Expected a Workbook object (e.g. from Workbook_Create or Workbook_Open).", nameof(ssWorkbook));
            }
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
            if (wb == null)
            {
                throw new ArgumentException("Expected a Workbook object (e.g. from Workbook_Create or Workbook_Open).", nameof(ssWorkbook));
            }
            var props = wb.Properties;
            var inProps = ssProperties.ssSTOfficeProperties;
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssAuthor)) { props.Author = inProps.ssAuthor ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssCategory)) { props.Category = inProps.ssCategory ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssComments)) { props.Comments = inProps.ssComments ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssCompany)) { props.Company = inProps.ssCompany ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssKeywords)) { props.Keywords = inProps.ssKeywords ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssLastModifiedBy)) { props.LastModifiedBy = inProps.ssLastModifiedBy ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssManager)) { props.Manager = inProps.ssManager ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssStatus)) { props.Status = inProps.ssStatus ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssSubject)) { props.Subject = inProps.ssSubject ?? string.Empty; }
            if (!ssIgnoreBlank || !string.IsNullOrEmpty(inProps.ssTitle)) { props.Title = inProps.ssTitle ?? string.Empty; }
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
            if (wb == null)
            {
                throw new ArgumentException("Expected a Workbook object (e.g. from Workbook_Create or Workbook_Open).", nameof(ssWorkbook));
            }
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
            int rowStart = address.Start.Row;
            int colStart = address.Start.Column;
            int rowEnd = address.End.Row;
            int colEnd = address.End.Column;
            // EPPlus parses "XFE1" or "A1048577" happily and only fails when the workbook is saved, so
            // reject a cell outside Excel's grid here, before the out parameters are set.
            if (rowStart > ExcelPackage.MaxRows || colStart > ExcelPackage.MaxColumns
                || rowEnd > ExcelPackage.MaxRows || colEnd > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).", nameof(ssAddress));
            }
            ssRowStart = rowStart;
            ssColStart = colStart;
            ssRowEnd = rowEnd;
            ssColEnd = colEnd;
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
            // Excel's grid ends at XFD1048576; a cell past it does not exist, so no address is returned.
            if (ssRowStart > ExcelPackage.MaxRows)
            {
                throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).", nameof(ssRowStart));
            }
            if (ssColStart > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).", nameof(ssColStart));
            }
            if (ssRowEnd > ExcelPackage.MaxRows)
            {
                throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).", nameof(ssRowEnd));
            }
            if (ssColEnd > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).", nameof(ssColEnd));
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
            RequireAddress(ssCellRange, nameof(ssCellRange));

            // Check the items before the rule is added: EPPlus only validates a list rule when the
            // workbook is saved, and a broken rule makes every later save fail.
            if (String.IsNullOrEmpty(ssItemsAddress))
            {
                if (ssItemsList == null || ssItemsList.Count == 0)
                {
                    throw new ArgumentException("The dropdown needs at least one item in ItemsList, or a range of cells in ItemsAddress.", nameof(ssItemsList));
                }
                // Excel limits an inline validation list to 255 characters in total (items plus
                // separators). Beyond that Excel silently drops or repairs the dropdown, so fail
                // clearly here instead.
                int totalLength = 0;
                for (int i = 0; i < ssItemsList.Count; i++)
                {
                    totalLength += (ssItemsList[i].ssSTItems.ssItemText ?? "").Length + 1;
                }
                if (totalLength > 256)
                {
                    throw new ArgumentException("The dropdown items exceed Excel's 255-character limit for inline validation lists. Write the items to a range of cells and pass that range in ItemsAddress instead.", nameof(ssItemsList));
                }
            }

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
            ExcelPackage ee = AsWorkbook(ssWorkbook);
            ExcelWorksheet target = null;

            if (!string.IsNullOrEmpty(ssWorksheetName))
            {
                target = ee.Workbook.Worksheets[ssWorksheetName];
            }
            if (ssWorksheetIndex > 0)
            {
                if (ssWorksheetIndex > ee.Workbook.Worksheets.Count)
                {
                    throw new ArgumentException("There is no worksheet " + ssWorksheetIndex + ": the workbook has " + ee.Workbook.Worksheets.Count + ".", "ssWorksheetIndex");
                }
                target = ee.Workbook.Worksheets[ssWorksheetIndex];
            }

            if (target == null)
            {
                if (!string.IsNullOrEmpty(ssWorksheetName))
                {
                    throw new ArgumentException("There is no worksheet named '" + ssWorksheetName + "'.", "ssWorksheetName");
                }
                return;
            }

            foreach (var sheet in ee.Workbook.Worksheets)
            {
                sheet.View.TabSelected = false;
            }
            target.View.TabSelected = true;
            ee.Workbook.View.ActiveTab = target.Index - 1;
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
            // No format input: keep the cell's own formatting (locked, font, fill). Applying an empty
            // CellFormat here unlocked the cell and removed a template's bold.
            MssCell_Write(ssWorksheet, null, ssRow, ssColumn, ssCellValue, ssCellType, true, new RCCellFormatRecord(null));
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
            // No format input: keep the cell's own formatting, as Cell_WriteByIndex does.
            MssCell_Write(ssWorksheet, ssCellName, 0, 0, ssCellValue, ssCellType, true, new RCCellFormatRecord(null));
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
            // No format input: keep the cells' own formatting.
            WriteColumnRange(ssWorksheet, ssRow, ssColumnStart, ssValueList, ssCellType, null);
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
            WriteColumnRange(ssWorksheet, ssRow, ssColumnStart, ssValueList, ssCellType, ssCellFormat);
        } // MssCell_WriteColumnRangeWithFormat

        /// <summary>
        /// Cell_WriteColumnRange(WithFormat): a null format keeps the cells' own formatting.
        /// </summary>
        private static void WriteColumnRange(object ssWorksheet, int ssRow, int ssColumnStart, RLValueRecordList ssValueList, string ssCellType, RCCellFormatRecord? format)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            DataTable dt = Util.RecordListToDataTable((RecordList)ssValueList);

            if (dt != null)
            {
                if (format.HasValue)
                {
                    Util.ValidateFormatNumbers(format.Value.ssSTCellFormat);
                }
                DataTable row = Util.Transpose(dt, ssCellType);
                ClearFormulas(ws, ssRow, ssColumnStart, row, false);
                ExcelRange range = (ExcelRange)ws.Cells[ssRow, ssColumnStart].LoadFromDataTable(row, false);
                Util.FormatDateCells(ws, range, row, false);

                if (format.HasValue)
                {
                    Util.ApplyFormatToRange(range, format.Value);
                }
            }
        } // WriteColumnRange

        /// <summary>
        /// Throws when a drawing of the given size, placed at the given cell (1-based) and offsets, would end beyond
        /// the last row or column, which makes the file invalid for Excel. EPPlus works out the lower-right anchor
        /// from the cell sizes, so a temporary shape is placed and measured, then removed. The real drawing is only
        /// added once it fits: EPPlus can't remove a picture cleanly (adding the same image again then fails).
        /// </summary>
        private static void RequireFitsInGrid(ExcelWorksheet ws, int row, int rowOffsetPixels, int column, int columnOffsetPixels,
            int widthPixels, int heightPixels, string what, string rowParam, string columnParam)
        {
            int toRow, toColumn;
            ExcelShape probe = ws.Drawings.AddShape("AdvancedExcelFitCheck_" + Guid.NewGuid().ToString("N"), eShapeStyle.Rect);
            try
            {
                probe.SetPosition(row - 1, rowOffsetPixels, column - 1, columnOffsetPixels);
                probe.SetSize(widthPixels, heightPixels);
                toRow = probe.To.Row;
                toColumn = probe.To.Column;
            }
            finally
            {
                ws.Drawings.Remove(probe);
            }
            if (toRow > ExcelPackage.MaxRows - 1)
            {
                throw new ArgumentException("The " + what + " does not fit on the sheet: starting at row " + row
                    + " with a height of " + heightPixels + " pixels it would end below row " + ExcelPackage.MaxRows + ".", rowParam);
            }
            if (toColumn > ExcelPackage.MaxColumns - 1)
            {
                throw new ArgumentException("The " + what + " does not fit on the sheet: starting at column " + column
                    + " with a width of " + widthPixels + " pixels it would end beyond column XFD (" + ExcelPackage.MaxColumns + ").", columnParam);
            }
        } // RequireFitsInGrid

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
            if (ssRow < 1 || ssRow > ExcelPackage.MaxRows)
            {
                throw new ArgumentException("Row must be between 1 and " + ExcelPackage.MaxRows + ".", nameof(ssRow));
            }
            if (ssColumn < 1 || ssColumn > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("Column must be between 1 and " + ExcelPackage.MaxColumns + ".", nameof(ssColumn));
            }

            // EPPlus keeps the image for Worksheet_GetImages, so it isn't disposed here.
            Image image = Util.ImageFromBytes(ssImage, nameof(ssImage));
            // EPPlus sizes a new picture from the image's pixels at its resolution, relative to 96 dpi.
            // A resolution of 0 (or a bogus one) would make that size infinite or negative, so the
            // resolution is used only when it is a usable dpi value.
            double horizontalResolution = image.HorizontalResolution;
            double verticalResolution = image.VerticalResolution;
            if (double.IsNaN(horizontalResolution) || double.IsInfinity(horizontalResolution) || horizontalResolution < 1)
            {
                horizontalResolution = 96;
            }
            if (double.IsNaN(verticalResolution) || double.IsInfinity(verticalResolution) || verticalResolution < 1)
            {
                verticalResolution = 96;
            }
            double widthPixels = Math.Round(image.Width / (horizontalResolution / 96.0));
            double heightPixels = Math.Round(image.Height / (verticalResolution / 96.0));
            if (widthPixels < 1 || widthPixels > 1000000 || heightPixels < 1 || heightPixels > 1000000)
            {
                image.Dispose();
                throw new ArgumentException("The image's size or resolution is out of range.", nameof(ssImage));
            }

            try
            {
                RequireFitsInGrid(ws, ssRow, 0, ssColumn, 0, (int)widthPixels, (int)heightPixels, "image", nameof(ssRow), nameof(ssColumn));
            }
            catch
            {
                // A rejected picture leaves no drawing behind, so nothing else owns the image.
                image.Dispose();
                throw;
            }

            ExcelPicture pic;
            try
            {
                pic = ws.Drawings.AddPicture(ssImageName, image);
            }
            catch
            {
                // AddPicture adds nothing to the sheet when it throws, so the image is still ours.
                image.Dispose();
                throw;
            }
            pic.SetPosition(ssRow - 1, 0, ssColumn - 1, 0);
            // Place the picture at exactly the size that was checked, whatever EPPlus derives from the dpi.
            pic.SetSize((int)widthPixels, (int)heightPixels);
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
            if (!(ssDataSet is RecordList))
            {
                throw new ArgumentException("DataSet must be a record list (a List of a Record or Structure).", nameof(ssDataSet));
            }
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            DataTable dt = Util.RecordListToDataTable((RecordList)ssDataSet);

            if (dt != null)
            {
                Util.ValidateFormatNumbers(ssCellFormat.ssSTCellFormat);
                ClearFormulas(ws, ssRowStart, ssColumnStart, dt, false);
                ExcelRange range = (ExcelRange)ws.Cells[ssRowStart, ssColumnStart].LoadFromDataTable(dt, false);
                Util.FormatDateCells(ws, range, dt, false);

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
            ExcelPackage p = AsWorkbook(ssWorkbook);
            Util.ValidateSheetName(ssWorksheetName, nameof(ssWorksheetName));
            ExcelWorksheet ws;
            ws = p.Workbook.Worksheets.Add(ssWorksheetName);
            ssWorksheet = ws;
        } // MssWorkbook_AddName

        /// <summary>
        /// Set the width of a column on a specific worksheet, in Excel character widths (the unit of Excel's Column Width).
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with</param>
        /// <param name="ssColumnNumber">The column number, starting at 1</param>
        /// <param name="ssDesiredWidth">The width in Excel character widths, as in Excel's Column Width (the default width is about 9.14). Not pixels.</param>
        public void MssColumn_SetWidth(object ssWorksheet, int ssColumnNumber, decimal ssDesiredWidth)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            ws.Column(ssColumnNumber).Width = (double)ssDesiredWidth;
        } // MssColumn_SetWidth

        /// <summary>
        /// Set the height of a row on a specific worksheet, in points (the unit of Excel's Row Height).
        /// </summary>
        /// <param name="ssWorksheet">The worksheet to work with
        /// </param>
        /// <param name="ssRowNumber">The number of the row to set the height for</param>
        /// <param name="ssDesiredHeight">The height in points, as in Excel's Row Height (the default height is 15). Not pixels.</param>
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
            if (ssRow < 1 || ssRow > ExcelPackage.MaxRows)
            {
                throw new ArgumentException("Row must be between 1 and " + ExcelPackage.MaxRows + ".", "ssRow");
            }
            if (ssColumn < 1 || ssColumn > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("Column must be between 1 and " + ExcelPackage.MaxColumns + ".", "ssColumn");
            }

            if (ssReadText)
                ssCellValue = CellText(ws.Cells[ssRow, ssColumn]);
            else
                ssCellValue = CellRawText(ws.GetValue(ssRow, ssColumn));
        } // MssCell_ReadByIndex

        /// <summary>
        /// Convert a cell value to the invariant text returned by Cell_Read when ReadText is false.
        /// </summary>
        private static string CellRawText(object value)
        {
            return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
        } // CellRawText

        /// <summary>
        /// The cell as Excel shows it. Excel shows a boolean as TRUE or FALSE whatever its number
        /// format; EPPlus's Text gives "1" or "0".
        /// </summary>
        private static string CellText(ExcelRange cell)
        {
            object value = cell.Value;
            if (value is bool)
            {
                return (bool)value ? "TRUE" : "FALSE";
            }
            return cell.Text;
        } // CellText

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
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            RequireAddress(ssCellName, "ssCellName");
            string address = (ssCellName ?? "").Trim();
            int bang = address.LastIndexOf('!');
            if (bang >= 0)
            {
                string sheet = address.Substring(0, bang).Trim('\'');
                if (!string.Equals(sheet, ws.Name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("'" + ssCellName + "' names another worksheet; read it from that worksheet.", "ssCellName");
                }
                address = address.Substring(bang + 1);
            }
            if (!ExcelCellBase.IsValidCellAddress(address))
            {
                throw new ArgumentException("'" + ssCellName + "' must identify a single cell address.", "ssCellName");
            }
            ExcelCellAddress addr = new ExcelCellAddress(address);
            if (addr.Row < 1 || addr.Row > ExcelPackage.MaxRows || addr.Column < 1 || addr.Column > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("Cell address must be inside the Excel grid.", "ssCellName");
            }
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
            // OOXML stores a formula without its '=', as Cell_Write's "formula" type does.
            ws.Cells[ssRow, ssColumn].Formula = ssFormula.TrimStart('=');
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
            ExcelPackage p = AsWorkbook(ssWorkbook);
            ExcelWorksheet wsToCopy = AsWorksheet(ssWorksheetToCopy);
            Util.ValidateSheetName(ssWorksheetName, nameof(ssWorksheetName));
            ExcelWorksheet ws;
            ws = p.Workbook.Worksheets.Add(ssWorksheetName, wsToCopy);
            Util.PreserveConditionalFormattingStyles(wsToCopy, ws);
            Util.PreserveHyperlinks(wsToCopy, ws);
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
            p = AsWorkbook(ssWorkbook);
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
            ExcelPackage p = AsWorkbook(ssWorkbook);
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
            ExcelPackage ee = AsWorkbook(ssWorkbook);
            ee.Workbook.Worksheets.Delete(ssNameToDelete);
        } // MssWorksheet_DeleteByName

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        /// <param name="ssIndexToDelete"></param>
        public void MssWorksheet_DeleteByIndex(object ssWorkbook, int ssIndexToDelete)
        {
            ExcelPackage p = AsWorkbook(ssWorkbook);
            if (ssIndexToDelete < 1 || ssIndexToDelete > p.Workbook.Worksheets.Count)
            {
                throw new ArgumentException($"Worksheet index {ssIndexToDelete} does not exist: the workbook has {p.Workbook.Worksheets.Count} worksheet(s), numbered from 1.", nameof(ssIndexToDelete));
            }
            p.Workbook.Worksheets.Delete(ssIndexToDelete);
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
        /// Column3D (not supported)
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
        /// StockHLC (not supported)
        /// StockOHLC (not supported)
        /// StockVHLC (not supported)
        /// StockVOHLC (not supported)
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
        /// <param name="ssRowPos">Row of the chart's upper-left corner, starting at 1</param>
        /// <param name="ssColPos">Column of the chart's upper-left corner, starting at 1</param>
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

            // The name is added to the workbook's names, so a repeat call is found there.
            if (!ws.Names.ContainsKey(ssName) && !ws.Workbook.Names.ContainsKey(ssName))
            {
                // No format input: keep the cells' own formatting.
                WriteRange(ssWorksheet, ssRowStart, ssColumnStart, ssDataSet, null, false);
                ws.Workbook.Names.Add(ssName, ws.Cells[ssRowStart, ssColumnStart, ssRowStart + ((RecordList)ssDataSet).Length - 1, ssColumnStart]);
            }
        } // MssWorksheet_AddName

        /// <summary>
        /// Opens an existing workbook for editing and keeps it in memory
        /// </summary>
        /// <param name="ssBinaryData"></param>
        /// <param name="ssWorkbook"></param>
        public void MssWorkbook_Open_BinaryData(byte[] ssBinaryData, string ssPassword, out object ssWorkbook)
        {
            ssWorkbook = OpenWorkbook("", ssBinaryData, ssPassword, nameof(ssBinaryData));
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
                Util.ValidateSheetName(ssFirstSheetName, nameof(ssFirstSheetName));
                // Strip trailing "1" so "Sheet1" produces Sheet1/Sheet2/Sheet3, not Sheet1/Sheet12/Sheet13.
                string baseName = (ssFirstSheetName.Length > 1 && ssFirstSheetName.EndsWith("1"))
                    ? ssFirstSheetName.Substring(0, ssFirstSheetName.Length - 1)
                    : ssFirstSheetName;
                if (ssNumberOfSheets > 1)
                {
                    Util.ValidateSheetName(baseName + ssNumberOfSheets, nameof(ssFirstSheetName));
                }

                wb.Worksheets.Add(ssFirstSheetName);
                if (ssNumberOfSheets > 1)
                {
                    for (int i = 2; i <= ssNumberOfSheets; i++)
                    {
                        wb.Worksheets.Add(baseName + i);
                    }
                }
            }
            else
            {
                foreach (RCNewSheetRecord item in ssSheetNames)
                {
                    Util.ValidateSheetName(item.ssSTNewSheet.ssName, nameof(ssSheetNames));
                }
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
            ExcelPackage p = AsWorkbook(ssWorkbook);
            ExcelWorkbook wb = p.Workbook;

            // EPPlus removes the whole workbookProtection element (password hash and the other locks
            // included) when a lock is set to false. Clear the locks first, set the ones asked for, and
            // then put back the password: the new one, or the one the workbook already had.
            string keptPasswordHash = Util.WorkbookPasswordHash(wb);
            wb.Protection.LockRevision = false;
            wb.Protection.LockStructure = false;
            wb.Protection.LockWindows = false;

            if (!ssLockStructure && !ssLockWindows && !ssLockRevision)
            {
                return;
            }

            wb.Protection.LockRevision = ssLockRevision;
            wb.Protection.LockStructure = ssLockStructure;
            wb.Protection.LockWindows = ssLockWindows;

            if (!string.IsNullOrEmpty(ssPassword))
            {
                wb.Protection.SetPassword(ssPassword);
            }
            else if (keptPasswordHash != null)
            {
                Util.SetWorkbookPasswordHash(wb, keptPasswordHash);
            }
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
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            if (string.IsNullOrEmpty(ssCellName) && (ssColumnNumber <= 0 || ssRowNumber <= 0))
            {
                throw new ArgumentException("You need to specify a valid cell name (i.e. A4) or cell index (row/column combination)", nameof(ssCellName));
            }

            ExcelRange range;
            if (!string.IsNullOrEmpty(ssCellName))
            {
                RequireAddress(ssCellName, nameof(ssCellName));
                ExcelAddress addr = new ExcelAddress(ssCellName);
                if (addr.Start.Row > ExcelPackage.MaxRows || addr.Start.Column > ExcelPackage.MaxColumns
                    || addr.End.Row > ExcelPackage.MaxRows || addr.End.Column > ExcelPackage.MaxColumns)
                {
                    throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).", nameof(ssCellName));
                }
                range = ws.Cells[ssCellName];
            }
            else
            {
                if (ssRowNumber > ExcelPackage.MaxRows)
                {
                    throw new ArgumentException("Row must be between 1 and " + ExcelPackage.MaxRows + ".", nameof(ssRowNumber));
                }
                if (ssColumnNumber > ExcelPackage.MaxColumns)
                {
                    throw new ArgumentException("Column must be between 1 and " + ExcelPackage.MaxColumns + ".", nameof(ssColumnNumber));
                }
                range = ws.Cells[ssRowNumber, ssColumnNumber];
            }

            if (ssImageFile == null || ssImageFile.Length == 0)
            {
                throw new ArgumentException("No image data supplied.", nameof(ssImageFile));
            }
            if (ssImageWidth < 1)
            {
                throw new ArgumentException("The image width must be at least 1 pixel.", nameof(ssImageWidth));
            }
            if (ssImageHeight < 1)
            {
                throw new ArgumentException("The image height must be at least 1 pixel.", nameof(ssImageHeight));
            }
            if (ssMarginTop < 0)
            {
                throw new ArgumentException("The top margin cannot be negative.", nameof(ssMarginTop));
            }
            if (ssMarginLeft < 0)
            {
                throw new ArgumentException("The left margin cannot be negative.", nameof(ssMarginLeft));
            }

            // EPPlus keeps the image for Worksheet_GetImages, as it does for pictures it loads from a
            // file, so neither the image nor the picture is disposed here.
            bool byName = !string.IsNullOrEmpty(ssCellName);
            RequireFitsInGrid(ws, range.Start.Row, ssMarginTop, range.Start.Column, ssMarginLeft, ssImageWidth, ssImageHeight, "image",
                byName ? nameof(ssCellName) : nameof(ssRowNumber), byName ? nameof(ssCellName) : nameof(ssColumnNumber));

            Image image = Util.ImageFromBytes(ssImageFile, nameof(ssImageFile));
            ExcelPicture picture;
            try
            {
                picture = ws.Drawings.AddPicture(ssImageName, image);
            }
            catch
            {
                // AddPicture adds nothing to the sheet when it throws, so the image is still ours.
                image.Dispose();
                throw;
            }
            picture.SetPosition(range.Start.Row - 1, ssMarginTop, range.Start.Column - 1, ssMarginLeft);
            picture.SetSize(ssImageWidth, ssImageHeight);
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

            // The OutSystems default Range (0,0,0,0) names no cell: deleting nothing is the whole job.
            if (ssRange.ssSTRange.ssStartRow == 0 && ssRange.ssSTRange.ssEndRow == 0
                && ssRange.ssSTRange.ssStartCol == 0 && ssRange.ssSTRange.ssEndCol == 0)
            {
                return;
            }

            RemoveCommentsInRange(ws, ssRange.ssSTRange.ssStartRow, ssRange.ssSTRange.ssEndRow,
                                      ssRange.ssSTRange.ssStartCol, ssRange.ssSTRange.ssEndCol);
        } // MssComment_Delete

        /// <summary>
        /// Remove every comment within the given 1-based cell rectangle. A zero/empty range
        /// (start &gt; end) simply iterates nothing.
        /// </summary>
        private static void RemoveCommentsInRange(ExcelWorksheet ws, int startRow, int endRow, int startCol, int endCol)
        {
            if (startRow > endRow || startCol > endCol || ws == null || ws.Comments == null || ws.Comments.Count == 0)
            {
                return;
            }

            for (int i = ws.Comments.Count - 1; i >= 0; i--)
            {
                ExcelComment comment = ws.Comments[i];
                if (comment == null || string.IsNullOrEmpty(comment.Address))
                {
                    continue;
                }

                ExcelAddress addr = new ExcelAddress(comment.Address);
                if (addr.Start.Row >= startRow && addr.Start.Row <= endRow &&
                    addr.Start.Column >= startCol && addr.Start.Column <= endCol)
                {
                    ws.Comments.RemoveAt(i);
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

            List<ExcelRangeBase> result = ws.Cells.Where(c => c.Value != null && CellRawText(c.Value) == ssValueToFind).ToList();

            foreach (ExcelRangeBase item in result)
            {
                RCRangeRecord rc = new RCRangeRecord();
                rc.ssSTRange.ssStartRow = item.Start.Row;
                rc.ssSTRange.ssStartCol = item.Start.Column;

                ssListOfCells.Add(rc);
            }
        } // MssCells_FindByValue

        /// <summary>
        /// True when the rule is stored with aboveAverage="0" (below the mean). Excel keeps rule priorities
        /// unique within a sheet, so the rule is found in the worksheet XML by its priority; EPPlus does not
        /// expose that attribute.
        /// </summary>
        private static bool IsBelowMeanRule(ExcelWorksheet ws, IExcelConditionalFormattingRule rule)
        {
            var xml = ws.WorksheetXml;
            var ns = new System.Xml.XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("d", xml.DocumentElement.NamespaceURI);
            var node = xml.SelectSingleNode("//d:cfRule[@stdDev][@priority='" + rule.Priority.ToString(System.Globalization.CultureInfo.InvariantCulture) + "']", ns) as System.Xml.XmlElement;
            string above = node == null ? "" : node.GetAttribute("aboveAverage");
            return above == "0" || string.Equals(above, "false", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// The colour of each stop of a colour-scale rule (low, [middle,] high) as #RRGGBB, read from the rule's
        /// XML. EPPlus leaves LowValue/MiddleValue/HighValue null on a rule loaded from a file, so its getters
        /// can't be used. A stop whose colour is not an rgb value (theme or indexed colour) reads as "".
        /// Always returns three entries; the unused ones are "". EPPlus has already checked, when it loaded
        /// the rule, that it has as many colours as stops (2 or 3).
        /// </summary>
        private static string[] ColorScaleColors(IExcelConditionalFormattingRule rule)
        {
            string[] colors = { "", "", "" };
            var node = ((ExcelConditionalFormattingRule)rule).Node;
            int i = 0;
            foreach (System.Xml.XmlElement c in node.SelectNodes("*[local-name()='colorScale']/*[local-name()='color']"))
            {
                string rgb = c.GetAttribute("rgb");
                colors[i++] = rgb.Length == 8 ? "#" + rgb.Substring(2).ToUpperInvariant()
                    : rgb.Length == 6 ? "#" + rgb.ToUpperInvariant() : "";
            }
            return colors;
        }

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
                RCConditionalFormatItemRecord newItem = new RCConditionalFormatItemRecord(null);
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
                        // EPPlus loads every stdDev rule as AboveStdDev and ignores aboveAverage="0", so a
                        // BelowStdDev rule from an opened file is recognised from its XML.
                        newItem.ssSTConditionalFormatItem.ssRuleType = IsBelowMeanRule(ws, item)
                            ? (int)eExcelConditionalFormattingRuleType.BelowStdDev
                            : (int)aboveStdDev.Type;
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
                        var rNotBetween = item as IExcelConditionalFormattingNotBetween;
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
                        string[] stops3 = ColorScaleColors(item);
                        newItem.ssSTConditionalFormatItem.ssLowColor = stops3[0];
                        newItem.ssSTConditionalFormatItem.ssMidColor = stops3[1];
                        newItem.ssSTConditionalFormatItem.ssHighColor = stops3[2];
                        break;
                    case eExcelConditionalFormattingRuleType.TwoColorScale:
                        newItem.ssSTConditionalFormatItem.ssRuleType = (int)item.Type;
                        string[] stops2 = ColorScaleColors(item);
                        newItem.ssSTConditionalFormatItem.ssLowColor = stops2[0];
                        newItem.ssSTConditionalFormatItem.ssHighColor = stops2[1];
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
                        var dataBarRule = item as IExcelConditionalFormattingDataBarGroup;
                        if (dataBarRule != null)
                        {
                            newItem.ssSTConditionalFormatItem.ssDataBarColor = ToHexColor(dataBarRule.Color);
                            newItem.ssSTConditionalFormatItem.ssShowValue = dataBarRule.ShowValue;
                        }
                        break;
                    default:
                        break;
                }

                // Read the rule's differential style back (skip the rule types that don't use one).
                switch (item.Type)
                {
                    case eExcelConditionalFormattingRuleType.TwoColorScale:
                    case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    case eExcelConditionalFormattingRuleType.FourIconSet:
                    case eExcelConditionalFormattingRuleType.FiveIconSet:
                    case eExcelConditionalFormattingRuleType.DataBar:
                        break;
                    default:
                        newItem.ssSTConditionalFormatItem.ssStyle = Util.ReadConditionalFormattingStyle(ws, item);
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
            if (ssHidden < (int)eWorkSheetHidden.Visible || ssHidden > (int)eWorkSheetHidden.VeryHidden)
            {
                throw new ArgumentException("Hidden must be 0 (visible), 1 (hidden) or 2 (very hidden).", nameof(ssHidden));
            }
            if (ssHidden != (int)eWorkSheetHidden.Visible && !Util.HasOtherVisibleSheet(ws))
            {
                throw new ArgumentException("At least one worksheet must stay visible: Excel can't open a workbook whose sheets are all hidden.", nameof(ssHidden));
            }
            ws.Hidden = (eWorkSheetHidden)ssHidden;
        } // MssWorksheet_Hide_Show

        /// <summary>
        /// Finds the first cell in Range whose value equals Value, compared as the text Cell_Read returns with ReadText False.
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssRange"></param>
        /// <param name="ssValue"></param>
        /// <param name="ssParameter1">Not used.</param>
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
                if (item.Value != null && CellRawText(item.Value) == ssValue)
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
            ExcelPackage p = AsWorkbook(ssWorkbook);
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

            // Showing keeps the column's width; only a column hidden by a zero width gets the default.
            if (!ssHidden && ws.Column(ssColumn).Width <= 0)
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

            if (string.IsNullOrEmpty(ssCellName))
            {
                if (ssCellRow < 1 || ssCellRow > ExcelPackage.MaxRows)
                {
                    throw new ArgumentException("Row must be between 1 and " + ExcelPackage.MaxRows + ".", "ssCellRow");
                }
                if (ssCellColumn < 1 || ssCellColumn > ExcelPackage.MaxColumns)
                {
                    throw new ArgumentException("Column must be between 1 and " + ExcelPackage.MaxColumns + ".", "ssCellColumn");
                }
                ExcelRange cell = ws.Cells[ssCellRow, ssCellColumn];
                ssCellValue = ssReadText ? CellText(cell) : CellRawText(cell.Value);
            }
            else
            {
                MssCell_ReadByName(ssWorksheet, ssCellName, ssReadText, out ssCellValue);
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
            if (address.End.Row > ExcelPackage.MaxRows || address.End.Column > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("The cell is beyond the last cell of a worksheet (XFD1048576).");
            }
            if (!ssPreserveFormat)
            {
                Util.ValidateFormatNumbers(ssCellFormat.ssSTCellFormat);
            }

            // A value replaces a formula. SetValue keeps the formula, and Excel would recalculate it on
            // open, silently losing the value.
            if (!string.Equals(ssCellType, "formula", StringComparison.OrdinalIgnoreCase))
            {
                ClearFormulas(ws, address.Start.Row, address.Start.Column, address.End.Row, address.End.Column);
            }

            switch (ssCellType.ToLower())
            {
                case "integer": ws.SetValue(address.Address, Convert.ToInt32(ssCellValue, System.Globalization.CultureInfo.InvariantCulture)); break;
                case "datetime":
                    ws.SetValue(address.Address, Convert.ToDateTime(ssCellValue, System.Globalization.CultureInfo.InvariantCulture));
                    Util.FormatDateCells(ws, address.Start.Row, address.Start.Column, address.Start.Row, address.Start.Column);
                    break;
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

            ExcelPackage ee = AsWorkbook(ssWorkbook);

            if (ssCurrentIndex > ee.Workbook.Worksheets.Count)
            {
                throw new ArgumentException("There is no worksheet " + ssCurrentIndex + ": the workbook has " + ee.Workbook.Worksheets.Count + ".", "ssCurrentIndex");
            }
            if (ssNewIndex == ssCurrentIndex)
            {
                return;
            }
            if (ssNewIndex >= ee.Workbook.Worksheets.Count)
            {
                ee.Workbook.Worksheets.MoveToEnd(ssCurrentIndex);
                return;
            }
            if (ssNewIndex < ssCurrentIndex)
            {
                ee.Workbook.Worksheets.MoveBefore(ssCurrentIndex, ssNewIndex);
                return;
            }
            ee.Workbook.Worksheets.MoveAfter(ssCurrentIndex, ssNewIndex);
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

            ExcelPackage p = AsWorkbook(ssWorkbook);
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

            ExcelPackage ee = AsWorkbook(ssWorkbook);
            if (ssIndexToDelete < 0 || ssIndexToDelete > ee.Workbook.Worksheets.Count)
            {
                throw new ArgumentException($"Worksheet index {ssIndexToDelete} does not exist: the workbook has {ee.Workbook.Worksheets.Count} worksheet(s), numbered from 1.", nameof(ssIndexToDelete));
            }
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
            ExcelPackage ee = AsWorkbook(ssWorkbook);
            ExcelWorksheet ws;
            ExcelWorksheet newSheet;

            if (ssWorksheet != null)
            {
                ws = AsWorksheet(ssWorksheet);
                string copyName = string.IsNullOrEmpty(ssWorksheetName) ? "Copy_" + ws.Name : ssWorksheetName;
                Util.ValidateSheetName(copyName, nameof(ssWorksheetName));
                newSheet = ee.Workbook.Worksheets.Add(copyName, ws);
                Util.PreserveConditionalFormattingStyles(ws, newSheet);
                Util.PreserveHyperlinks(ws, newSheet);
                if (ssIndexWhereToAdd > 0)
                {
                    MssWorkbook_ChangeSheetIndex((object)ee, newSheet.Index, ssIndexWhereToAdd);
                }
                return;
            }

            string effectiveName = string.IsNullOrEmpty(ssWorksheetName) ? "Sheet1" : ssWorksheetName;
            Util.ValidateSheetName(effectiveName, nameof(ssWorksheetName));
            newSheet = ee.Workbook.Worksheets.Add(effectiveName);
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

            ExcelPackage p = AsWorkbook(ssWorkbook);
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
            ssProperties = new RCWorksheetRecord(null);

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
            WriteRange(ssWorksheet, ssRowStart, ssColumnStart, ssDataSet, ssCellFormat, ssExportHeaders);
        } // MssCell_WriteRange

        /// <summary>
        /// Cell_WriteRange and Worksheet_AddName: a null format keeps the cells' own formatting.
        /// </summary>
        private static void WriteRange(object ssWorksheet, int ssRowStart, int ssColumnStart, object ssDataSet, RCCellFormatRecord? format, bool ssExportHeaders)
        {
            // A null or non-record-list DataSet cannot be written; say so instead of failing with a
            // NullReferenceException/InvalidCastException deep inside the write.
            if (!(ssDataSet is RecordList))
            {
                throw new ArgumentException("DataSet must be a record list (a List of a Record or Structure).", nameof(ssDataSet));
            }
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

                if (format.HasValue)
                {
                    Util.ValidateFormatNumbers(format.Value.ssSTCellFormat);
                }
                ClearFormulas(ws, ssRowStart, ssColumnStart, dt, ssExportHeaders);
                ExcelRange range = (ExcelRange)ws.Cells[ssRowStart, ssColumnStart].LoadFromDataTable(dt, ssExportHeaders);
                Util.FormatDateCells(ws, range, dt, ssExportHeaders);

                if (format.HasValue)
                {
                    Util.ApplyFormatToRange(range, format.Value);
                }
            }
        } // WriteRange

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ssWorkbook"></param>
        public void MssWorkbook_Close(object ssWorkbook)
        {
            if (ssWorkbook == null)
            {
                return;
            }
            AsWorkbook(ssWorkbook).Dispose();
        } // MssWorkbook_Close

        /// <summary>
        /// Rename a worksheet
        /// </summary>
        /// <param name="ssWorksheet"></param>
        /// <param name="ssName"></param>
        public void MssWorksheet_Rename(object ssWorksheet, string ssName)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);
            Util.ValidateSheetName(ssName, nameof(ssName));
            ws.Name = ssName;
        } // MssWorksheet_Rename

        /// <summary>
        /// Get the in-memory binary data of the specified workbook
        /// </summary>
        /// <param name="ssWorkbook">The workbook you want the binary data for</param>
        /// <param name="ssBinaryData">The binary data of the file</param>
        public void MssWorkbook_GetBinaryData(object ssWorkbook, out byte[] ssBinaryData)
        {
            ExcelPackage p = AsWorkbook(ssWorkbook);
            Util.PreserveVisibleRowsForZeroHeightSheets(p);
            ssBinaryData = SaveAndReload(p);
        } // MssWorkbook_GetBinaryData

        /// <summary>
        /// Opens an existing workbook for editing by either specifying a name or the binary data.
        /// </summary>
        /// <param name="ssFileName">Location of the file that you want to open. Set to empty string "" when using binary data. URLs on local or private networks and files on network shares are not allowed.</param>
        /// <param name="ssBinary_Data">Binary data of the file that you want to open. Set to nullbinary() if using FileName</param>
        /// <param name="ssWorkbook">The workbook that you want to work with.</param>
        public void MssWorkbook_Open(string ssFileName, byte[] ssBinary_Data, string ssPassword, out object ssWorkbook)
        {
            ssWorkbook = OpenWorkbook(ssFileName, ssBinary_Data, ssPassword, nameof(ssBinary_Data));
        } // MssWorkbook_Open

        /// <summary>
        /// Workbook_Open and Workbook_Open_BinaryData: open from a URL, a file path or the bytes given.
        /// binaryParamName is the calling action's name for the bytes, used in the error for a damaged file.
        /// </summary>
        private static ExcelPackage OpenWorkbook(string fileName, byte[] data, string password, string binaryParamName)
        {
            fileName = fileName ?? "";
            bool hasBinaryData = data != null && data.LongLength > 0;

            if (!hasBinaryData && string.IsNullOrEmpty(fileName))
            {
                throw new ArgumentException("You need to specify at least one of FileName or Binary_Data");
            }

            ExcelPackage p = new ExcelPackage();
            if (fileName.StartsWith("http:", StringComparison.OrdinalIgnoreCase) || fileName.StartsWith("https:", StringComparison.OrdinalIgnoreCase))
            {
                const int timeoutMs = 30000;                 // fail fast instead of hanging on a slow/dead host
                const int overallDeadlineMs = 60000;         // overall deadline across the whole download including redirects
                const int maxRedirects = 5;
                const long maxDownloadBytes = 100L * 1024 * 1024; // cap the download to guard against runaway responses

                System.Diagnostics.Stopwatch stopwatch = System.Diagnostics.Stopwatch.StartNew();
                string currentUrl = fileName;

                for (int redirects = 0; ; redirects++)
                {
                    // Known residual risk: DNS rebinding between our host check and the connection opened by HttpWebRequest.
                    Uri currentUri = ValidateAndResolveHttpUrl(currentUrl);

                    if (stopwatch.ElapsedMilliseconds > overallDeadlineMs)
                    {
                        throw new IOException("Downloading workbook from '" + fileName + "' took longer than 60 seconds.");
                    }

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(currentUri);
                    request.Method = "GET";
                    request.AllowAutoRedirect = false;
                    request.UseDefaultCredentials = false;
                    request.Timeout = timeoutMs;
                    request.ReadWriteTimeout = timeoutMs;

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    {
                        if (stopwatch.ElapsedMilliseconds > overallDeadlineMs)
                        {
                            throw new IOException("Downloading workbook from '" + fileName + "' took longer than 60 seconds.");
                        }

                        int statusCode = (int)response.StatusCode;
                        if (statusCode == 301 || statusCode == 302 || statusCode == 303 || statusCode == 307 || statusCode == 308)
                        {
                            string location = response.Headers["Location"];
                            if (!string.IsNullOrEmpty(location))
                            {
                                if (redirects >= maxRedirects)
                                {
                                    throw new IOException("Too many redirects (more than " + maxRedirects + ") while downloading workbook from '" + fileName + "'.");
                                }
                                if (!Uri.TryCreate(currentUri, location, out Uri nextUri))
                                {
                                    throw new ArgumentException("Invalid redirect URL '" + location + "'.", "ssFileName");
                                }
                                currentUrl = nextUri.AbsoluteUri;
                                continue;
                            }
                        }

                        if (response.StatusCode != HttpStatusCode.OK)
                        {
                            throw new IOException("Failed to download workbook from '" + fileName + "': HTTP " + (int)response.StatusCode + " " + response.StatusCode + ".");
                        }

                        if (response.ContentLength > maxDownloadBytes)
                        {
                            throw new IOException("The downloaded workbook exceeds the maximum allowed size (100 MB).");
                        }

                        using (Stream responseStream = response.GetResponseStream())
                        using (MemoryStream buffer = new MemoryStream())
                        {
                            byte[] chunk = new byte[81920];
                            long total = 0;
                            int read;
                            while ((read = responseStream.Read(chunk, 0, chunk.Length)) > 0)
                            {
                                if (stopwatch.ElapsedMilliseconds > overallDeadlineMs)
                                {
                                    throw new IOException("Downloading workbook from '" + fileName + "' took longer than 60 seconds.");
                                }
                                total += read;
                                if (total > maxDownloadBytes)
                                {
                                    throw new IOException("The downloaded workbook exceeds the maximum allowed size (100 MB).");
                                }
                                buffer.Write(chunk, 0, read);
                            }

                            byte[] raw = buffer.ToArray();
                            Util.EnsureCompleteZipPackage(raw, "ssFileName");
                            byte[] patched = Util.EnsureStylesPartDefaults(raw, "ssFileName");
                            LoadPackage(p, patched, password, "ssFileName");
                            break;
                        }
                    }
                }
            }
            else if (!string.IsNullOrEmpty(fileName))
            {
                const long maxFileBytes = 100L * 1024 * 1024;

                if (fileName.StartsWith(@"\\", StringComparison.Ordinal)
                    || fileName.StartsWith("//", StringComparison.Ordinal)
                    || IsMappedNetworkDrive(fileName))
                {
                    throw new ArgumentException("Workbook_Open does not open files on network shares or devices: '" + fileName + "'.", "ssFileName");
                }

                if (new FileInfo(fileName).Length > maxFileBytes)
                {
                    throw new ArgumentException("The file is larger than 100 MB.", "ssFileName");
                }

                byte[] raw = System.IO.File.ReadAllBytes(fileName);
                Util.EnsureCompleteZipPackage(raw, "ssFileName");
                byte[] patched = Util.EnsureStylesPartDefaults(raw, "ssFileName");
                LoadPackage(p, patched, password, "ssFileName");
            }
            else if (hasBinaryData)
            {
                Util.EnsureCompleteZipPackage(data, binaryParamName);
                byte[] patched = Util.EnsureStylesPartDefaults(data, binaryParamName);
                LoadPackage(p, patched, password, binaryParamName);
            }
            else
            {
                throw new FileNotFoundException("Could not open a file with the given information. Please verify your filename/binary data and try again.");
            }

            return p;
        }

        // Lets the test suite download from its own loopback server; only loopback, every other range stays
        // blocked. Nothing in OutSystems can set it: the tests set it through reflection, which any code in
        // the process could do anyway.
        private static bool allowLoopbackForTests = false;

        private static Uri ValidateAndResolveHttpUrl(string url)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri uri)
                || (!string.Equals(uri.Scheme, "http", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(uri.Scheme, "https", StringComparison.OrdinalIgnoreCase))
                || string.IsNullOrEmpty(uri.Host))
            {
                throw new ArgumentException("Invalid HTTP/HTTPS URL '" + url + "'.", "ssFileName");
            }

            IPAddress[] addresses;
            string host = uri.Host.Trim('[', ']');
            if ((uri.HostNameType == UriHostNameType.IPv4 || uri.HostNameType == UriHostNameType.IPv6)
                && IPAddress.TryParse(host, out IPAddress literalIp))
            {
                addresses = new IPAddress[] { literalIp };
            }
            else
            {
                try
                {
                    addresses = Dns.GetHostAddresses(uri.DnsSafeHost);
                }
                catch (Exception ex)
                {
                    throw new IOException("Could not resolve the host of '" + url + "'.", ex);
                }
                if (addresses == null || addresses.Length == 0)
                {
                    throw new IOException("Could not resolve the host of '" + url + "'.");
                }
            }

            foreach (IPAddress address in addresses)
            {
                if (IsBlockedAddress(address) && !(allowLoopbackForTests && IPAddress.IsLoopback(address)))
                {
                    throw new ArgumentException("Workbook_Open does not open URLs on local or private networks: '" + url + "'.", "ssFileName");
                }
            }

            return uri;
        }

        private static bool IsBlockedAddress(IPAddress a)
        {
            if (a == null)
            {
                return true;
            }

            if (a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
            {
                byte[] b = a.GetAddressBytes();
                if (b[0] == 0) return true;                                      // 0.0.0.0/8
                if (b[0] == 10) return true;                                     // 10.0.0.0/8
                if (b[0] == 100 && (b[1] & 0xC0) == 64) return true;             // 100.64.0.0/10
                if (b[0] == 127) return true;                                    // 127.0.0.0/8
                if (b[0] == 169 && b[1] == 254) return true;                     // 169.254.0.0/16
                if (b[0] == 172 && (b[1] & 0xF0) == 16) return true;             // 172.16.0.0/12
                if (b[0] == 192 && b[1] == 0 && b[2] == 0) return true;          // 192.0.0.0/24
                if (b[0] == 192 && b[1] == 168) return true;                     // 192.168.0.0/16
                if (b[0] == 198 && (b[1] & 0xFE) == 18) return true;             // 198.18.0.0/15
                if ((b[0] & 0xF0) == 224 || (b[0] & 0xF0) == 240) return true;   // 224.0.0.0/4 and 240.0.0.0/4 (includes 255.255.255.255)
                return false;
            }

            if (a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetworkV6)
            {
                byte[] b = a.GetAddressBytes();

                bool first80Zero = true;
                for (int i = 0; i < 10; i++)
                {
                    if (b[i] != 0)
                    {
                        first80Zero = false;
                        break;
                    }
                }

                if (first80Zero && b[10] == 0 && b[11] == 0)
                {
                    // ::/128 (unspecified) or ::1/128 (loopback)
                    if (b[12] == 0 && b[13] == 0 && b[14] == 0 && (b[15] == 0 || b[15] == 1))
                    {
                        return true;
                    }
                    // IPv4-compatible IPv6 address (::a.b.c.d)
                    return IsBlockedAddress(new IPAddress(new byte[] { b[12], b[13], b[14], b[15] }));
                }

                if (first80Zero && b[10] == 0xFF && b[11] == 0xFF)
                {
                    // IPv4-mapped IPv6 address (::ffff:a.b.c.d)
                    return IsBlockedAddress(new IPAddress(new byte[] { b[12], b[13], b[14], b[15] }));
                }

                if (b[0] == 0x00 && b[1] == 0x64 && b[2] == 0xFF && b[3] == 0x9B && AllZero(b, 4, 12))
                {
                    // NAT64 64:ff9b::/96 reaches the embedded IPv4 address
                    return IsBlockedAddress(new IPAddress(new byte[] { b[12], b[13], b[14], b[15] }));
                }
                if (b[0] == 0x20 && b[1] == 0x02)
                {
                    // 6to4 2002::/16 tunnels to the IPv4 address in bytes 2-5
                    return IsBlockedAddress(new IPAddress(new byte[] { b[2], b[3], b[4], b[5] }));
                }
                if ((b[0] & 0xFE) == 0xFC) return true;                          // fc00::/7
                if (b[0] == 0xFE && (b[1] & 0xC0) == 0x80) return true;          // fe80::/10
                if (b[0] == 0xFE && (b[1] & 0xC0) == 0xC0) return true;          // fec0::/10
                if (b[0] == 0xFF) return true;                                   // ff00::/8
                return false;
            }

            return true;
        }

        // True when bytes [from, to) of b are all zero.
        private static bool AllZero(byte[] b, int from, int to)
        {
            for (int i = from; i < to; i++)
            {
                if (b[i] != 0) return false;
            }
            return true;
        }

        private static bool IsMappedNetworkDrive(string fileName)
        {
            if (Path.DirectorySeparatorChar != '\\')
            {
                return false;
            }
            try
            {
                string root = Path.GetPathRoot(Path.GetFullPath(fileName));
                return !string.IsNullOrEmpty(root) && new DriveInfo(root).DriveType == DriveType.Network;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Load workbook bytes into the package, decrypting with the password when one is supplied.
        /// A missing/incorrect password on an encrypted file is surfaced as a clear error.
        /// EPPlus parses most parts lazily, so the workbook, its styles and every worksheet are
        /// forced to be parsed here: a part that is not valid XML fails on open, with the caller's
        /// parameter name, instead of on the next action.
        /// Residual risk (accepted): an encrypted workbook is an OLE compound document, not a zip,
        /// until EPPlus decrypts it inside Load, so the zip limits in Util do not apply to it.
        /// </summary>
        private static void LoadPackage(ExcelPackage p, byte[] data, string password, string paramName)
        {
            var tablesWithoutFilter = Util.CollectTablesWithoutAutoFilter(data);
            var zeroHeightRowRecords = Util.CollectZeroHeightRowRecords(data);
            try
            {
                if (string.IsNullOrEmpty(password))
                {
                    p.Load(new MemoryStream(data));
                }
                else
                {
                    p.Load(new MemoryStream(data), password);
                }
                ForceParse(p);
            }
            catch (System.Xml.XmlException)
            {
                throw new ArgumentException("The workbook is damaged: one of its parts is not valid XML.", paramName);
            }
            catch (Exception ex) when (WrapsXmlException(ex))
            {
                throw new ArgumentException("The workbook is damaged: one of its parts is not valid XML.", paramName);
            }
            catch (System.Security.SecurityException)
            {
                throw new ArgumentException("The workbook is password-protected and the supplied password is incorrect.");
            }
            catch (InvalidDataException)
            {
                // an encrypted (OLE compound) file opened without a password lands here
                if (string.IsNullOrEmpty(password) && data != null && data.Length > 1 && data[0] == 0xD0 && data[1] == 0xCF)
                {
                    throw new ArgumentException("The workbook is password-protected. Supply the Password to open it.");
                }
                throw;
            }
            Util.RestoreTablesWithoutAutoFilter(p, tablesWithoutFilter);
            Util.MarkZeroHeightRowRecords(p, zeroHeightRowRecords);
        }

        /// <summary>
        /// Makes EPPlus parse the parts it normally defers, so broken XML fails here.
        /// Verified against EPPlus 4.5: Worksheets.Count parses xl/workbook.xml and every worksheet
        /// part, Styles.NamedStyles parses xl/styles.xml, and a worksheet's Dimension loads its XML.
        /// </summary>
        private static void ForceParse(ExcelPackage p)
        {
            // the reads are the point of this method
            _ = p.Workbook.Styles.NamedStyles.Count;
            int sheetCount = p.Workbook.Worksheets.Count;
            for (int i = 1; i <= sheetCount; i++)
            {
                _ = p.Workbook.Worksheets[i].Dimension;
            }
        }

        /// <summary>True when the exception, or anything it wraps, is an XML error.</summary>
        private static bool WrapsXmlException(Exception ex)
        {
            for (Exception inner = ex == null ? null : ex.InnerException; inner != null; inner = inner.InnerException)
            {
                if (inner is System.Xml.XmlException)
                {
                    return true;
                }
            }
            return false;
        }

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
        /// Column3D (not supported)
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
        /// StockHLC (not supported)
        /// StockOHLC (not supported)
        /// StockVHLC (not supported)
        /// StockVOHLC (not supported)
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
        /// <param name="ssRowPos">Row of the chart's upper-left corner, starting at 1</param>
        /// <param name="ssColPos">Column of the chart's upper-left corner, starting at 1</param>
        public void MssChart_Create(object ssWorksheet, string ssChartType, string ssChartName, RLDataSeriesRecordList ssDataSeries_List, int ssHeight, int ssWidth, int ssRowPos, int ssColPos)
        {
            ExcelWorksheet ws = AsWorksheet(ssWorksheet);

            // Both positions count from 1 (A1), like the image actions; EPPlus counts from 0.
            if (ssRowPos < 1 || ssRowPos > ExcelPackage.MaxRows)
            {
                throw new ArgumentException("Row position must be between 1 and " + ExcelPackage.MaxRows + ".", nameof(ssRowPos));
            }
            if (ssColPos < 1 || ssColPos > ExcelPackage.MaxColumns)
            {
                throw new ArgumentException("Column position must be between 1 and " + ExcelPackage.MaxColumns + ".", nameof(ssColPos));
            }

            // Every check happens before AddChart, so a rejected call leaves no drawing on the sheet.
            if (ssWidth < 1)
            {
                throw new ArgumentException("The chart width must be at least 1 pixel.", nameof(ssWidth));
            }
            if (ssHeight < 1)
            {
                throw new ArgumentException("The chart height must be at least 1 pixel.", nameof(ssHeight));
            }

            var chartType = Util.StringToChartType(ssChartType);

            RequireFitsInGrid(ws, ssRowPos, 0, ssColPos, 0, ssWidth, ssHeight, "chart", nameof(ssRowPos), nameof(ssColPos));

            var chart = ws.Drawings.AddChart(ssChartName, chartType);
            chart.SetPosition(ssRowPos - 1, 0, ssColPos - 1, 0);
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

        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<ExcelWorksheet, object> _repairedHeaderFooters =
            new System.Runtime.CompilerServices.ConditionalWeakTable<ExcelWorksheet, object>();

        /// <summary>
        /// Return the worksheet's HeaderFooter, repairing EPPlus 4.5's section split from the raw
        /// worksheet XML on the first access for this worksheet instance.
        /// </summary>
        private static ExcelHeaderFooter HeaderFooterOf(ExcelWorksheet ws)
        {
            _repairedHeaderFooters.GetValue(ws, RepairHeaderFooter);
            return ws.HeaderFooter;
        }

        // EPPlus 4.5 mis-splits a header or footer with an empty section when it reads a saved file
        // (&La&C&Rc reads back center "&Rc"), so the sections are split again from the raw XML.
        private static object RepairHeaderFooter(ExcelWorksheet ws)
        {
            var xml = ws.WorksheetXml;
            if (xml?.DocumentElement != null)
            {
                var ns = new System.Xml.XmlNamespaceManager(xml.NameTable);
                ns.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                var hfNode = xml.SelectSingleNode("//d:headerFooter", ns);
                if (hfNode != null)
                {
                    RepairHeaderFooterElement(hfNode, ns, "d:oddHeader", () => ws.HeaderFooter.OddHeader);
                    RepairHeaderFooterElement(hfNode, ns, "d:evenHeader", () => ws.HeaderFooter.EvenHeader);
                    RepairHeaderFooterElement(hfNode, ns, "d:firstHeader", () => ws.HeaderFooter.FirstHeader);
                    RepairHeaderFooterElement(hfNode, ns, "d:oddFooter", () => ws.HeaderFooter.OddFooter);
                    RepairHeaderFooterElement(hfNode, ns, "d:evenFooter", () => ws.HeaderFooter.EvenFooter);
                    RepairHeaderFooterElement(hfNode, ns, "d:firstFooter", () => ws.HeaderFooter.FirstFooter);
                }
            }

            return new object();
        }

        private static void RepairHeaderFooterElement(System.Xml.XmlNode hfNode, System.Xml.XmlNamespaceManager ns, string name, Func<ExcelHeaderFooterText> targetSelector)
        {
            var node = hfNode.SelectSingleNode(name, ns);
            if (node != null)
            {
                Util.SplitHeaderFooterSections(node.InnerText, out string left, out string center, out string right);
                ExcelHeaderFooterText target = targetSelector();
                target.LeftAlignedText = left;
                target.CenteredText = center;
                target.RightAlignedText = right;
            }
        }

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
            ExcelHeaderFooter headerFooter = HeaderFooterOf(ws);
            headerFooter.OddFooter.LeftAlignedText = ssLeftSection;
            headerFooter.OddFooter.CenteredText = ssCenterSection;
            headerFooter.OddFooter.RightAlignedText = ssRightSection;
            headerFooter.EvenFooter.LeftAlignedText = ssLeftSection;
            headerFooter.EvenFooter.CenteredText = ssCenterSection;
            headerFooter.EvenFooter.RightAlignedText = ssRightSection;
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
            ExcelHeaderFooter headerFooter = HeaderFooterOf(ws);
            headerFooter.OddHeader.LeftAlignedText = ssLeftSection;
            headerFooter.OddHeader.CenteredText = ssCenterSection;
            headerFooter.OddHeader.RightAlignedText = ssRightSection;
            headerFooter.EvenHeader.LeftAlignedText = ssLeftSection;
            headerFooter.EvenHeader.CenteredText = ssCenterSection;
            headerFooter.EvenHeader.RightAlignedText = ssRightSection;
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
            ExcelHeaderFooter headerFooter = HeaderFooterOf(ws);
            ssLeftSection = (ssIsEven ? headerFooter.EvenHeader.LeftAlignedText : headerFooter.OddHeader.LeftAlignedText);
            ssCenterSection = (ssIsEven ? headerFooter.EvenHeader.CenteredText : headerFooter.OddHeader.CenteredText);
            ssRightSection = (ssIsEven ? headerFooter.EvenHeader.RightAlignedText : headerFooter.OddHeader.RightAlignedText);
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
            ExcelHeaderFooter headerFooter = HeaderFooterOf(ws);
            ssLeftSection = (ssIsEven ? headerFooter.EvenFooter.LeftAlignedText : headerFooter.OddFooter.LeftAlignedText);
            ssCenterSection = (ssIsEven ? headerFooter.EvenFooter.CenteredText : headerFooter.OddFooter.CenteredText);
            ssRightSection = (ssIsEven ? headerFooter.EvenFooter.RightAlignedText : headerFooter.OddFooter.RightAlignedText);
        } // MssWorksheet_GetFooter

        /// <summary>
        /// Clear value of a cell, defined by its index.
        /// Option to specify whether the cell is part of a merged group or not.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssRow">Row Number</param>
        /// <param name="ssStartColumn">Column Number</param>
        /// <param name="ssEndColumn">Column Number. Mandatory if IsMerged is True</param>
        /// <param name="ssIsMerged">If True, the cells from StartColumn to EndColumn are a merged range: its value is cleared and the range stays merged.</param>
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
        /// Clear the value of a cell, defined by its name.
        /// Option to specify whether the cell is part of a merged group or not.
        /// </summary>
        /// <param name="ssWorksheet">Worksheet on which the cell resides</param>
        /// <param name="ssCellName">Cell-name (eg A1:B1, if cells are merged; eg A1, if single cell)</param>
        /// <param name="ssIsMerged">If True, CellName is a merged range (eg A1:B1): its value is cleared and the range stays merged.</param>
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
        /// Cast the workbook handle to an ExcelPackage, throwing a clear error when another object
        /// (for example a worksheet) is passed as the workbook.
        /// </summary>
        private static ExcelPackage AsWorkbook(object ssWorkbook)
        {
            if (ssWorkbook is ExcelPackage p) { return p; }
            throw new ArgumentException("Expected a Workbook object (e.g. from Workbook_Create or Workbook_Open).", nameof(ssWorkbook));
        }

        /// <summary>
        /// Cast the worksheet handle to an ExcelWorksheet, throwing a clear error instead of
        /// the cryptic NullReferenceException that a failed "as" cast would surface later.
        /// </summary>
        private static ExcelWorksheet AsWorksheet(object ssWorksheet)
        {
            if (ssWorksheet is ExcelWorksheet ws)
            {
                return FollowWorksheetRedirects(ws);
            }
            throw new ArgumentException("Expected a Worksheet object (e.g. from Worksheet_Select, Worksheet_SelectByName/Index, or Workbook_AddName).", nameof(ssWorksheet));
        }

        /// <summary>
        /// Remove the formulas from the cells LoadFromDataTable is about to fill with this data.
        /// </summary>
        private static void ClearFormulas(ExcelWorksheet ws, int fromRow, int fromColumn, DataTable data, bool withHeaders)
        {
            int rows = data.Rows.Count + (withHeaders ? 1 : 0);
            if (rows > 0 && data.Columns.Count > 0)
            {
                ClearFormulas(ws, fromRow, fromColumn, fromRow + rows - 1, fromColumn + data.Columns.Count - 1);
            }
        }

        /// <summary>
        /// Remove the formulas from a range before values are written to it: EPPlus's SetValue and
        /// LoadFromDataTable keep a cell's formula, and Excel recalculates it on open, so the written
        /// value would be lost.
        /// </summary>
        private static void ClearFormulas(ExcelWorksheet ws, int fromRow, int fromColumn, int toRow, int toColumn)
        {
            foreach (ExcelRangeBase cell in ws.Cells[fromRow, fromColumn, toRow, toColumn])
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    cell.Formula = "";
                }
            }
        }

        /// <summary>
        /// Throw a clear ArgumentException when an address is empty or not a cell or range address.
        /// EPPlus accepts such addresses and only fails when the workbook is saved, after which the
        /// workbook can no longer be saved at all. A sheet prefix (Sheet1!A1:B2) is allowed.
        /// </summary>
        private static void RequireAddress(string address, string paramName)
        {
            string a = (address ?? "").Trim();
            int bang = a.LastIndexOf('!');
            bool emptySheet = false;
            if (bang >= 0)
            {
                string sheet = a.Substring(0, bang).Trim();
                if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'')
                {
                    sheet = sheet.Substring(1, sheet.Length - 2);
                }
                emptySheet = sheet.Length == 0;
                a = a.Substring(bang + 1);
            }
            if (emptySheet || a.Length == 0 || !(ExcelCellBase.IsValidAddress(a) || ExcelCellBase.IsValidCellAddress(a)))
            {
                throw new ArgumentException("'" + address + "' is not a valid cell or range address (for example A1 or B2:D10).", paramName);
            }
        }

        /*
         * GetBinaryData/SaveRightToLeft have to reload the package after GetAsByteArray (EPPlus closes
         * it on save), which rebuilds every worksheet object. A worksheet handle selected BEFORE that
         * call would then point at the detached old object model: writes to it are silently lost and
         * some operations crash inside EPPlus. Each reload therefore records old->new worksheet
         * mappings here, and AsWorksheet transparently follows them so stale handles keep working.
         */
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<ExcelWorksheet, ExcelWorksheet> _worksheetRedirects =
            new System.Runtime.CompilerServices.ConditionalWeakTable<ExcelWorksheet, ExcelWorksheet>();

        private static ExcelWorksheet FollowWorksheetRedirects(ExcelWorksheet ws)
        {
            // follow the chain (a workbook may be saved multiple times); guard against cycles
            for (int hops = 0; hops < 64; hops++)
            {
                if (!_worksheetRedirects.TryGetValue(ws, out var next))
                {
                    return ws;
                }
                ws = next;
            }
            return ws;
        }

        /// <summary>
        /// Save the package to a byte array and reload it so the workbook object stays usable,
        /// recording old-to-new worksheet redirects so previously selected worksheet handles
        /// keep pointing at live worksheets.
        /// </summary>
        private static byte[] SaveAndReload(ExcelPackage p, string password = null)
        {
            var oldSheets = new List<ExcelWorksheet>();
            foreach (var ws in p.Workbook.Worksheets)
            {
                oldSheets.Add(ws);
            }

            byte[] bytes = p.GetAsByteArray();
            var tablesWithoutFilter = Util.CollectTablesWithoutAutoFilter(bytes);
            var zeroHeightRowRecords = Util.CollectZeroHeightRowRecords(bytes);
            if (string.IsNullOrEmpty(password))
            {
                p.Load(new System.IO.MemoryStream(bytes));
            }
            else
            {
                p.Load(new System.IO.MemoryStream(bytes), password);
            }
            Util.RestoreTablesWithoutAutoFilter(p, tablesWithoutFilter);
            Util.MarkZeroHeightRowRecords(p, zeroHeightRowRecords);

            foreach (var old in oldSheets)
            {
                ExcelWorksheet fresh = null;
                try { fresh = p.Workbook.Worksheets[old.Name]; } catch { }
                if (fresh != null && !ReferenceEquals(old, fresh))
                {
                    _worksheetRedirects.Remove(old);
                    _worksheetRedirects.Add(old, fresh);
                }
            }

            return bytes;
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
        /// Format a colour as a #RRGGBB hex string (used when reading a data-bar colour back).
        /// </summary>
        private static string ToHexColor(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        /// <summary>
        /// Apply the prompt/error-message and allow-blank settings shared by every data-validation type.
        /// </summary>
        private static void ApplyValidationCommon(OfficeOpenXml.DataValidation.Contracts.IExcelDataValidation val, STDataValidationStructure v)
        {
            val.AllowBlank = v.ssAllowBlank;
            val.ShowErrorMessage = v.ssShowErrorMessage;
            val.ShowInputMessage = v.ssShowPrompt;
            if (!string.IsNullOrEmpty(v.ssErrorStyle))
            {
                val.ErrorStyle = ParseEnum(v.ssErrorStyle, OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop);
            }
            if (!string.IsNullOrEmpty(v.ssErrorTitle)) val.ErrorTitle = v.ssErrorTitle;
            if (!string.IsNullOrEmpty(v.ssErrorMessage)) val.Error = v.ssErrorMessage;
            if (!string.IsNullOrEmpty(v.ssPromptTitle)) val.PromptTitle = v.ssPromptTitle;
            if (!string.IsNullOrEmpty(v.ssPromptMessage)) val.Prompt = v.ssPromptMessage;
        }

    } // CssAdvanced_Excel

} // OutSystems.NssAdvanced_Excel

