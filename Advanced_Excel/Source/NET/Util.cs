using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using OutSystems.HubEdition.RuntimePlatform.Db;
using System.Globalization;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OutSystems.HubEdition.RuntimePlatform;
using OfficeOpenXml.Drawing.Chart;
using System.IO;

namespace OutSystems.NssAdvanced_Excel
{
    class Util
    {
        public static DataTable Transpose(DataTable dt, string ssCellType)
        {
            DataTable dtNew = new DataTable();

            //adding columns    
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataColumn c = dtNew.Columns.Add(i.ToString());
                switch (ssCellType.ToLower())
                {
                    case "integer": c.DataType = typeof(int); break;
                    case "datetime": c.DataType = typeof(DateTime); break;
                    case "decimal": c.DataType = typeof(decimal); break;
                    case "boolean": c.DataType = typeof(bool); break;
                }
            }

            //Adding Row Data
            for (int k = 0; k < dt.Columns.Count; k++)
            {
                DataRow r = dtNew.NewRow();
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    switch (ssCellType.ToLower())
                    {
                        case "integer": r[j] = Convert.ToInt32(dt.Rows[j][k], CultureInfo.InvariantCulture); break;
                        case "datetime": r[j] = Convert.ToDateTime(dt.Rows[j][k], CultureInfo.InvariantCulture); break;
                        case "decimal": r[j] = Convert.ToDecimal(dt.Rows[j][k], CultureInfo.InvariantCulture); break;
                        case "boolean": r[j] = Convert.ToBoolean(dt.Rows[j][k], CultureInfo.InvariantCulture); break;
                        default: r[j] = dt.Rows[j][k]; break;
                    }
                }

                dtNew.Rows.Add(r);
            }

            return dtNew;
        }

        public static byte[] ImageToByteArray(Image imageIn)
        {
            if (imageIn == null)
            {
                return new byte[0];
            }

            using (var ms = new MemoryStream())
            {
                System.Drawing.Imaging.ImageFormat format = System.Drawing.Imaging.ImageFormat.Png;
                try
                {
                    var raw = imageIn.RawFormat;
                    if (raw != null && raw.Guid != System.Drawing.Imaging.ImageFormat.MemoryBmp.Guid)
                    {
                        format = raw;
                    }
                }
                catch
                {
                }

                try
                {
                    imageIn.Save(ms, format);
                }
                catch
                {
                    try
                    {
                        ms.SetLength(0);
                        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    }
                    catch
                    {
                        return new byte[0];
                    }
                }

                return ms.ToArray();
            }
        }
        public static DataTable ConvertArrayListToDataTable(IList<IRecord> arrayList)
        {
            DataTable dt = new DataTable();

            if (arrayList.Count != 0)
            {
                dt = ConvertObjectToDataTableSchema(arrayList[0]);


                FillData(arrayList, dt);
            }

            return dt;
        }

        /// <summary>
        /// Reset the record list, convert it to a DataTable and drop the platform-generated
        /// columns (OptimizedAttributes / OriginalKey). Returns null when the list is empty.
        /// </summary>
        public static DataTable RecordListToDataTable(RecordList rl)
        {
            rl.Reset();
            if (rl.Data.Count == 0)
            {
                return null;
            }

            DataTable dt = ConvertArrayListToDataTable(rl.Data);
            if (dt.Columns.Contains("OptimizedAttributes")) dt.Columns.Remove("OptimizedAttributes");
            if (dt.Columns.Contains("OriginalKey")) dt.Columns.Remove("OriginalKey");
            return dt;
        }

        public static DataTable ConvertObjectToDataTableSchema(Object o)
        {
            DataTable dt = new DataTable();
            // get all fields for given row 
            FieldInfo fieldInfo = o.GetType().GetFields()[0];
            foreach (FieldInfo field in fieldInfo.GetValue(o).GetType().GetFields()) // columns/fields            
            {
                DataColumn dc = new DataColumn(field.Name);
                dc.DataType = field.FieldType;
                dt.Columns.Add(dc);
            }
            return dt;
        }

        private static void FillData(IList<IRecord> arrayList, DataTable dt)
        {
            foreach (Object o in arrayList)
            {
                DataRow dr = dt.NewRow();
                FieldInfo fieldInfo = o.GetType().GetFields()[0];

                DateTime nullDate = new DateTime(1901, 01, 01);
                DateTime d = DateTime.MinValue;

                foreach (FieldInfo field in fieldInfo.GetValue(o).GetType().GetFields()) // columns/fields                
                {
                    if (field.FieldType == typeof(System.DateTime))
                    {
                        d = Convert.ToDateTime(field.GetValue(fieldInfo.GetValue(o)));
                        if (d.CompareTo(nullDate) >= 1) dr[field.Name] = field.GetValue(fieldInfo.GetValue(o));
                    }
                    else
                        dr[field.Name] = field.GetValue(fieldInfo.GetValue(o));
                }
                dt.Rows.Add(dr);
            }
        }

        public static Color ConvertFromColorCode(string colorCode)
        {
            try
            {
                string hex = colorCode.Replace("#", "");
                int argb = unchecked((int)long.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture));

                if (hex.Length <= 6)
                {
                    argb = unchecked((int)((uint)argb | 0xFF000000));
                }

                return Color.FromArgb(argb);
            }
            catch
            {
                // Assume white is the default color (instead of giving an error)
                return Color.White;
            }
        }

        /// <summary>
        /// Worksheets saved with sheetFormatPr/@zeroHeight="1" hide every row by default and rely on
        /// explicit row records to mark the visible ones. EPPlus's save prunes empty row records, which
        /// turns those visible blank rows hidden (content ends up squashed). For such sheets, pin a real
        /// record on each visible row in the used range so it survives the save. Other sheets are untouched.
        /// </summary>
        internal static void PreserveVisibleRowsForZeroHeightSheets(ExcelPackage package)
        {
            if (package == null)
            {
                return;
            }

            foreach (var ws in package.Workbook.Worksheets)
            {
                if (ws.Dimension == null || !WorksheetUsesZeroHeight(ws))
                {
                    continue;
                }

                int end = ws.Dimension.End.Row;
                for (int r = 1; r <= end; r++)
                {
                    if (!ws.Row(r).Hidden)
                    {
                        ws.Row(r).CustomHeight = true;
                    }
                }
            }
        }

        /// <summary>
        /// True when the worksheet's sheetFormatPr has zeroHeight set (all rows hidden by default).
        /// </summary>
        private static bool WorksheetUsesZeroHeight(ExcelWorksheet ws)
        {
            try
            {
                var xml = ws.WorksheetXml;
                if (xml == null || xml.DocumentElement == null)
                {
                    return false;
                }

                var nsmgr = new System.Xml.XmlNamespaceManager(xml.NameTable);
                nsmgr.AddNamespace("d", xml.DocumentElement.NamespaceURI);
                var node = xml.SelectSingleNode("//d:sheetFormatPr", nsmgr) as System.Xml.XmlElement;
                if (node == null)
                {
                    return false;
                }

                string z = node.GetAttribute("zeroHeight");
                return z == "1" || string.Equals(z, "true", System.StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Apply the specified format to a range of cells
        /// </summary>
        /// <param name="range">The range of cells to apply the formatting to</param>
        /// <param name="format">The format to apply to the range of cells</param>
        internal static void ApplyFormatToRange(ExcelRange range, RCCellFormatRecord format)
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
                Color color = ConvertFromColorCode(format.ssSTCellFormat.ssBackgroundColor);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(color);
            }

            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssFontColor))
            {
                Color color = ConvertFromColorCode(format.ssSTCellFormat.ssFontColor);
                range.Style.Font.Color.SetColor(color);
            }

            range.Style.Font.Bold = format.ssSTCellFormat.ssBold;
            range.Style.Font.Italic = format.ssSTCellFormat.ssItalic;
            range.Style.Font.UnderLine = format.ssSTCellFormat.ssUnderline;
            range.Style.Font.Strike = format.ssSTCellFormat.ssStrikethrough;

            /*
             * Deprecated, must use specific styles for BorderBottom,BorderTop,BorderLeft,BorderRight
             */
            if (format.ssSTCellFormat.ssBorderStyle > 0)
            {
                Color borderColor = new Color();
                if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBorderColor))
                {
                    borderColor = ConvertFromColorCode(format.ssSTCellFormat.ssBorderColor);
                }
                ExcelBorderStyle borderStyle = (ExcelBorderStyle)format.ssSTCellFormat.ssBorderStyle;
                range.Style.Border.BorderAround(borderStyle, borderColor);
            }

            /*
             * Border styling. BorderBottom,BorderTop,BorderLeft,BorderRight
             */
            //BorderBottom
            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBorderBottom.ssSTBorderStyle.ssColor) ||
                          format.ssSTCellFormat.ssBorderBottom.ssSTBorderStyle.ssStyle != 0)
            {
                Color borderColor = ConvertFromColorCode(format.ssSTCellFormat.ssBorderBottom.ssSTBorderStyle.ssColor);
                ExcelBorderStyle borderStyle = (ExcelBorderStyle)format.ssSTCellFormat.ssBorderBottom.ssSTBorderStyle.ssStyle;
                range.Style.Border.Bottom.Style = borderStyle;
                range.Style.Border.Bottom.Color.SetColor(borderColor);
            }

            //BorderTop
            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBorderTop.ssSTBorderStyle.ssColor) ||
                          format.ssSTCellFormat.ssBorderTop.ssSTBorderStyle.ssStyle != 0)
            {
                Color borderColor = ConvertFromColorCode(format.ssSTCellFormat.ssBorderTop.ssSTBorderStyle.ssColor);
                ExcelBorderStyle borderStyle = (ExcelBorderStyle)format.ssSTCellFormat.ssBorderTop.ssSTBorderStyle.ssStyle;
                range.Style.Border.Top.Style = borderStyle;
                range.Style.Border.Top.Color.SetColor(borderColor);
            }

            //BorderLeft
            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBorderLeft.ssSTBorderStyle.ssColor) ||
                          format.ssSTCellFormat.ssBorderLeft.ssSTBorderStyle.ssStyle != 0)
            {
                Color borderColor = ConvertFromColorCode(format.ssSTCellFormat.ssBorderLeft.ssSTBorderStyle.ssColor);
                ExcelBorderStyle borderStyle = (ExcelBorderStyle)format.ssSTCellFormat.ssBorderLeft.ssSTBorderStyle.ssStyle;
                range.Style.Border.Left.Style = borderStyle;
                range.Style.Border.Left.Color.SetColor(borderColor);
            }

            //BorderRight
            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssBorderRight.ssSTBorderStyle.ssColor) ||
                          format.ssSTCellFormat.ssBorderRight.ssSTBorderStyle.ssStyle != 0)
            {
                Color borderColor = ConvertFromColorCode(format.ssSTCellFormat.ssBorderRight.ssSTBorderStyle.ssColor);
                ExcelBorderStyle borderStyle = (ExcelBorderStyle)format.ssSTCellFormat.ssBorderRight.ssSTBorderStyle.ssStyle;
                range.Style.Border.Right.Style = borderStyle;
                range.Style.Border.Right.Color.SetColor(borderColor);
            }

            if (format.ssSTCellFormat.ssAutofitColumn)
            {
                range.AutoFitColumns();
            }

            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssNumberFormat))
            {
                range.Style.Numberformat.Format = format.ssSTCellFormat.ssNumberFormat;
            }

            if (format.ssSTCellFormat.ssHorizontalAlignment >= 0)
            {
                range.Style.HorizontalAlignment = (ExcelHorizontalAlignment)format.ssSTCellFormat.ssHorizontalAlignment;
            }

            if (format.ssSTCellFormat.ssVerticalAlignment >= 0)
            {
                range.Style.VerticalAlignment = (ExcelVerticalAlignment)format.ssSTCellFormat.ssVerticalAlignment;
            }

            range.Style.WrapText = format.ssSTCellFormat.ssWrapText;
            range.Style.TextRotation = format.ssSTCellFormat.ssTextRotation;
            range.Style.ShrinkToFit = format.ssSTCellFormat.ssShrinkToFit;
            range.Style.ReadingOrder = (ExcelReadingOrder)format.ssSTCellFormat.ssReadingOrder;
            range.Style.QuotePrefix = format.ssSTCellFormat.ssQuotePrefix;
            range.Style.Locked = format.ssSTCellFormat.ssLocked;
            range.Style.Indent = format.ssSTCellFormat.ssIndent;
            range.Style.Hidden = format.ssSTCellFormat.ssHidden;
        }

        /// <summary>
        /// Apply conditional formatting style to rule style property
        /// </summary>
        /// <param name="style"></param>
        /// <param name="ssStyle"></param>
        internal static void ApplyConditionalFormattingStyle(ExcelDxfStyleConditionalFormatting style, RCConditionalFormatStyleRecord ssStyle)
        {
            ExcelUnderLineType underline = (ExcelUnderLineType)ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssUnderline;
            ExcelBorderStyle bTop = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssStyle;
            ExcelBorderStyle bBottom = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssStyle;
            ExcelBorderStyle bLeft = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssStyle;
            ExcelBorderStyle bRight = (ExcelBorderStyle)ssStyle.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssStyle;
            ExcelFillStyle patternType = (ExcelFillStyle)ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternType;

            style.Border.Bottom.Style = bBottom;

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssColor))
            {
                style.Border.Bottom.Color.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssColor);
            }

            style.Border.Left.Style = bLeft;

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssColor))
            {
                style.Border.Left.Color.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssColor);
            }

            style.Border.Right.Style = bRight;

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssColor))
            {
                style.Border.Right.Color.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssColor);
            }

            style.Border.Top.Style = bTop;

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssColor))
            {
                style.Border.Top.Color.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssColor);
            }

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssBackgroundColor))
            {
                style.Fill.BackgroundColor.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssBackgroundColor);
            }

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternColor))
            {
                style.Fill.PatternColor.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternColor);
            }

            style.Fill.PatternType = patternType;
            style.Font.Bold = ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssBold;
            style.Font.Italic = ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssItalic;

            if (ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssStrike)
            {
                style.Font.Strike = ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssStrike;
            }
            else
            {
                style.Font.Strike = null;
            }

            style.Font.Underline = underline;

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssColor))
            {
                style.Font.Color.Color = ConvertFromColorCode(ssStyle.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssColor);
            }

            if (!string.IsNullOrEmpty(ssStyle.ssSTConditionalFormatStyle.ssNumberFormat))
            {
                style.NumberFormat.Format = ssStyle.ssSTConditionalFormatStyle.ssNumberFormat;
            }
        }

        /// <summary>
        /// Read a conditional-formatting rule's differential style (fill, font, borders, number
        /// format) back into a record. Mirror of <see cref="ApplyConditionalFormattingStyle"/> so
        /// GetAllRules returns the rule's actual style instead of defaults.
        /// </summary>
        internal static RCConditionalFormatStyleRecord ReadConditionalFormattingStyle(ExcelDxfStyleConditionalFormatting style)
        {
            var rec = new RCConditionalFormatStyleRecord(null);
            if (style == null)
            {
                return rec;
            }

            // Fill
            if (style.Fill != null)
            {
                if (style.Fill.PatternType.HasValue)
                {
                    rec.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternType = (int)style.Fill.PatternType.Value;
                }
                if (style.Fill.BackgroundColor != null && style.Fill.BackgroundColor.Color.HasValue)
                {
                    rec.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssBackgroundColor = ToHexRgb(style.Fill.BackgroundColor.Color.Value);
                }
                if (style.Fill.PatternColor != null && style.Fill.PatternColor.Color.HasValue)
                {
                    rec.ssSTConditionalFormatStyle.ssFill.ssSTFillStyle.ssPatternColor = ToHexRgb(style.Fill.PatternColor.Color.Value);
                }
            }

            // Font
            if (style.Font != null)
            {
                if (style.Font.Bold.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssBold = style.Font.Bold.Value;
                if (style.Font.Italic.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssItalic = style.Font.Italic.Value;
                if (style.Font.Strike.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssStrike = style.Font.Strike.Value;
                if (style.Font.Underline.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssUnderline = (int)style.Font.Underline.Value;
                if (style.Font.Color != null && style.Font.Color.Color.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssColor = ToHexRgb(style.Font.Color.Color.Value);
            }

            // Borders
            if (style.Border != null)
            {
                if (style.Border.Top.Style.HasValue) rec.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssStyle = (int)style.Border.Top.Style.Value;
                if (style.Border.Top.Color != null && style.Border.Top.Color.Color.HasValue) rec.ssSTConditionalFormatStyle.ssBorderTop.ssSTBorderStyle.ssColor = ToHexRgb(style.Border.Top.Color.Color.Value);

                if (style.Border.Bottom.Style.HasValue) rec.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssStyle = (int)style.Border.Bottom.Style.Value;
                if (style.Border.Bottom.Color != null && style.Border.Bottom.Color.Color.HasValue) rec.ssSTConditionalFormatStyle.ssBorderBottom.ssSTBorderStyle.ssColor = ToHexRgb(style.Border.Bottom.Color.Color.Value);

                if (style.Border.Left.Style.HasValue) rec.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssStyle = (int)style.Border.Left.Style.Value;
                if (style.Border.Left.Color != null && style.Border.Left.Color.Color.HasValue) rec.ssSTConditionalFormatStyle.ssBorderLeft.ssSTBorderStyle.ssColor = ToHexRgb(style.Border.Left.Color.Color.Value);

                if (style.Border.Right.Style.HasValue) rec.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssStyle = (int)style.Border.Right.Style.Value;
                if (style.Border.Right.Color != null && style.Border.Right.Color.Color.HasValue) rec.ssSTConditionalFormatStyle.ssBorderRight.ssSTBorderStyle.ssColor = ToHexRgb(style.Border.Right.Color.Color.Value);
            }

            // Number format
            if (style.NumberFormat != null && !string.IsNullOrEmpty(style.NumberFormat.Format))
            {
                rec.ssSTConditionalFormatStyle.ssNumberFormat = style.NumberFormat.Format;
            }

            return rec;
        }

        private static string ToHexRgb(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dimension"></param>
        /// <returns></returns>
        internal static RCDimensionRecord CastDimension(ExcelAddressBase dimension)
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
        internal static RCAddressRecord CastAddress(ExcelCellAddress address)
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
        /// Log a message to the General Log
        /// </summary>
        /// <param name="message">What to log</param>
        internal static void LogMessage(object message)
        {
            try
            {
                GenericExtendedActions.LogMessage(AppInfo.GetAppInfo().OsContext, message?.ToString() ?? "", "AdvXL");
            }
            catch
            {
                // Logging must never throw (e.g. when no OutSystems runtime context is available).
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chartType"></param>
        /// <returns></returns>
        internal static eChartType StringToChartType(string chartType)
        {
            switch (chartType)
            {
                case "Area3D": return eChartType.Area3D;
                case "AreaStacked3D": return eChartType.AreaStacked3D;
                case "AreaStacked1003D": return eChartType.AreaStacked1003D;
                case "BarClustered3D": return eChartType.BarClustered3D;
                case "BarStacked3D": return eChartType.BarStacked3D;
                case "BarStacked1003D": return eChartType.BarStacked1003D;
                case "Column3D": return eChartType.Column3D;
                case "ColumnClustered3D": return eChartType.ColumnClustered3D;
                case "ColumnStacked3D": return eChartType.ColumnStacked3D;
                case "ColumnStacked1003D": return eChartType.ColumnStacked1003D;
                case "Line3D": return eChartType.Line3D;
                case "Pie3D": return eChartType.Pie3D;
                case "PieExploded3D": return eChartType.PieExploded3D;
                case "Area": return eChartType.Area;
                case "AreaStacked": return eChartType.AreaStacked;
                case "AreaStacked100": return eChartType.AreaStacked100;
                case "BarClustered": return eChartType.BarClustered;
                case "BarOfPie": return eChartType.BarOfPie;
                case "BarStacked": return eChartType.BarStacked;
                case "BarStacked100": return eChartType.BarStacked100;
                case "Bubble": return eChartType.Bubble;
                case "Bubble3DEffect": return eChartType.Bubble3DEffect;
                case "ColumnClustered": return eChartType.ColumnClustered;
                case "ColumnStacked": return eChartType.ColumnStacked;
                case "ColumnStacked100": return eChartType.ColumnStacked100;
                case "ConeBarClustered": return eChartType.ConeBarClustered;
                case "ConeBarStacked": return eChartType.ConeBarStacked;
                case "ConeBarStacked100": return eChartType.ConeBarStacked100;
                case "ConeCol": return eChartType.ConeCol;
                case "ConeColClustered": return eChartType.ConeColClustered;
                case "ConeColStacked": return eChartType.ConeColStacked;
                case "ConeColStacked100": return eChartType.ConeColStacked100;
                case "CylinderBarClustered": return eChartType.CylinderBarClustered;
                case "CylinderBarStacked": return eChartType.CylinderBarStacked;
                case "CylinderBarStacked100": return eChartType.CylinderBarStacked100;
                case "CylinderCol": return eChartType.CylinderCol;
                case "CylinderColClustered": return eChartType.CylinderColClustered;
                case "CylinderColStacked": return eChartType.CylinderColStacked;
                case "CylinderColStacked100": return eChartType.CylinderColStacked100;
                case "Doughnut": return eChartType.Doughnut;
                case "DoughnutExploded": return eChartType.DoughnutExploded;
                case "Line": return eChartType.Line;
                case "LineMarkers": return eChartType.LineMarkers;
                case "LineMarkersStacked": return eChartType.LineMarkersStacked;
                case "LineMarkersStacked100": return eChartType.LineMarkersStacked100;
                case "LineStacked": return eChartType.LineStacked;
                case "LineStacked100": return eChartType.LineStacked100;
                case "Pie": return eChartType.Pie;
                case "PieExploded": return eChartType.PieExploded;
                case "PieOfPie": return eChartType.PieOfPie;
                case "PyramidBarClustered": return eChartType.PyramidBarClustered;
                case "PyramidBarStacked": return eChartType.PyramidBarStacked;
                case "PyramidBarStacked100": return eChartType.PyramidBarStacked100;
                case "PyramidCol": return eChartType.PyramidCol;
                case "PyramidColClustered": return eChartType.PyramidColClustered;
                case "PyramidColStacked": return eChartType.PyramidColStacked;
                case "PyramidColStacked100": return eChartType.PyramidColStacked100;
                case "Radar": return eChartType.Radar;
                case "RadarFilled": return eChartType.RadarFilled;
                case "RadarMarkers": return eChartType.RadarMarkers;
                case "StockHLC": return eChartType.StockHLC;
                case "StockOHLC": return eChartType.StockOHLC;
                case "StockVHLC": return eChartType.StockVHLC;
                case "StockVOHLC": return eChartType.StockVOHLC;
                case "Surface": return eChartType.Surface;
                case "SurfaceTopView": return eChartType.SurfaceTopView;
                case "SurfaceTopViewWireframe": return eChartType.SurfaceTopViewWireframe;
                case "SurfaceWireframe": return eChartType.SurfaceWireframe;
                case "XYScatter": return eChartType.XYScatter;
                case "XYScatterLines": return eChartType.XYScatterLines;
                case "XYScatterLinesNoMarkers": return eChartType.XYScatterLinesNoMarkers;
                case "XYScatterSmooth": return eChartType.XYScatterSmooth;
                case "XYScatterSmoothNoMarkers": return eChartType.XYScatterSmoothNoMarkers;
            }
            return eChartType.Column3D;
        }

    }
}
