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

namespace OutSystems.NssAdvanced_Excel
{
    class Util
    {
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
                return Color.FromArgb(Int32.Parse(colorCode.Replace("#", ""), NumberStyles.HexNumber));
            }
            catch
            {
                // Assume white is the default color (instead of giving an error)
                return Color.White;
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
            GenericExtendedActions.LogMessage(AppInfo.GetAppInfo().OsContext, message.ToString(), "AdvXL");
        }


    }
}
