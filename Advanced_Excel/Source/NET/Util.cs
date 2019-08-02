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

            if (!string.IsNullOrEmpty(format.ssSTCellFormat.ssNumberFormat))
            {
                range.Style.Numberformat.Format = format.ssSTCellFormat.ssNumberFormat;
            }
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


        /// <summary>
        /// 
        /// </summary>
        /// <param name="chartType"></param>
        /// <returns></returns>
        internal static eChartType stringToChartType(string chartType)
        {
            switch (chartType)
            {
                case "Area3D": return eChartType.Area3D;
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


    }
}
