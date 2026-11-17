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

        /// <summary>
        /// A date written with number format General shows in Excel as its serial number (45292).
        /// Give the DateTime columns of a range just loaded from dt Excel's built-in date format.
        /// </summary>
        internal static void FormatDateCells(ExcelWorksheet ws, ExcelRangeBase range, DataTable dt, bool hasHeaders)
        {
            int fromRow = range.Start.Row + (hasHeaders ? 1 : 0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                if (dt.Columns[i].DataType == typeof(DateTime))
                {
                    int column = range.Start.Column + i;
                    FormatDateCells(ws, fromRow, column, range.End.Row, column);
                }
            }
        }

        /// <summary>
        /// Give the cells of the area that hold a date and have number format General Excel's built-in
        /// date format (14), or date and time (22) in a column where a date has a time. Excel shows both
        /// in the user's regional format. A cell with its own number format keeps it.
        /// </summary>
        internal static void FormatDateCells(ExcelWorksheet ws, int fromRow, int fromColumn, int toRow, int toColumn)
        {
            for (int c = fromColumn; c <= toColumn; c++)
            {
                var rows = new List<int>();
                bool hasTime = false;
                for (int r = fromRow; r <= toRow; r++)
                {
                    object value = ws.Cells[r, c].Value;
                    if (value is DateTime && ws.Cells[r, c].Style.Numberformat.NumFmtID == 0)
                    {
                        rows.Add(r);
                        hasTime |= ((DateTime)value).TimeOfDay != TimeSpan.Zero;
                    }
                }
                string format = hasTime ? "m/d/yy h:mm" : "mm-dd-yy";
                // one style change per run of consecutive rows, not per cell
                for (int i = 0; i < rows.Count; i++)
                {
                    int start = rows[i];
                    while (i + 1 < rows.Count && rows[i + 1] == rows[i] + 1)
                    {
                        i++;
                    }
                    ws.Cells[start, c, rows[i], c].Style.Numberformat.Format = format;
                }
            }
        }

        /// <summary>
        /// An image for EPPlus to keep. GDI+ needs the stream for as long as the image lives, so the
        /// stream is left open and goes away with the image.
        /// </summary>
        public static Image ImageFromBytes(byte[] bytes)
        {
            return Image.FromStream(new MemoryStream(bytes));
        }

        /// <summary>
        /// True when a worksheet other than ws is visible.
        /// </summary>
        public static bool HasOtherVisibleSheet(ExcelWorksheet ws)
        {
            foreach (ExcelWorksheet other in ws.Workbook.Worksheets)
            {
                if (other != ws && other.Hidden == eWorkSheetHidden.Visible)
                {
                    return true;
                }
            }
            return false;
        }

        private static System.Xml.XmlElement WorkbookProtectionNode(ExcelWorkbook wb)
        {
            System.Xml.XmlElement root = wb.WorkbookXml.DocumentElement;
            foreach (System.Xml.XmlNode node in root.ChildNodes)
            {
                if (node.LocalName == "workbookProtection" && node.NamespaceURI == root.NamespaceURI)
                {
                    return (System.Xml.XmlElement)node;
                }
            }
            return null;
        }

        /// <summary>
        /// The workbook's password hash (the workbookPassword attribute), or null when it has none.
        /// </summary>
        public static string WorkbookPasswordHash(ExcelWorkbook wb)
        {
            System.Xml.XmlElement node = WorkbookProtectionNode(wb);
            return node != null && node.HasAttribute("workbookPassword") ? node.GetAttribute("workbookPassword") : null;
        }

        /// <summary>
        /// Puts back a password hash read with WorkbookPasswordHash. The workbookProtection element
        /// must exist, which it does once a lock is set.
        /// </summary>
        public static void SetWorkbookPasswordHash(ExcelWorkbook wb, string hash)
        {
            WorkbookProtectionNode(wb).SetAttribute("workbookPassword", hash);
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
        /// When the input starts with the zip signature PK (0x50 0x4B), verify that the archive
        /// is a complete zip package by reading its central directory and streaming every entry to
        /// the end. Non-zip inputs (such as encrypted OLE compound workbooks) are untouched.
        /// </summary>
        internal static void EnsureCompleteZipPackage(byte[] bytes, string paramName)
        {
            if (bytes == null || bytes.Length < 2 || bytes[0] != 0x50 || bytes[1] != 0x4B)
            {
                return;
            }

            try
            {
                using (var ms = new MemoryStream(bytes, false))
                using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read))
                {
                    if (zip.Entries.Count == 0)
                    {
                        throw new InvalidDataException("Empty zip archive.");
                    }
                    byte[] buffer = new byte[8192];
                    foreach (var entry in zip.Entries)
                    {
                        long expected = entry.Length;
                        long total = 0;
                        using (var s = entry.Open())
                        {
                            int n;
                            while ((n = s.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                total += n;
                            }
                        }
                        if (total != expected)
                        {
                            throw new InvalidDataException("Truncated zip entry.");
                        }
                    }
                }
            }
            catch (Exception)
            {
                throw new ArgumentException("The file is not a complete .xlsx workbook: it is damaged or was cut off.", paramName);
            }
        }

        /// <summary>
        /// Workbooks saved without required style lists (fonts, fills, borders, cellStyleXfs, cellXfs) or
        /// without a &lt;dxfs&gt; element (e.g. an empty &lt;styleSheet/&gt; part or files from minimal generators) make
        /// EPPlus throw a NullReferenceException when cells or conditional formats are written and saved.
        /// Inject any missing lists with Excel's defaults in CT_Stylesheet schema order, preserving the
        /// root element's namespace prefix. Returns the original bytes unchanged when all lists exist,
        /// the input isn't an xlsx package, or anything goes wrong (opening must never be blocked by this).
        /// </summary>
        internal static byte[] EnsureStylesPartDefaults(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 4 || bytes[0] != 0x50 || bytes[1] != 0x4B)
            {
                return bytes;
            }

            try
            {
                using (var srcMs = new MemoryStream(bytes, false))
                using (var srcZip = new System.IO.Compression.ZipArchive(srcMs, System.IO.Compression.ZipArchiveMode.Read))
                {
                    var stylesEntry = srcZip.GetEntry("xl/styles.xml");
                    if (stylesEntry == null)
                    {
                        return bytes;
                    }
                    string xml;
                    using (var r = new StreamReader(stylesEntry.Open()))
                    {
                        xml = r.ReadToEnd();
                    }
                    string patched = EnsureStylesPartLists(xml);
                    if (patched == xml)
                    {
                        return bytes;
                    }

                    using (var dstMs = new MemoryStream())
                    {
                        using (var dstZip = new System.IO.Compression.ZipArchive(dstMs, System.IO.Compression.ZipArchiveMode.Create, true))
                        {
                            foreach (var entry in srcZip.Entries)
                            {
                                var newEntry = dstZip.CreateEntry(entry.FullName, System.IO.Compression.CompressionLevel.Optimal);
                                newEntry.LastWriteTime = entry.LastWriteTime;
                                if (entry.FullName == "xl/styles.xml")
                                {
                                    using (var w = new StreamWriter(newEntry.Open(), new System.Text.UTF8Encoding(false)))
                                    {
                                        w.Write(patched);
                                    }
                                }
                                else
                                {
                                    using (var srcStream = entry.Open())
                                    using (var dstStream = newEntry.Open())
                                    {
                                        srcStream.CopyTo(dstStream);
                                    }
                                }
                            }
                        }
                        return dstMs.ToArray();
                    }
                }
            }
            catch
            {
                return bytes;
            }
        }

        private static readonly string[] StyleSheetSchemaOrder = {
            "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
            "cellStyles", "dxfs", "tableStyles", "colors", "extLst"
        };

        private static readonly string[] RequiredStyleLists = {
            "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "dxfs"
        };

        /// <summary>
        /// Ensure xl/styles.xml contains fonts, fills, borders, cellStyleXfs, cellXfs and dxfs in
        /// CT_Stylesheet sequence, using the root styleSheet element's namespace and prefix.
        /// </summary>
        private static string EnsureStylesPartLists(string stylesXml)
        {
            var doc = new System.Xml.XmlDocument { PreserveWhitespace = true };
            using (var reader = System.Xml.XmlReader.Create(
                new StringReader(stylesXml),
                new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null }))
            {
                doc.Load(reader);
            }
            var root = doc.DocumentElement;
            if (root == null || root.LocalName != "styleSheet")
            {
                return stylesXml;
            }

            bool changed = false;
            foreach (string listName in RequiredStyleLists)
            {
                var existing = FindStyleChild(root, listName);
                if (existing == null)
                {
                    var el = CreateDefaultStyleList(doc, root.Prefix, root.NamespaceURI, listName);
                    InsertInStyleSchemaOrder(root, el);
                    changed = true;
                }
                else if (EnsureDefaultStyleListChildren(doc, existing, root.Prefix, root.NamespaceURI, listName))
                {
                    changed = true;
                }
            }

            return changed ? doc.OuterXml : stylesXml;
        }

        private static System.Xml.XmlElement FindStyleChild(System.Xml.XmlElement root, string localName)
        {
            foreach (System.Xml.XmlNode child in root.ChildNodes)
            {
                if (child.NodeType == System.Xml.XmlNodeType.Element && child.LocalName == localName)
                {
                    return (System.Xml.XmlElement)child;
                }
            }
            return null;
        }

        private static int CountStyleElementChildren(System.Xml.XmlElement list)
        {
            int count = 0;
            foreach (System.Xml.XmlNode child in list.ChildNodes)
            {
                if (child.NodeType == System.Xml.XmlNodeType.Element)
                {
                    count++;
                }
            }
            return count;
        }

        private static int GetStyleSchemaRank(string localName)
        {
            for (int i = 0; i < StyleSheetSchemaOrder.Length; i++)
            {
                if (StyleSheetSchemaOrder[i] == localName)
                {
                    return i;
                }
            }
            return StyleSheetSchemaOrder.Length;
        }

        private static void InsertInStyleSchemaOrder(System.Xml.XmlElement root, System.Xml.XmlElement newChild)
        {
            int targetRank = GetStyleSchemaRank(newChild.LocalName);
            foreach (System.Xml.XmlNode child in root.ChildNodes)
            {
                if (child.NodeType == System.Xml.XmlNodeType.Element && GetStyleSchemaRank(child.LocalName) > targetRank)
                {
                    root.InsertBefore(newChild, child);
                    return;
                }
            }
            root.AppendChild(newChild);
        }

        private static System.Xml.XmlElement CreateStyleElement(System.Xml.XmlDocument doc, string prefix, string ns, string localName)
        {
            return string.IsNullOrEmpty(prefix)
                ? doc.CreateElement(localName, ns)
                : doc.CreateElement(prefix, localName, ns);
        }

        private static System.Xml.XmlElement CreateDefaultStyleList(System.Xml.XmlDocument doc, string prefix, string ns, string listName)
        {
            var list = CreateStyleElement(doc, prefix, ns, listName);
            if (listName == "dxfs")
            {
                list.SetAttribute("count", "0");
            }
            else
            {
                EnsureDefaultStyleListChildren(doc, list, prefix, ns, listName);
            }
            return list;
        }

        private static bool EnsureDefaultStyleListChildren(System.Xml.XmlDocument doc, System.Xml.XmlElement list, string prefix, string ns, string listName)
        {
            int existingCount = CountStyleElementChildren(list);
            switch (listName)
            {
                case "fonts":
                {
                    if (existingCount > 0) return false;
                    list.SetAttribute("count", "1");
                    var font = CreateStyleElement(doc, prefix, ns, "font");
                    var sz = CreateStyleElement(doc, prefix, ns, "sz");
                    sz.SetAttribute("val", "11");
                    var name = CreateStyleElement(doc, prefix, ns, "name");
                    name.SetAttribute("val", "Calibri");
                    font.AppendChild(sz);
                    font.AppendChild(name);
                    list.AppendChild(font);
                    return true;
                }
                case "fills":
                {
                    if (existingCount >= 2) return false;
                    string[] defaults = { "none", "gray125" };
                    for (int i = existingCount; i < defaults.Length; i++)
                    {
                        var fill = CreateStyleElement(doc, prefix, ns, "fill");
                        var patternFill = CreateStyleElement(doc, prefix, ns, "patternFill");
                        patternFill.SetAttribute("patternType", defaults[i]);
                        fill.AppendChild(patternFill);
                        list.AppendChild(fill);
                    }
                    list.SetAttribute("count", "2");
                    return true;
                }
                case "borders":
                {
                    if (existingCount > 0) return false;
                    list.SetAttribute("count", "1");
                    var border = CreateStyleElement(doc, prefix, ns, "border");
                    foreach (string side in new[] { "left", "right", "top", "bottom", "diagonal" })
                    {
                        border.AppendChild(CreateStyleElement(doc, prefix, ns, side));
                    }
                    list.AppendChild(border);
                    return true;
                }
                case "cellStyleXfs":
                {
                    if (existingCount > 0) return false;
                    list.SetAttribute("count", "1");
                    var xf = CreateStyleElement(doc, prefix, ns, "xf");
                    xf.SetAttribute("numFmtId", "0");
                    xf.SetAttribute("fontId", "0");
                    xf.SetAttribute("fillId", "0");
                    xf.SetAttribute("borderId", "0");
                    list.AppendChild(xf);
                    return true;
                }
                case "cellXfs":
                {
                    if (existingCount > 0) return false;
                    list.SetAttribute("count", "1");
                    var xf = CreateStyleElement(doc, prefix, ns, "xf");
                    xf.SetAttribute("numFmtId", "0");
                    xf.SetAttribute("fontId", "0");
                    xf.SetAttribute("fillId", "0");
                    xf.SetAttribute("borderId", "0");
                    xf.SetAttribute("xfId", "0");
                    list.AppendChild(xf);
                    return true;
                }
                default:
                    return false;
            }
        }

        /// <summary>
        /// Collect the table parts (xl/tables/*.xml) whose &lt;table&gt; has a header row
        /// (headerRowCount absent or not "0") and no &lt;autoFilter&gt; child element, so their
        /// ShowFilter=false state can be restored after EPPlus loads the package.
        /// </summary>
        internal static HashSet<string> CollectTablesWithoutAutoFilter(byte[] bytes)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (bytes == null || bytes.Length < 4 || bytes[0] != 0x50 || bytes[1] != 0x4B)
            {
                return result;
            }

            try
            {
                using (var ms = new MemoryStream(bytes, false))
                using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read))
                {
                    foreach (var entry in zip.Entries)
                    {
                        string path = entry.FullName.Replace('\\', '/').TrimStart('/');
                        if (!path.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase)
                            || !path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                            || path.IndexOf('/', "xl/tables/".Length) >= 0)
                        {
                            continue;
                        }

                        var doc = new System.Xml.XmlDocument();
                        using (var s = entry.Open())
                        using (var reader = System.Xml.XmlReader.Create(
                            s,
                            new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null }))
                        {
                            doc.Load(reader);
                        }

                        var root = doc.DocumentElement;
                        if (root == null || root.LocalName != "table")
                        {
                            continue;
                        }

                        if (root.HasAttribute("headerRowCount") && root.GetAttribute("headerRowCount") == "0")
                        {
                            continue;
                        }

                        if (FindStyleChild(root, "autoFilter") == null)
                        {
                            result.Add(path);
                            string tableName = root.GetAttribute("name");
                            if (!string.IsNullOrEmpty(tableName))
                            {
                                result.Add(tableName);
                            }
                        }
                    }
                }
            }
            catch
            {
                // Any error in this step must not block opening.
            }

            return result;
        }

        /// <summary>
        /// After a load, turn the filter off again on the tables CollectTablesWithoutAutoFilter found:
        /// EPPlus re-creates the autoFilter of every table with a header row when it loads one.
        /// </summary>
        internal static void RestoreTablesWithoutAutoFilter(ExcelPackage package, HashSet<string> tablesWithoutAutoFilter)
        {
            if (package == null || tablesWithoutAutoFilter == null || tablesWithoutAutoFilter.Count == 0)
            {
                return;
            }

            try
            {
                foreach (var ws in package.Workbook.Worksheets)
                {
                    foreach (var tbl in ws.Tables)
                    {
                        string partPath = tbl.TableUri != null ? tbl.TableUri.OriginalString.Replace('\\', '/').TrimStart('/') : null;
                        if ((partPath != null && tablesWithoutAutoFilter.Contains(partPath))
                            || (!string.IsNullOrEmpty(tbl.Name) && tablesWithoutAutoFilter.Contains(tbl.Name)))
                        {
                            tbl.ShowFilter = false;
                        }
                    }
                }
            }
            catch
            {
                // Any error in this step must not block opening.
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

            ValidateFormatNumbers(format.ssSTCellFormat);

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
        /// Reject alignment and border-style numbers EPPlus has no value for, before any cell is changed.
        /// EPPlus stored them as they were and the next save threw NullReferenceException, so the workbook
        /// could no longer be saved. The numbers that mean "leave as is" stay accepted: a negative alignment
        /// and a legacy BorderStyle of 0 or less. The write actions call it before writing any value.
        /// </summary>
        internal static void ValidateFormatNumbers(STCellFormatStructure f)
        {
            if (f.ssHorizontalAlignment >= 0) RequireEnumNumber(typeof(ExcelHorizontalAlignment), f.ssHorizontalAlignment, "HorizontalAlignment");
            if (f.ssVerticalAlignment >= 0) RequireEnumNumber(typeof(ExcelVerticalAlignment), f.ssVerticalAlignment, "VerticalAlignment");
            if (f.ssBorderStyle > 0) RequireEnumNumber(typeof(ExcelBorderStyle), f.ssBorderStyle, "BorderStyle");
            RequireEnumNumber(typeof(ExcelBorderStyle), f.ssBorderBottom.ssSTBorderStyle.ssStyle, "BorderBottom.Style");
            RequireEnumNumber(typeof(ExcelBorderStyle), f.ssBorderTop.ssSTBorderStyle.ssStyle, "BorderTop.Style");
            RequireEnumNumber(typeof(ExcelBorderStyle), f.ssBorderLeft.ssSTBorderStyle.ssStyle, "BorderLeft.Style");
            RequireEnumNumber(typeof(ExcelBorderStyle), f.ssBorderRight.ssSTBorderStyle.ssStyle, "BorderRight.Style");
        }

        /// <summary>
        /// Throw a clear ArgumentException when the number is not one of the enum's values (they run from 0 up).
        /// </summary>
        private static void RequireEnumNumber(Type enumType, int value, string name)
        {
            if (!Enum.IsDefined(enumType, value))
            {
                int last = Enum.GetValues(enumType).Length - 1;
                throw new ArgumentException(name + " " + value.ToString(CultureInfo.InvariantCulture)
                    + " is not valid. Use a number from 0 to " + last.ToString(CultureInfo.InvariantCulture) + ".");
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

            if (underline != ExcelUnderLineType.None)
            {
                style.Font.Underline = underline;
            }
            else
            {
                style.Font.Underline = null;
            }

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
        /// <param name="ws">The worksheet that holds the rule (its workbook holds the rule's dxf)</param>
        /// <param name="rule">The rule to read</param>
        internal static RCConditionalFormatStyleRecord ReadConditionalFormattingStyle(ExcelWorksheet ws, OfficeOpenXml.ConditionalFormatting.Contracts.IExcelConditionalFormattingRule rule)
        {
            var rec = new RCConditionalFormatStyleRecord(null);
            ExcelDxfStyleConditionalFormatting style = rule.Style;
            if (style == null)
            {
                return rec;
            }
            System.Xml.XmlElement dxfFont = DxfFontNode(ws, rule);

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
                else if (HasFlagWithoutVal(dxfFont, "b")) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssBold = true;
                if (style.Font.Italic.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssItalic = style.Font.Italic.Value;
                else if (HasFlagWithoutVal(dxfFont, "i")) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssItalic = true;
                if (style.Font.Strike.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssStrike = style.Font.Strike.Value;
                else if (HasFlagWithoutVal(dxfFont, "strike")) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssStrike = true;
                if (style.Font.Underline.HasValue) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssUnderline = (int)style.Font.Underline.Value;
                else if (HasFlagWithoutVal(dxfFont, "u")) rec.ssSTConditionalFormatStyle.ssFont.ssSTFontStyle.ssUnderline = (int)ExcelUnderLineType.Single;
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

        /// <summary>
        /// The font element of the dxf a rule loaded from a file points to, or null (a rule added in this
        /// session has no dxfId until the workbook is saved). EPPlus reads the spec's short form of a font
        /// flag (&lt;b/&gt;, &lt;i/&gt;, &lt;strike/&gt;, &lt;u/&gt;: no val means on, as Excel writes it) as not set,
        /// so <see cref="ReadConditionalFormattingStyle"/> reads those flags here.
        /// </summary>
        private static System.Xml.XmlElement DxfFontNode(ExcelWorksheet ws, OfficeOpenXml.ConditionalFormatting.Contracts.IExcelConditionalFormattingRule rule)
        {
            var cfRule = ((OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule)rule).Node as System.Xml.XmlElement;
            int dxfId;
            if (cfRule == null || !int.TryParse(cfRule.GetAttribute("dxfId"), NumberStyles.None, CultureInfo.InvariantCulture, out dxfId))
            {
                return null;
            }
            var styles = ws.Workbook.StylesXml;
            var ns = new System.Xml.XmlNamespaceManager(styles.NameTable);
            ns.AddNamespace("d", styles.DocumentElement.NamespaceURI);
            return styles.SelectSingleNode("/d:styleSheet/d:dxfs/d:dxf[" + (dxfId + 1).ToString(CultureInfo.InvariantCulture) + "]/d:font", ns) as System.Xml.XmlElement;
        }

        /// <summary>
        /// True when the font element has the named flag child without a val attribute (which means on).
        /// </summary>
        private static bool HasFlagWithoutVal(System.Xml.XmlElement font, string name)
        {
            if (font == null)
            {
                return false;
            }
            foreach (System.Xml.XmlNode child in font.ChildNodes)
            {
                var e = child as System.Xml.XmlElement;
                if (e != null && e.LocalName == name && !e.HasAttribute("val"))
                {
                    return true;
                }
            }
            return false;
        }

        private static string ToHexRgb(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        /// <summary>
        /// Copying a worksheet into another workbook carries its conditional-formatting rules but not
        /// the differential styles (dxfs) they reference, so the destination rules lose their format.
        /// Re-read each source rule's style and re-apply it to the matching copied rule, which makes
        /// EPPlus materialise a fresh dxf in the destination workbook. Colour-scale / icon-set / data-bar
        /// rules don't use a dxf and are left alone (their colours travel in the rule itself).
        /// </summary>
        internal static void PreserveConditionalFormattingStyles(ExcelWorksheet source, ExcelWorksheet copy)
        {
            if (source == null || copy == null)
            {
                return;
            }

            if (source.Workbook != copy.Workbook)
            {
                RemoveDxfIds(copy);
            }

            try
            {
                int count = Math.Min(source.ConditionalFormatting.Count, copy.ConditionalFormatting.Count);
                for (int i = 0; i < count; i++)
                {
                    var src = source.ConditionalFormatting[i];
                    if (!RuleUsesDxfStyle(src.Type))
                    {
                        continue;
                    }

                    var styleRecord = ReadConditionalFormattingStyle(source, src);
                    ApplyConditionalFormattingStyle(copy.ConditionalFormatting[i].Style, styleRecord);
                }
            }
            catch
            {
                // never block a worksheet copy because of this
            }
        }

        /// <summary>
        /// A copied rule keeps its dxfId, an index into the source workbook's dxfs. In another workbook that index
        /// points at a different style, or past the end of the list, and then every save of the workbook fails.
        /// Drop it before EPPlus first reads the copied rules: <see cref="PreserveConditionalFormattingStyles"/>
        /// re-applies each style, and EPPlus writes it as a new dxf of this workbook when it is saved.
        /// </summary>
        private static void RemoveDxfIds(ExcelWorksheet copy)
        {
            var xml = copy.WorksheetXml;
            var ns = new System.Xml.XmlNamespaceManager(xml.NameTable);
            ns.AddNamespace("d", xml.DocumentElement.NamespaceURI);
            foreach (System.Xml.XmlElement rule in xml.SelectNodes("//d:conditionalFormatting/d:cfRule[@dxfId]", ns))
            {
                rule.RemoveAttribute("dxfId");
            }
        }

        /// <summary>
        /// Copying a worksheet into another workbook keeps each cell's display text but drops its
        /// hyperlink. Re-apply the hyperlinks from the source sheet's allocated cells onto the copy.
        /// </summary>
        internal static void PreserveHyperlinks(ExcelWorksheet source, ExcelWorksheet copy)
        {
            if (source == null || copy == null || source.Dimension == null)
            {
                return;
            }

            try
            {
                foreach (var cell in source.Cells[source.Dimension.Address])
                {
                    if (cell.Hyperlink != null && copy.Cells[cell.Address].Hyperlink == null)
                    {
                        copy.Cells[cell.Address].Hyperlink = cell.Hyperlink;
                    }
                }
            }
            catch
            {
                // never block a worksheet copy because of this
            }
        }

        private static bool RuleUsesDxfStyle(OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType type)
        {
            switch (type)
            {
                case OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.TwoColorScale:
                case OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.ThreeColorScale:
                case OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.ThreeIconSet:
                case OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.FourIconSet:
                case OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.FiveIconSet:
                case OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingRuleType.DataBar:
                    return false;
                default:
                    return true;
            }
        }

        private static readonly char[] ForbiddenSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

        /// <summary>
        /// Throw an ArgumentException when a worksheet name breaks Excel's rules: at most 31 characters,
        /// none of : \ / ? * [ ], and no apostrophe at the start or end. EPPlus accepts some of these
        /// (for example *), and Excel then can't open the file. Empty names are left to EPPlus.
        /// </summary>
        internal static void ValidateSheetName(string name, string paramName)
        {
            if (string.IsNullOrEmpty(name))
            {
                return;
            }
            if (name.Length > 31 || name.IndexOfAny(ForbiddenSheetNameChars) >= 0 || name[0] == '\'' || name[name.Length - 1] == '\'')
            {
                throw new ArgumentException($"'{name}' is not a valid worksheet name: it must be at most 31 characters, must not contain : \\ / ? * [ ] and must not start or end with an apostrophe.", paramName);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dimension"></param>
        /// <returns></returns>
        internal static RCDimensionRecord CastDimension(ExcelAddressBase dimension)
        {
            RCDimensionRecord dim = new RCDimensionRecord(null);

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
            RCAddressRecord add = new RCAddressRecord(null);

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
