using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutSystems.NssAdvanced_Excel;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            CssAdvanced_Excel excel = new CssAdvanced_Excel();

            excel.MssWorkbook_Create(1, "Test", new RLNewSheetRecordList(), out object workBook);

            excel.MssWorksheet_Select(workBook, 0, "Test", out object workSheet);

            excel.MssCell_Write(workSheet, "A4", 0, 0, "8000", "", new RCCellFormatRecord());

        }
    }
}
