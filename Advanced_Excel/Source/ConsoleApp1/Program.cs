using System;
using System.CodeDom;
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
        // To read the file as well, pass in the name of an existing Excel file,
        // otherwise it will just be generated and saved to Test.xlsx
        static void Main(string[] args)
        {
            CssAdvanced_Excel excel = new CssAdvanced_Excel();

            excel.MssAddress_From_RowCol(3, 3, 3, 6, out string address);
            Console.WriteLine("3,3,3,6 -> {0}", address);
            excel.MssAddress_From_RowCol(3, 3, 0, -13, out address);
            Console.WriteLine("3,3,0,-13 -> {0}", address);

            address = "AB47";
            excel.MssAddress_From_Text(address, out int rStart, out int cStart, out int rEnd, out int cEnd);
            Console.WriteLine("{0} -> {1} {2} {3} {4}", address, rStart, cStart, rEnd, cEnd);

            var path = "Test.xlsx";

            var exists = false;

            if (args.Length == 1)
            {
                path = args[0];
                exists = File.Exists(path);
            }


            object workBook;

            if (exists)
            {
                Console.WriteLine($"Opening Excel {path}...");
                var existingContent = File.ReadAllBytes(path);
                excel.MssWorkbook_Open_BinaryData(existingContent, out workBook);
            }
            else
            {
                Console.WriteLine($"Generating Excel...");
                excel.MssWorkbook_Create(1, "Test", new RLNewSheetRecordList(), out workBook);
            }

            excel.MssWorksheet_Select(workBook, 0, "Test", out object workSheet);

            // Write a cell
            excel.MssCell_Write(workSheet, "A4", 0, 0, "8000", "", new RCCellFormatRecord());

            RCConditionalFormatItemRecord r1 = new RCConditionalFormatItemRecord
            {
                ssSTConditionalFormatItem = new STConditionalFormatItemStructure
                {
                    ssRuleType = 20
                }
            };
            RCAddressRecord addressRecord = new RCAddressRecord
            {
                ssSTAddress = new STAddressStructure()
            };
            addressRecord.ssSTAddress.ssColumn = 1;
            addressRecord.ssSTAddress.ssRow = 1;
            r1.ssSTConditionalFormatItem.ssAddress = addressRecord;
            excel.MssConditionalFormatting_AddRule(workSheet, r1);

            // Header and Footer
            if (exists)
            {
                excel.MssWorksheet_GetHeader(workSheet, false,
                    out string leftSection, out string centerSection, out string rightSection);
                Console.WriteLine($"Existing header left: {leftSection}, center: {centerSection}, right: {rightSection}");
                excel.MssWorksheet_GetFooter(workSheet, false,
                    out leftSection, out centerSection, out rightSection);
                Console.WriteLine($"Existing footer left: {leftSection}, center: {centerSection}, right: {rightSection}");
            }
            excel.MssWorksheet_SetHeader(workSheet, "&KFF0000The left header section (red)", "Sheet name: &A", "The date &D and time &T");

            excel.MssWorksheet_SetFooter(workSheet, "The filename &F", "The center footer section", @"Page &P of &N");

            excel.MssWorkbook_GetBinaryData(workBook, out byte[] content);

            Console.WriteLine($"Writing Excel file to {path}...");
            File.WriteAllBytes(path, content);
            Console.WriteLine($"Done.");
        }
    }
}
