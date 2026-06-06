// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file under the MIT license.
// See LICENSE file in the project root for full information.
//
// EPPlus is licensed under the Polyform Noncommercial 1.0.0 license.
// For commercial use, a commercial license must be obtained from EPPlus Software.
//
// This test project is provided as a scaffold for unit testing the
// OutSystems Advanced_Excel component. It is not affiliated with or
// endorsed by OutSystems.

using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using FluentAssertions;
using System.Globalization;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for workbook and worksheet creation, selection, deletion, and copying.
/// Mirrors MssWorkbook_Create, MssWorkbook_Open_BinaryData, MssWorksheet_SelectByIndex,
/// MssWorksheet_SelectByName, MssWorksheet_DeleteByIndex, MssWorksheet_DeleteByName,
/// MssWorkbook_AddCopyWorksheet, MssWorksheet_SetActive, MssWorksheet_CopyRows.
/// </summary>
public class WorkbookTests : IDisposable
{
    private readonly ExcelPackage _package;
    private bool _disposed;

    public WorkbookTests()
    {
        _package = new ExcelPackage();
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _package.Dispose();
            _disposed = true;
        }
    }

    [Fact]
    public void Workbook_Create_Default_SingleSheet()
    {
        // MssWorkbook_Create with default params → single "Sheet1"
        var wb = _package.Workbook;
        wb.Worksheets.Add("Sheet1");

        wb.Worksheets.Count.Should().Be(1);
        wb.Worksheets[0].Name.Should().Be("Sheet1");
    }

    [Fact]
    public void Workbook_Create_MultipleSheets()
    {
        // MssWorkbook_Create with ssNumberOfSheets = 3, ssFirstSheetName = "Data"
        var wb = _package.Workbook;
        wb.Worksheets.Add("Data");
        wb.Worksheets.Add("Data2");
        wb.Worksheets.Add("Data3");

        wb.Worksheets.Count.Should().Be(3);
        wb.Worksheets[0].Name.Should().Be("Data");
        wb.Worksheets[1].Name.Should().Be("Data2");
        wb.Worksheets[2].Name.Should().Be("Data3");
    }

    [Fact]
    public void Workbook_Create_CustomFirstSheetName()
    {
        var wb = _package.Workbook;
        wb.Worksheets.Add("Report");

        wb.Worksheets.Count.Should().Be(1);
        wb.Worksheets[0].Name.Should().Be("Report");
    }

    [Fact]
    public void Worksheet_SelectByIndex_ReturnsCorrectWorksheet()
    {
        // MssWorksheet_SelectByIndex
        _package.Workbook.Worksheets.Add("Sheet1");
        _package.Workbook.Worksheets.Add("Sheet2");
        _package.Workbook.Worksheets.Add("Sheet3");

        var ws = _package.Workbook.Worksheets[1]; // 0-based index
        ws.Name.Should().Be("Sheet2");
    }

    [Fact]
    public void Worksheet_SelectByName_ReturnsCorrectWorksheet()
    {
        // MssWorksheet_SelectByName
        _package.Workbook.Worksheets.Add("Alpha");
        _package.Workbook.Worksheets.Add("Beta");

        var ws = _package.Workbook.Worksheets["Beta"];
        ws.Should().NotBeNull();
        ws!.Name.Should().Be("Beta");
    }

    [Fact]
    public void Worksheet_DeleteByName_RemovesSheet()
    {
        // MssWorksheet_DeleteByName
        _package.Workbook.Worksheets.Add("Keep");
        _package.Workbook.Worksheets.Add("Remove");
        _package.Workbook.Worksheets.Count.Should().Be(2);

        _package.Workbook.Worksheets.Delete("Remove");
        _package.Workbook.Worksheets.Count.Should().Be(1);
        _package.Workbook.Worksheets[0].Name.Should().Be("Keep");
    }

    [Fact]
    public void Worksheet_DeleteByIndex_RemovesSheet()
    {
        // MssWorksheet_DeleteByIndex
        _package.Workbook.Worksheets.Add("First");
        _package.Workbook.Worksheets.Add("Second");
        _package.Workbook.Worksheets.Count.Should().Be(2);

        _package.Workbook.Worksheets.Delete(0); // Delete first
        _package.Workbook.Worksheets.Count.Should().Be(1);
        _package.Workbook.Worksheets[0].Name.Should().Be("Second");
    }

    [Fact]
    public void Worksheet_CopyRows_CopiesCellValues()
    {
        // MssWorksheet_CopyRows: copy A1:B2 to A4:B5
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Hello";
        ws.Cells["B1"].Value = "World";
        ws.Cells["A2"].Value = 42;
        ws.Cells["B2"].Value = 3.14;

        ws.Cells["A1:B2"].Copy(ws.Cells["A4:B5"]);

        ws.Cells["A4"].Text.Should().Be("Hello");
        ws.Cells["B4"].Text.Should().Be("World");
        ws.Cells["A5"].Text.Should().Be("42");
        ws.Cells["B5"].Text.Should().Be("3.14");
    }

    [Fact]
    public void Worksheet_SetActive_ByName()
    {
        // MssWorksheet_SetActive by name
        _package.Workbook.Worksheets.Add("Sheet1");
        _package.Workbook.Worksheets.Add("Sheet2");

        _package.Workbook.View.ActiveTab = _package.Workbook.Worksheets["Sheet2"].Index;
        _package.Workbook.View.ActiveTab.Should().Be(1);
    }

    [Fact]
    public void Worksheet_SetActive_ByIndex()
    {
        // MssWorksheet_SetActive by index
        _package.Workbook.Worksheets.Add("Sheet1");
        _package.Workbook.Worksheets.Add("Sheet2");
        _package.Workbook.Worksheets.Add("Sheet3");

        _package.Workbook.View.ActiveTab = 2; // Third sheet (0-based)
        _package.Workbook.View.ActiveTab.Should().Be(2);
    }

    [Fact]
    public void Workbook_AddCopyWorksheet_CopiesContent()
    {
        // MssWorkbook_AddCopyWorksheet
        var original = _package.Workbook.Worksheets.Add("Original");
        original.Cells["A1"].Value = "Test Data";
        original.Cells["B2"].Value = 123;

        var copy = _package.Workbook.Worksheets.Add("Copy", original);

        copy.Cells["A1"].Text.Should().Be("Test Data");
        copy.Cells["B2"].Text.Should().Be("123");
    }

    [Fact]
    public void Workbook_Open_FromBinaryData()
    {
        // MssWorkbook_Open_BinaryData: create package, save to bytes, re-open
        var ws = _package.Workbook.Worksheets.Add("Data");
        ws.Cells["A1"].Value = "RoundTrip";
        byte[] data = _package.GetAsByteArray();

        using var reopened = new ExcelPackage(new MemoryStream(data));
        reopened.Workbook.Worksheets[0].Name.Should().Be("Data");
        reopened.Workbook.Worksheets[0].Cells["A1"].Text.Should().Be("RoundTrip");
    }

    [Fact]
    public void Workbook_AddName_CreatesNamedRange()
    {
        // MssWorkbook_AddName
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = 100;
        ws.Cells["A2"].Value = 200;

        ws.Workbook.Names.Add("MyRange", ws.Cells["A1:A2"]);
        ws.Workbook.Names["MyRange"].Should().NotBeNull();
    }

    [Fact]
    public void Worksheet_AddName_CreatesNamedRangeOnWorksheet()
    {
        // MssWorksheet_AddName
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["B1"].Value = "Named";
        ws.Cells["B2"].Value = "Range";

        ws.Names.Add("LocalName", ws.Cells["B1:B2"]);
        ws.Names["LocalName"].Should().NotBeNull();
    }
}
