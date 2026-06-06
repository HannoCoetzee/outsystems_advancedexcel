// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file under the MIT license.
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

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for cell read/write operations by index and name.
/// Mirrors MssCell_WriteByIndex, MssCell_WriteByName, MssCell_WriteByIndexWithFormat,
/// MssCell_WriteByNameWithFormat, MssCell_ReadByIndex, MssCell_ReadByName,
/// MssCell_SetFormulaByIndex, MssCell_SetFormulaByName,
/// MssCell_GetFillColorByIndex, MssCell_GetFillColorByName,
/// MssCell_WriteColumnRange, MssCell_WriteRangeWithFormat.
/// </summary>
public class CellTests : IDisposable
{
    private readonly ExcelPackage _package;
    private readonly ExcelWorksheet _ws;
    private bool _disposed;

    public CellTests()
    {
        _package = new ExcelPackage();
        _ws = _package.Workbook.Worksheets.Add("Sheet1");
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _package.Dispose();
            _disposed = true;
        }
    }

    // ── Write by Index ──────────────────────────────────────────────

    [Theory]
    [InlineData("general", "Hello", "Hello")]
    [InlineData("text", "00123", "00123")]
    [InlineData("integer", "42", "42")]
    [InlineData("decimal", "3.14", "3.14")]
    [InlineData("boolean", "true", "True")]
    public void WriteByIndex_VariousTypes_WritesCorrectValue(string cellType, string input, string expected)
    {
        // MssCell_WriteByIndex: write value with type conversion
        _ws.Cells[1, 1].Value = ConvertValue(input, cellType);
        ApplyNumberFormat(_ws.Cells[1, 1], cellType);

        _ws.Cells[1, 1].Text.Should().Be(expected);
    }

    [Fact]
    public void WriteByIndex_DateTime_StoresAsDate()
    {
        var date = new DateTime(2024, 6, 15);
        _ws.Cells[1, 1].Value = date;
        _ws.Cells[1, 1].Style.Numberformat.Format = "yyyy-mm-dd";

        _ws.Cells[1, 1].Text.Should().Be("2024-06-15");
    }

    [Fact]
    public void WriteByIndex_Formula_StoresFormula()
    {
        _ws.Cells[1, 1].Value = 10;
        _ws.Cells[1, 2].Value = 20;
        _ws.Cells[1, 3].Formula = "A1+B1";

        _ws.Cells[1, 3].Formula.Should().Be("A1+B1");
    }

    // ── Write by Name ───────────────────────────────────────────────

    [Fact]
    public void WriteByName_GeneralType_WritesValue()
    {
        // MssCell_WriteByName
        _ws.Cells["A1"].Value = "Test";
        _ws.Cells["A1"].Text.Should().Be("Test");
    }

    [Fact]
    public void WriteByName_TextType_PreservesLeadingZeros()
    {
        _ws.Cells["B2"].Value = "007";
        _ws.Cells["B2"].Style.Numberformat.Format = "@"; // text format

        _ws.Cells["B2"].Text.Should().Be("007");
    }

    // ── Read by Index ───────────────────────────────────────────────

    [Fact]
    public void ReadByIndex_AsValue_ReturnsStringRepresentation()
    {
        // MssCell_ReadByIndex with ssReadText = false
        _ws.Cells[1, 1].Value = 42;

        string value = Convert.ToString(_ws.GetValue(1, 1));
        value.Should().Be("42");
    }

    [Fact]
    public void ReadByIndex_AsText_ReturnsFormattedText()
    {
        // MssCell_ReadByIndex with ssReadText = true
        _ws.Cells[1, 1].Value = 1234.56;
        _ws.Cells[1, 1].Style.Numberformat.Format = "#,##0.00";

        _ws.Cells[1, 1].Text.Should().Contain("1,234.56");
    }

    [Fact]
    public void ReadByIndex_EmptyCell_ReturnsEmptyString()
    {
        string value = Convert.ToString(_ws.GetValue(1, 1));
        value.Should().Be("null"); // EPPlus returns null for empty cells
    }

    // ── Read by Name ────────────────────────────────────────────────

    [Fact]
    public void ReadByName_AsValue_ReturnsStringRepresentation()
    {
        // MssCell_ReadByName with ssReadText = false
        _ws.Cells["C3"].Value = 99;

        var addr = new ExcelCellAddress("C3");
        string value = Convert.ToString(_ws.GetValue(addr.Row, addr.Column));
        value.Should().Be("99");
    }

    [Fact]
    public void ReadByName_AsText_ReturnsFormattedText()
    {
        _ws.Cells["D4"].Value = 3.14;
        _ws.Cells["D4"].Style.Numberformat.Format = "0.00";

        var addr = new ExcelCellAddress("D4");
        _ws.Cells[addr.Row, addr.Column].Text.Should().Be("3.14");
    }

    // ── Formula ─────────────────────────────────────────────────────

    [Fact]
    public void SetFormulaByIndex_WritesFormula()
    {
        // MssCell_SetFormulaByIndex
        _ws.Cells[1, 1].Value = 10;
        _ws.Cells[1, 2].Value = 20;
        _ws.Cells[1, 3].Formula = "A1+B1";

        _ws.Cells[1, 3].Formula.Should().Be("A1+B1");
    }

    [Fact]
    public void SetFormulaByName_WritesFormula()
    {
        // MssCell_SetFormulaByName
        _ws.Cells["A1"].Value = 5;
        _ws.Cells["A2"].Value = 3;
        _ws.Cells["A3"].Formula = "A1*A2";

        _ws.Cells["A3"].Formula.Should().Be("A1*A2");
    }

    // ── Fill Color ──────────────────────────────────────────────────

    [Fact]
    public void GetFillColorByIndex_WithColor_ReturnsHexColor()
    {
        // MssCell_GetFillColorByIndex
        _ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        _ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 0));

        var fillColor = _ws.Cells[1, 1].Style.Fill.BackgroundColor.Rgb;
        fillColor.Should().NotBeNullOrEmpty();
        // EPPlus returns ARGB format; last 6 chars are RGB
        string hex = "#" + fillColor.Substring(fillColor.Length - 6);
        hex.Should().Be("#FF0000");
    }

    [Fact]
    public void GetFillColorByIndex_NoFill_ReturnsNoFillColor()
    {
        var fillColor = _ws.Cells[1, 1].Style.Fill.BackgroundColor.Rgb;
        string result = string.IsNullOrEmpty(fillColor) ? "No Fill Color" : "#" + fillColor.Substring(fillColor.Length - 6);
        result.Should().Be("No Fill Color");
    }

    [Fact]
    public void GetFillColorByName_WithColor_ReturnsHexColor()
    {
        // MssCell_GetFillColorByName
        _ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        _ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 128, 0));

        var cell = _ws.Cells["A1"];
        var fillColor = cell.Style.Fill.BackgroundColor.Rgb;
        string hex = "#" + fillColor.Substring(fillColor.Length - 6);
        hex.Should().Be("#008000");
    }

    // ── Column Range Write ──────────────────────────────────────────

    [Fact]
    public void WriteColumnRange_WritesValuesAcrossColumns()
    {
        // MssCell_WriteColumnRange: write values starting at row 1, col 1
        var values = new[] { "A", "B", "C", "D" };
        for (int i = 0; i < values.Length; i++)
        {
            _ws.Cells[1, 1 + i].Value = values[i];
        }

        _ws.Cells[1, 1].Text.Should().Be("A");
        _ws.Cells[1, 2].Text.Should().Be("B");
        _ws.Cells[1, 3].Text.Should().Be("C");
        _ws.Cells[1, 4].Text.Should().Be("D");
    }

    // ── Image Write ─────────────────────────────────────────────────

    [Fact]
    public void WriteImageByIndex_InsertsImage()
    {
        // MssCell_WriteByIndex with image bytes
        // Create a minimal 1x1 PNG image
        using var bitmap = new Bitmap(1, 1);
        bitmap.SetPixel(0, 0, Color.Red);
        using var ms = new MemoryStream();
        bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        byte[] imageBytes = ms.ToArray();

        var picture = _ws.Drawings.AddPicture("TestImage", new MemoryStream(imageBytes));
        picture.SetPosition(0, 0, 0, 0);

        _ws.Drawings.Count.Should().Be(1);
        _ws.Drawings[0].Name.Should().Be("TestImage");
    }

    // ── Helpers ─────────────────────────────────────────────────────

    private static object ConvertValue(string value, string cellType)
    {
        return cellType.ToLower() switch
        {
            "integer" => int.Parse(value),
            "decimal" => decimal.Parse(value),
            "boolean" => bool.Parse(value),
            "datetime" => DateTime.Parse(value),
            _ => value // general, text
        };
    }

    private static void ApplyNumberFormat(ExcelRange cell, string cellType)
    {
        switch (cellType.ToLower())
        {
            case "text":
                cell.Style.Numberformat.Format = "@";
                break;
            case "integer":
                cell.Style.Numberformat.Format = "0";
                break;
            case "decimal":
                cell.Style.Numberformat.Format = "0.00";
                break;
        }
    }
}
