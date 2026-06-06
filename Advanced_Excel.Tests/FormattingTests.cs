using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using FluentAssertions;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for cell range formatting operations.
/// Mirrors MssCell_FormatRange and Util.ApplyFormatToRange.
/// </summary>
public class FormattingTests : IDisposable
{
    private readonly ExcelPackage _package;
    private readonly ExcelWorksheet _ws;
    private bool _disposed;

    public FormattingTests()
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

    private void ApplyRangeFormat(ExcelRange range, Action<ExcelRangeStyle> configure)
    {
        configure(range.Style);
    }

    // ── Font Name ───────────────────────────────────────────────────

    [Fact]
    public void FormatRange_FontName_AppliedToRange()
    {
        _ws.Cells["A1:C3"].Style.Font.Name = "Consolas";
        _ws.Cells["A1"].Style.Font.Name.Should().Be("Consolas");
        _ws.Cells["C3"].Style.Font.Name.Should().Be("Consolas");
    }

    // ── Font Size ───────────────────────────────────────────────────

    [Fact]
    public void FormatRange_FontSize_AppliedToRange()
    {
        _ws.Cells["A1:B2"].Style.Font.Size = 14;
        _ws.Cells["A1"].Style.Font.Size.Should().Be(14);
        _ws.Cells["B2"].Style.Font.Size.Should().Be(14);
    }

    // ── Background Color ────────────────────────────────────────────

    [Fact]
    public void FormatRange_BackgroundColor_AppliedToRange()
    {
        Color expected = ColorTranslator.FromHtml("#FFCC00");
        var style = _ws.Cells["A1:C3"].Style;
        style.Fill.PatternType = ExcelFillStyle.Solid;
        style.Fill.BackgroundColor.SetColor(expected);

        _ws.Cells["A1"].Style.Fill.PatternType.Should().Be(ExcelFillStyle.Solid);
        _ws.Cells["B2"].Style.Fill.BackgroundColor.Rgb.Should().Contain("FFCC00");
    }

    // ── Font Color ──────────────────────────────────────────────────

    [Fact]
    public void FormatRange_FontColor_AppliedToRange()
    {
        Color expected = Color.Red;
        _ws.Cells["A1:C1"].Style.Font.Color.SetColor(expected);

        _ws.Cells["A1"].Style.Font.Color.Rgb.Should().Contain("FF0000");
    }

    // ── Bold ────────────────────────────────────────────────────────

    [Fact]
    public void FormatRange_Bold_AppliedToRange()
    {
        _ws.Cells["A1:A5"].Style.Font.Bold = true;
        foreach (var cell in _ws.Cells["A1:A5"])
            cell.Style.Font.Bold.Should().BeTrue();
    }

    // ── Border Around ───────────────────────────────────────────────

    [Fact]
    public void FormatRange_BorderAround_AllCellBordersSet()
    {
        _ws.Cells["A1:C3"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

        _ws.Cells["B2"].Style.Border.Top.Style.Should().Be(ExcelBorderStyle.Thin);
    }

    // ── Individual Borders ──────────────────────────────────────────

    [Theory]
    [InlineData("Bottom", ExcelBorderStyle.Medium)]
    [InlineData("Top", ExcelBorderStyle.Thick)]
    [InlineData("Left", ExcelBorderStyle.Hair)]
    [InlineData("Right", ExcelBorderStyle.Dotted)]
    public void FormatRange_BorderIndividual_AppliedCorrectly(string borderSide, ExcelBorderStyle borderStyle)
    {
        var range = _ws.Cells["A1:B2"];
        switch (borderSide)
        {
            case "Bottom": range.Style.Border.Bottom.Style = borderStyle; break;
            case "Top": range.Style.Border.Top.Style = borderStyle; break;
            case "Left": range.Style.Border.Left.Style = borderStyle; break;
            case "Right": range.Style.Border.Right.Style = borderStyle; break;
        }

        var actual = borderSide switch
        {
            "Bottom" => _ws.Cells["A1"].Style.Border.Bottom.Style,
            "Top" => _ws.Cells["A1"].Style.Border.Top.Style,
            "Left" => _ws.Cells["A1"].Style.Border.Left.Style,
            "Right" => _ws.Cells["A1"].Style.Border.Right.Style,
            _ => throw new InvalidOperationException()
        };
        actual.Should().Be(borderStyle);
    }

    // ── Number Format ───────────────────────────────────────────────

    [Fact]
    public void FormatRange_NumberFormat_AppliedToRange()
    {
        _ws.Cells["A1:A3"].Style.Numberformat.Format = "#,##0.00";
        _ws.Cells["A1"].Style.Numberformat.Format.Should().Be("#,##0.00");
    }

    [Fact]
    public void FormatRange_DateFormat_FormatsDate()
    {
        _ws.Cells["A1"].Value = new DateTime(2024, 6, 15);
        _ws.Cells["A1"].Style.Numberformat.Format = "yyyy-mm-dd";
        _ws.Cells["A1"].Text.Should().Be("2024-06-15");
    }

    // ── Horizontal Alignment ────────────────────────────────────────

    [Theory]
    [InlineData(ExcelHorizontalAlignment.Left)]
    [InlineData(ExcelHorizontalAlignment.Center)]
    [InlineData(ExcelHorizontalAlignment.Right)]
    [InlineData(ExcelHorizontalAlignment.CenterContinuous)]
    public void FormatRange_HorizontalAlignment_Applied(ExcelHorizontalAlignment alignment)
    {
        _ws.Cells["A1:B2"].Style.HorizontalAlignment = alignment;
        _ws.Cells["A1"].Style.HorizontalAlignment.Should().Be(alignment);
    }

    // ── Vertical Alignment ──────────────────────────────────────────

    [Theory]
    [InlineData(ExcelVerticalAlignment.Top)]
    [InlineData(ExcelVerticalAlignment.Center)]
    [InlineData(ExcelVerticalAlignment.Bottom)]
    public void FormatRange_VerticalAlignment_Applied(ExcelVerticalAlignment alignment)
    {
        _ws.Cells["A1:B2"].Style.VerticalAlignment = alignment;
        _ws.Cells["A1"].Style.VerticalAlignment.Should().Be(alignment);
    }

    // ── Wrap Text ───────────────────────────────────────────────────

    [Fact]
    public void FormatRange_WrapText_Applied()
    {
        _ws.Cells["A1:A3"].Style.WrapText = true;
        _ws.Cells["A2"].Style.WrapText.Should().BeTrue();
    }

    // ── Text Rotation ───────────────────────────────────────────────

    [Fact]
    public void FormatRange_TextRotation_Applied()
    {
        _ws.Cells["A1"].Style.TextRotation = 45;
        _ws.Cells["A1"].Style.TextRotation.Should().Be(45);
    }

    // ── Shrink to Fit ───────────────────────────────────────────────

    [Fact]
    public void FormatRange_ShrinkToFit_Applied()
    {
        _ws.Cells["A1"].Style.ShrinkToFit = true;
        _ws.Cells["A1"].Style.ShrinkToFit.Should().BeTrue();
    }

    // ── Locked ──────────────────────────────────────────────────────

    [Fact]
    public void FormatRange_Locked_Applied()
    {
        _ws.Cells["A1:B2"].Style.Locked = true;
        _ws.Cells["A1"].Style.Locked.Should().BeTrue();
    }

    // ── Indent ──────────────────────────────────────────────────────

    [Fact]
    public void FormatRange_Indent_Applied()
    {
        _ws.Cells["A1"].Style.Indent = 3;
        _ws.Cells["A1"].Style.Indent.Should().Be(3);
    }

    // ── Quote Prefix ────────────────────────────────────────────────

    [Fact]
    public void FormatRange_QuotePrefix_Applied()
    {
        _ws.Cells["A1"].Style.QuotePrefix = true;
        _ws.Cells["A1"].Style.QuotePrefix.Should().BeTrue();
    }

    // ── Autofit Columns ─────────────────────────────────────────────

    [Fact]
    public void FormatRange_AutofitColumns_DoesNotThrow()
    {
        _ws.Cells["A1"].Value = "Some content";
        var act = () => _ws.Cells["A1"].AutoFitColumns();
        act.Should().NotThrow();
    }

    // ── Combined Format ─────────────────────────────────────────────

    [Fact]
    public void FormatRange_CombinedFormat_AllApplied()
    {
        var range = _ws.Cells["A1:C3"];
        range.Style.Font.Name = "Arial";
        range.Style.Font.Size = 12;
        range.Style.Font.Bold = true;
        range.Style.Font.Color.SetColor(Color.Blue);
        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
        range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        range.Style.Numberformat.Format = "#,##0";

        _ws.Cells["B2"].Style.Font.Name.Should().Be("Arial");
        _ws.Cells["B2"].Style.Font.Size.Should().Be(12);
        _ws.Cells["B2"].Style.Font.Bold.Should().BeTrue();
        _ws.Cells["B2"].Style.HorizontalAlignment.Should().Be(ExcelHorizontalAlignment.Center);
    }
}
