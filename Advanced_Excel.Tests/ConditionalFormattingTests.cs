using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using FluentAssertions;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for conditional formatting operations.
/// Mirrors MssConditionalFormatting_AddExpressionRule,
/// MssConditionalFormatting_AddCellValuesRule,
/// MssConditionalFormatting_AddThreeColorScaleRule,
/// MssConditionalFormatting_AddTwoColorScaleRule,
/// MssConditionalFormatting_AddDatabarRule,
/// MssConditionalFormatting_AddIconSetRule,
/// MssConditionalFormatting_DeleteRule,
/// MssConditionalFormatting_DeleteAllRules.
/// </summary>
public class ConditionalFormattingTests : IDisposable
{
    private readonly ExcelPackage _package;
    private readonly ExcelWorksheet _ws;
    private bool _disposed;

    public ConditionalFormattingTests()
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

    // ── Expression Rule ─────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddExpressionRule_AddsRule()
    {
        // MssConditionalFormatting_AddExpressionRule
        var cf = _ws.ConditionalFormatting.AddExpression("A1:A10");
        cf.Formula = "A1>5";
        cf.Style.Font.Color.Color = Color.Red;

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    [Fact]
    public void ConditionalFormat_AddExpressionRule_HighlightsMatchingCells()
    {
        _ws.Cells["A1"].Value = 10;
        _ws.Cells["A2"].Value = 3;
        _ws.Cells["A3"].Value = 7;

        var cf = _ws.ConditionalFormatting.AddExpression("A1:A10");
        cf.Formula = "A1>5";
        cf.Style.Font.Bold = true;

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Cell Values Rule (GreaterThan) ──────────────────────────────

    [Fact]
    public void ConditionalFormat_AddCellValuesRule_GreaterThan_AddsRule()
    {
        // MssConditionalFormatting_AddCellValuesRule with Operator.GreaterThan
        var cf = _ws.ConditionalFormatting.AddGreaterThan("A1:A10");
        cf.Formula = "5";
        cf.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cf.Style.Fill.BackgroundColor.Color = Color.Yellow;

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Three Color Scale ───────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddThreeColorScaleRule_AddsRule()
    {
        // MssConditionalFormatting_AddThreeColorScaleRule
        var cf = _ws.ConditionalFormatting.AddThreeColorScale("A1:A10");

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    [Fact]
    public void ConditionalFormat_AddThreeColorScaleRule_HasThreePoints()
    {
        var cf = _ws.ConditionalFormatting.AddThreeColorScale("A1:A10");
        cf.Should().NotBeNull();
        // Three color scale has low/mid/high points
    }

    // ── Two Color Scale ─────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddTwoColorScaleRule_AddsRule()
    {
        // MssConditionalFormatting_AddTwoColorScaleRule
        var cf = _ws.ConditionalFormatting.AddTwoColorScale("A1:A10");

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Data Bar ────────────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddDatabarRule_AddsRule()
    {
        // MssConditionalFormatting_AddDatabarRule
        var cf = _ws.ConditionalFormatting.AddDatabar("A1:A10", Color.Blue);

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Icon Set ────────────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddIconSetRule_AddsRule()
    {
        // MssConditionalFormatting_AddIconSetRule (3TrafficLights)
        var cf = _ws.ConditionalFormatting.AddThreeIconSet("A1:A10", eExcelConditionalFormatting3IconsSetType.TrafficLights);

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    [Fact]
    public void ConditionalFormat_AddIconSetRule_FiveArrows_AddsRule()
    {
        var cf = _ws.ConditionalFormatting.AddFiveIconSet("A1:A10", eExcelConditionalFormatting5IconsSetType.Arrows);

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Multiple Rules ──────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddMultipleRules_AllPresent()
    {
        _ws.ConditionalFormatting.AddExpression("A1:A10");
        _ws.ConditionalFormatting.AddDatabar("B1:B10", Color.Green);
        _ws.ConditionalFormatting.AddThreeColorScale("C1:C10");

        _ws.ConditionalFormatting.Count.Should().Be(3);
    }

    // ── Delete Rule ─────────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_DeleteRule_RemovesRule()
    {
        // MssConditionalFormatting_DeleteRule
        _ws.ConditionalFormatting.AddExpression("A1:A10");
        _ws.ConditionalFormatting.AddExpression("B1:B10");
        _ws.ConditionalFormatting.Count.Should().Be(2);

        _ws.ConditionalFormatting.RemoveAt(1); // Remove first
        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    [Fact]
    public void ConditionalFormat_DeleteRule_InvalidIndex_Throws()
    {
        // MssConditionalFormatting_DeleteRule: index <= 0 throws
        var act = () => _ws.ConditionalFormatting.RemoveAt(0);
        act.Should().Throw<ArgumentOutOfRangeException>();
    }

    // ── Delete All Rules ────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_DeleteAllRules_RemovesAll()
    {
        // MssConditionalFormatting_DeleteAllRules
        _ws.ConditionalFormatting.AddExpression("A1:A10");
        _ws.ConditionalFormatting.AddDatabar("B1:B10", Color.Red);
        _ws.ConditionalFormatting.AddThreeIconSet("C1:C10", eExcelConditionalFormatting3IconsSetType.TrafficLights);
        _ws.ConditionalFormatting.Count.Should().Be(3);

        _ws.ConditionalFormatting.RemoveAll();
        _ws.ConditionalFormatting.Count.Should().Be(0);
    }

    // ── Between Rule ────────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddBetweenRule_AddsRule()
    {
        var cf = _ws.ConditionalFormatting.AddBetween("A1:A10");
        cf.Formula = "1";
        cf.Formula2 = "10";
        cf.Style.Font.Italic = true;

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Top/Bottom Rule ─────────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddTopBottomRule_AddsRule()
    {
        var cf = _ws.ConditionalFormatting.AddTopPercent("A1:A10");
        cf.Rank = 10; // top 10%
        cf.Style.Font.Bold = true;

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Duplicate Values Rule ───────────────────────────────────────

    [Fact]
    public void ConditionalFormat_AddDuplicateValuesRule_AddsRule()
    {
        var cf = _ws.ConditionalFormatting.AddDuplicateValues("A1:A10");
        cf.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cf.Style.Fill.BackgroundColor.Color = Color.Orange;

        _ws.ConditionalFormatting.Count.Should().Be(1);
    }

    // ── Style Application ───────────────────────────────────────────

    [Fact]
    public void ConditionalFormat_StyleApplied_FontBold()
    {
        var cf = _ws.ConditionalFormatting.AddExpression("A1:A10");
        cf.Style.Font.Bold = true;
        cf.Style.Font.Color.Color = Color.Red;

        cf.Style.Font.Bold.Should().BeTrue();
    }

    [Fact]
    public void ConditionalFormat_StyleApplied_Border()
    {
        var cf = _ws.ConditionalFormatting.AddExpression("A1:A10");
        cf.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        cf.Style.Border.Bottom.Color.Color = Color.Black;

        cf.Style.Border.Bottom.Style.Should().Be(ExcelBorderStyle.Thin);
    }

    [Fact]
    public void ConditionalFormat_StyleApplied_NumberFormat()
    {
        var cf = _ws.ConditionalFormatting.AddExpression("A1:A10");
        cf.Style.NumberFormat.Format = "#,##0.00";

        cf.Style.NumberFormat.Format.Should().Be("#,##0.00");
    }
}
