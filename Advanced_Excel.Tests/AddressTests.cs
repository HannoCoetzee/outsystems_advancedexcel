using System.Globalization;
using System.Drawing;
using OfficeOpenXml;
using FluentAssertions;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for address conversion and color utility methods.
/// Mirrors MssAddress_From_RowCol, MssAddress_From_Text,
/// and MssUtil_ConvertHexCodeToRGB (Util.ConvertFromColorCode).
/// </summary>
public class AddressTests
{
    // ── MssAddress_From_RowCol ──────────────────────────────────────

    [Theory]
    [InlineData(1, 1, "A1")]
    [InlineData(1, 2, "B1")]
    [InlineData(1, 26, "Z1")]
    [InlineData(1, 27, "AA1")]
    [InlineData(1, 28, "AB1")]
    [InlineData(2, 1, "A2")]
    [InlineData(100, 10, "J100")]
    [InlineData(1, 52, "AZ1")]
    [InlineData(1, 53, "BA1")]
    [InlineData(1, 702, "ZZ1")]
    [InlineData(999, 1, "A999")]
    public void Address_FromRowCol_ConvertsCorrectly(int row, int col, string expected)
    {
        var addr = new ExcelAddressBase(row, col, row, col);
        addr.Address.Should().Be(expected);
    }

    // ── MssAddress_From_Text ────────────────────────────────────────

    [Theory]
    [InlineData("A1", 1, 1)]
    [InlineData("B2", 2, 2)]
    [InlineData("Z1", 1, 26)]
    [InlineData("AA1", 1, 27)]
    [InlineData("J100", 100, 10)]
    [InlineData("AB1", 1, 28)]
    public void Address_FromText_ParsesExpectedAddress(string address, int expectedRow, int expectedCol)
    {
        var addr = new ExcelCellAddress(address);
        addr.Row.Should().Be(expectedRow);
        addr.Column.Should().Be(expectedCol);
    }

    [Fact]
    public void Address_FromText_RoundTrip()
    {
        // Address "C5" → row 5, col 3 → reconstruct
        var addr = new ExcelCellAddress("C5");
        var reconstructed = new ExcelAddressBase(addr.Row, addr.Column, addr.Row, addr.Column);
        reconstructed.Address.Should().Be("C5");
    }

    // ── ConvertFromColorCode (MssUtil_ConvertHexCodeToRGB) ──────────

    [Theory]
    [InlineData("#FF0000", 255, 0, 0)]
    [InlineData("#00FF00", 0, 255, 0)]
    [InlineData("#0000FF", 0, 0, 255)]
    [InlineData("#FFFFFF", 255, 255, 255)]
    [InlineData("#000000", 0, 0, 0)]
    [InlineData("#ff00ff", 255, 0, 255)]
    public void ConvertFromColorCode_ValidHex_ReturnsExpectedRGB(string hex, int r, int g, int b)
    {
        // Mirrors Util.ConvertFromColorCode
        Color color = ColorFromHex(hex);
        color.R.Should().Be(r);
        color.G.Should().Be(g);
        color.B.Should().Be(b);
    }

    [Theory]
    [InlineData("FF0000")]
    [InlineData("")]
    [InlineData("not-a-color")]
    [InlineData("#GGGGGG")]
    public void ConvertFromColorCode_InvalidHex_ReturnsWhite(string invalid)
    {
        Color color = ColorFromHex(invalid);
        color.Should().Be(Color.White);
    }

    [Fact]
    public void ConvertFromColorCode_WithHash_ParsesCorrectly()
    {
        Color color = ColorFromHex("#1A2B3C");
        color.R.Should().Be(0x1A);
        color.G.Should().Be(0x2B);
        color.B.Should().Be(0x3C);
    }

    // ── Workbook/GetAllWorksheets cell value round-trip ─────────────

    [Fact]
    public void Cell_RoundTrip_WriteReadString()
    {
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("S1");
        ws.Cells[1, 1].Value = "RoundTrip";
        ws.Cells[1, 1].Text.Should().Be("RoundTrip");
    }

    [Fact]
    public void Cell_RoundTrip_WriteReadInteger()
    {
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("S1");
        ws.Cells[1, 1].Value = 42;
        Convert.ToInt32(ws.GetValue(1, 1)).Should().Be(42);
    }

    [Fact]
    public void Cell_RoundTrip_WriteReadDecimal()
    {
        using var pkg = new ExcelPackage();
        var ws = pkg.Workbook.Worksheets.Add("S1");
        ws.Cells[1, 1].Value = 3.14m;
        Convert.ToDecimal(ws.GetValue(1, 1)).Should().Be(3.14m);
    }

    // ── Helper ──────────────────────────────────────────────────────

    private static Color ColorFromHex(string colorCode)
    {
        try
        {
            return Color.FromArgb(int.Parse(colorCode.Replace("#", ""), NumberStyles.HexNumber));
        }
        catch
        {
            return Color.White;
        }
    }
}
