using OfficeOpenXml;
using FluentAssertions;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for column and row hide/show operations.
/// Mirrors MssColumn_HideByIndex, MssColumn_ShowByIndex, MssColumn_HideByName,
/// MssColumn_ShowByName, MssRow_HideByIndex, MssRow_ShowByIndex.
/// </summary>
public class ColumnRowTests : IDisposable
{
    private readonly ExcelPackage _package;
    private readonly ExcelWorksheet _ws;
    private bool _disposed;

    public ColumnRowTests()
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

    // ── Column Hide/Show ────────────────────────────────────────────

    [Fact]
    public void Column_HideByIndex_HidesColumn()
    {
        // MssColumn_HideByIndex
        _ws.Column(3).Hidden = true;
        _ws.Column(3).Hidden.Should().BeTrue();
    }

    [Fact]
    public void Column_ShowByIndex_UnhidesColumn()
    {
        // MssColumn_ShowByIndex
        _ws.Column(3).Hidden = true;
        _ws.Column(3).Hidden.Should().BeTrue();

        _ws.Column(3).Hidden = false;
        _ws.Column(3).Hidden.Should().BeFalse();
    }

    [Fact]
    public void Column_HideByName_HidesColumn()
    {
        // MssColumn_HideByName: "B" → column 2
        _ws.Column(2).Hidden = true;
        _ws.Column(2).Hidden.Should().BeTrue();
    }

    [Fact]
    public void Column_ShowByName_UnhidesColumn()
    {
        // MssColumn_ShowByName
        _ws.Column(2).Hidden = true;
        _ws.Column(2).Hidden = false;
        _ws.Column(2).Hidden.Should().BeFalse();
    }

    [Fact]
    public void Column_HideMultiple_HidesCorrectColumns()
    {
        _ws.Column(1).Hidden = true;
        _ws.Column(3).Hidden = true;
        _ws.Column(5).Hidden = true;

        _ws.Column(1).Hidden.Should().BeTrue();
        _ws.Column(2).Hidden.Should().BeFalse();
        _ws.Column(3).Hidden.Should().BeTrue();
        _ws.Column(4).Hidden.Should().BeFalse();
        _ws.Column(5).Hidden.Should().BeTrue();
    }

    // ── Row Hide/Show ──────────────────────────────────────────────

    [Fact]
    public void Row_HideByIndex_HidesRow()
    {
        // MssRow_HideByIndex
        _ws.Row(3).Hidden = true;
        _ws.Row(3).Hidden.Should().BeTrue();
    }

    [Fact]
    public void Row_ShowByIndex_UnhidesRow()
    {
        // MssRow_ShowByIndex
        _ws.Row(3).Hidden = true;
        _ws.Row(3).Hidden.Should().BeTrue();

        _ws.Row(3).Hidden = false;
        _ws.Row(3).Hidden.Should().BeFalse();
    }

    [Fact]
    public void Row_HideMultiple_HidesCorrectRows()
    {
        _ws.Row(1).Hidden = true;
        _ws.Row(3).Hidden = true;

        _ws.Row(1).Hidden.Should().BeTrue();
        _ws.Row(2).Hidden.Should().BeFalse();
        _ws.Row(3).Hidden.Should().BeTrue();
    }

    // ── Row Height / Column Width ───────────────────────────────────

    [Fact]
    public void Row_SetHeight_StoresHeight()
    {
        _ws.Row(1).Height = 30.0;
        _ws.Row(1).Height.Should().Be(30.0);
    }

    [Fact]
    public void Column_SetWidth_StoresWidth()
    {
        _ws.Column(1).Width = 25.0;
        _ws.Column(1).Width.Should().Be(25.0);
    }

    // ── Group/Ungroup (mirrors MssRow_GroupByIndex etc.) ───────────

    [Fact]
    public void Row_GroupCollapsed_CollapsesGroup()
    {
        // MssRow_GroupByIndex with ssCollapsed = true
        for (int i = 1; i <= 5; i++)
            _ws.Row(i).OutlineLevel = 1;

        _ws.Row(1).Collapsed = true;
        _ws.Row(1).Collapsed.Should().BeTrue();
    }

    [Fact]
    public void Column_GroupCollapsed_CollapsesGroup()
    {
        for (int i = 1; i <= 5; i++)
            _ws.Column(i).OutlineLevel = 1;

        _ws.Column(1).Collapsed = true;
        _ws.Column(1).Collapsed.Should().BeTrue();
    }
}
