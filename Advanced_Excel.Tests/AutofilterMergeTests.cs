using OfficeOpenXml;
using FluentAssertions;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for comment add/delete, autofilter, merge/unmerge, and column/row insert/delete.
/// Mirrors MssComment_Add, MssComment_Delete,
/// MssWorksheet_AddAutoFilter, MssWorksheet_AutofitColumns,
/// MssRange_Merge, MssRange_UnMerge,
/// MssColumn_Insert, MssColumn_Delete, MssRow_Insert, MssRow_Delete.
/// </summary>
public class AutofilterMergeTests : IDisposable
{
    private readonly ExcelPackage _package;
    private bool _disposed;

    public AutofilterMergeTests()
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

    // ── Comments ────────────────────────────────────────────────────

    [Fact]
    public void Comment_Add_HasCorrectTextAndAuthor()
    {
        // MssComment_Add
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        var comment = ws.Comments.Add(ws.Cells[1, 1], "Test comment", "Author");

        comment.Text.Should().Be("Test comment");
        comment.Author.Should().Be("Author");
    }

    [Fact]
    public void Comment_Add_AutoFit_SetsProperty()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        var comment = ws.Comments.Add(ws.Cells[1, 1], "Long comment text", "Author");
        comment.AutoFit = true;

        comment.AutoFit.Should().BeTrue();
    }

    [Fact]
    public void Comment_Delete_RemovesComment()
    {
        // MssComment_Delete
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        var comment = ws.Comments.Add(ws.Cells[1, 1], "To remove", "Author");
        ws.Cells[1, 1].Comment.Should().NotBeNull();

        ws.Comments.Remove(comment);
        ws.Cells[1, 1].Comment.Should().BeNull();
    }

    [Fact]
    public void Comment_DeleteRange_RemovesAllInRange()
    {
        // MssComment_Delete with range
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Comments.Add(ws.Cells[1, 1], "C1", "A");
        ws.Comments.Add(ws.Cells[1, 2], "C2", "A");
        ws.Comments.Add(ws.Cells[1, 3], "C3", "A");
        ws.Cells[1, 2].Comment.Should().NotBeNull();

        // Remove comments in range A1:B2
        for (int row = 1; row <= 2; row++)
        {
            for (int col = 1; col <= 2; col++)
            {
                if (ws.Cells[row, col].Comment != null)
                    ws.Comments.Remove(ws.Cells[row, col].Comment);
            }
        }

        ws.Cells[1, 1].Comment.Should().BeNull();
        ws.Cells[1, 2].Comment.Should().BeNull();
        // C3 comment survives (col 3 is outside range)
        ws.Cells[1, 3].Comment.Should().NotBeNull();
    }

    // ── AutoFilter ──────────────────────────────────────────────────

    [Fact]
    public void AutoFilter_Add_FilterRange()
    {
        // MssWorksheet_AddAutoFilter
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Name";
        ws.Cells["B1"].Value = "Age";
        ws.Cells["A2"].Value = "Alice";
        ws.Cells["A3"].Value = "Bob";

        ws.AutoFilterAddress = new OfficeOpenXml.ExcelAddress("A1:B10");
        ws.AutoFilterAddress.Should().NotBeNull();
    }

    [Fact]
    public void AutoFilter_Add_NonZeroRows()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Header";
        ws.Cells["A2"].Value = "Data";

        ws.AutoFilterAddress = new OfficeOpenXml.ExcelAddress("A1:A2");
        ws.AutoFilterAddress.Address.Should().Be("A1:A2");
    }

    [Fact]
    public void AutoFilter_SkipEmptySheet_DoesNothing()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        // Empty sheet — MssWorksheet_AddAutoFilter skips when UsedRange is null

        // No exception should occur
        var act = () =>
        {
            if (ws.Dimension != null)
            {
                ws.AutoFilterAddress = new OfficeOpenXml.ExcelAddress("A1:A10");
            }
        };
        act.Should().NotThrow();
    }

    // ── AutoFit Columns ─────────────────────────────────────────────

    [Fact]
    public void AutofitColumns_DoesNotThrow()
    {
        // MssWorksheet_AutofitColumns
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Some content in cell";

        var act = () => ws.Cells.AutoFitColumns();
        act.Should().NotThrow();
    }

    // ── Merge ───────────────────────────────────────────────────────

    [Fact]
    public void Range_Merge_CellsBecomeMerged()
    {
        // MssRange_Merge
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:C3"].Merge = true;

        ws.Cells["A1:C3"].Merge.Should().BeTrue();
    }

    [Fact]
    public void Range_Merge_SingleCellAddressIsStartCell()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["B2:D5"].Merge = true;

        // After merge, the merged range address is the top-left cell
        ws.Cells["B2"].Merge.Should().BeTrue();
        ws.Cells["D5"].Merge.Should().BeTrue();
    }

    [Fact]
    public void Range_UnMerge_RemovesMerge()
    {
        // MssRange_UnMerge
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:C3"].Merge = true;
        ws.Cells["A1:C3"].Merge.Should().BeTrue();

        ws.Cells["A1:C3"].Merge = false;
        ws.Cells["A1:C3"].Merge.Should().BeFalse();
    }

    [Fact]
    public void Range_MergeThenUnMerge_NoLongerMerged()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:B2"].Merge = true;
        ws.Cells["A1:B2"].Merge = false;

        ws.Cells["A1:B2"].Merge.Should().BeFalse();
    }

    // ── Column Insert ───────────────────────────────────────────────

    [Fact]
    public void Column_Insert_AddsColumn()
    {
        // MssColumn_Insert
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "A";
        ws.Cells[1, 2].Value = "B";
        ws.Cells[1, 3].Value = "C";

        ws.InsertColumn(2, 1); // Insert 1 column at position 2

        // Original B should now be at column 3
        ws.Cells[1, 3].Text.Should().Be("B");
        ws.Cells[1, 4].Text.Should().Be("C");
    }

    [Fact]
    public void Column_Insert_Multiple_AddsCorrectCount()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "A";
        ws.Cells[1, 2].Value = "B";

        ws.InsertColumn(2, 2); // Insert 2 columns at position 2

        ws.Cells[1, 4].Text.Should().Be("B");
    }

    // ── Column Delete ───────────────────────────────────────────────

    [Fact]
    public void Column_Delete_RemovesColumn()
    {
        // MssColumn_Delete
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "A";
        ws.Cells[1, 2].Value = "B";
        ws.Cells[1, 3].Value = "C";

        ws.DeleteColumn(2); // Delete column 2

        ws.Cells[1, 1].Text.Should().Be("A");
        ws.Cells[1, 2].Text.Should().Be("C");
    }

    [Fact]
    public void Column_Delete_WithComments_RemovesCommentsFirst()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "A";
        ws.Cells[1, 2].Value = "B";
        ws.Cells[2, 1].Value = "X";
        ws.Comments.Add(ws.Cells[1, 1], "Comment on A", "Author");

        // Delete comments in column before deleting column
        for (int row = 1; row <= ws.Dimension?.End.Row; row++)
        {
            if (ws.Cells[row, 1].Comment != null)
                ws.Comments.Remove(ws.Cells[row, 1].Comment);
        }

        ws.DeleteColumn(1);
        ws.Cells[1, 1].Text.Should().Be("B");
    }

    // ── Row Insert ──────────────────────────────────────────────────

    [Fact]
    public void Row_Insert_AddsRow()
    {
        // MssRow_Insert
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "Row1";
        ws.Cells[2, 1].Value = "Row2";
        ws.Cells[3, 1].Value = "Row3";

        ws.InsertRow(2, 1); // Insert 1 row at position 2

        ws.Cells[1, 1].Text.Should().Be("Row1");
        ws.Cells[3, 1].Text.Should().Be("Row2");
        ws.Cells[4, 1].Text.Should().Be("Row3");
    }

    [Fact]
    public void Row_Insert_Multiple_AddsCorrectCount()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "A";
        ws.Cells[2, 1].Value = "B";

        ws.InsertRow(2, 3); // Insert 3 rows at position 2

        ws.Cells[1, 1].Text.Should().Be("A");
        ws.Cells[5, 1].Text.Should().Be("B");
    }

    // ── Row Delete ──────────────────────────────────────────────────

    [Fact]
    public void Row_Delete_RemovesRow()
    {
        // MssRow_Delete
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "R1";
        ws.Cells[2, 1].Value = "R2";
        ws.Cells[3, 1].Value = "R3";

        ws.DeleteRow(2);

        ws.Cells[1, 1].Text.Should().Be("R1");
        ws.Cells[2, 1].Text.Should().Be("R3");
    }

    [Fact]
    public void Row_Delete_Multiple_RemovesCorrectCount()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells[1, 1].Value = "R1";
        ws.Cells[2, 1].Value = "R2";
        ws.Cells[3, 1].Value = "R3";
        ws.Cells[4, 1].Value = "R4";

        ws.DeleteRow(2, 2); // Delete 2 rows starting at row 2

        ws.Cells[1, 1].Text.Should().Be("R1");
        ws.Cells[2, 1].Text.Should().Be("R4");
    }

    // ── Save/Serialize After Operations ─────────────────────────────

    [Fact]
    public void Save_WorksheetWithMergedCells_PreservesMerge()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1:B2"].Merge = true;
        ws.Cells["A1"].Value = "Merged";

        byte[] data = _package.GetAsByteArray();

        using var reopened = new ExcelPackage(new MemoryStream(data));
        reopened.Workbook.Worksheets[0].Cells["A1:B2"].Merge.Should().BeTrue();
    }

    [Fact]
    public void Save_WorksheetWithAutofilter_PreservesFilter()
    {
        var ws = _package.Workbook.Worksheets.Add("Sheet1");
        ws.Cells["A1"].Value = "Header";
        ws.Cells["A2"].Value = "Data";
        ws.AutoFilterAddress = new OfficeOpenXml.ExcelAddress("A1:A10");

        byte[] data = _package.GetAsByteArray();

        using var reopened = new ExcelPackage(new MemoryStream(data));
        reopened.Workbook.Worksheets[0].AutoFilterAddress.Should().NotBeNull();
    }
}
