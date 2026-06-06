using OfficeOpenXml;
using OfficeOpenXml.Style;
using FluentAssertions;

namespace AdvancedExcel.Tests;

/// <summary>
/// Tests for document property get/set/clear operations.
/// Mirrors MssProperty_Set, MssProperty_Get, MssProperty_Clear.
/// </summary>
public class PropertyTests : IDisposable
{
    private readonly ExcelPackage _package;
    private bool _disposed;

    public PropertyTests()
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
    public void Property_SetTitle_RoundTrips()
    {
        _package.Workbook.Properties.Title = "Test Workbook";
        _package.Workbook.Properties.Title.Should().Be("Test Workbook");
    }

    [Fact]
    public void Property_SetAuthor_RoundTrips()
    {
        _package.Workbook.Properties.Author = "Test Author";
        _package.Workbook.Properties.Author.Should().Be("Test Author");
    }

    [Fact]
    public void Property_SetSubject_RoundTrips()
    {
        _package.Workbook.Properties.Subject = "Test Subject";
        _package.Workbook.Properties.Subject.Should().Be("Test Subject");
    }

    [Fact]
    public void Property_SetKeywords_RoundTrips()
    {
        _package.Workbook.Properties.Keywords = "excel, test, automation";
        _package.Workbook.Properties.Keywords.Should().Be("excel, test, automation");
    }

    [Fact]
    public void Property_SetDescription_RoundTrips()
    {
        _package.Workbook.Properties.Description = "A test workbook";
        _package.Workbook.Properties.Description.Should().Be("A test workbook");
    }

    [Fact]
    public void Property_SetCategory_RoundTrips()
    {
        _package.Workbook.Properties.Category = "Testing";
        _package.Workbook.Properties.Category.Should().Be("Testing");
    }

    [Fact]
    public void Property_SetCompany_RoundTrips()
    {
        _package.Workbook.Properties.Company = "Acme Corp";
        _package.Workbook.Properties.Company.Should().Be("Acme Corp");
    }

    [Fact]
    public void Property_SetManager_RoundTrips()
    {
        _package.Workbook.Properties.Manager = "Jane Doe";
        _package.Workbook.Properties.Manager.Should().Be("Jane Doe");
    }

    [Fact]
    public void Property_SetCreated_RoundTrips()
    {
        var created = new DateTime(2024, 1, 15, 10, 30, 0);
        _package.Workbook.Properties.Created = created;
        _package.Workbook.Properties.Created.Should().Be(created);
    }

    [Fact]
    public void Property_SetModifiedBy_RoundTrips()
    {
        _package.Workbook.Properties.ModifiedBy = "TestUser";
        _package.Workbook.Properties.ModifiedBy.Should().Be("TestUser");
    }

    [Fact]
    public void Property_Get_ReturnsSetValue()
    {
        // MssProperty_Get: set then get
        _package.Workbook.Properties.Title = "My Title";
        string title = _package.Workbook.Properties.Title;
        title.Should().Be("My Title");
    }

    [Fact]
    public void Property_Clear_RemovesValue()
    {
        // MssProperty_Clear: set, clear, verify empty
        _package.Workbook.Properties.Title = "To Clear";
        _package.Workbook.Properties.Title.Should().Be("To Clear");

        _package.Workbook.Properties.Title = string.Empty;
        _package.Workbook.Properties.Title.Should().BeNullOrEmpty();
    }

    [Fact]
    public void Property_ClearAll_RemovesAllProperties()
    {
        _package.Workbook.Properties.Title = "T";
        _package.Workbook.Properties.Author = "A";
        _package.Workbook.Properties.Subject = "S";

        _package.Workbook.Properties.Title = string.Empty;
        _package.Workbook.Properties.Author = string.Empty;
        _package.Workbook.Properties.Subject = string.Empty;

        _package.Workbook.Properties.Title.Should().BeNullOrEmpty();
        _package.Workbook.Properties.Author.Should().BeNullOrEmpty();
        _package.Workbook.Properties.Subject.Should().BeNullOrEmpty();
    }

    [Fact]
    public void Property_SetAndGet_MultipleProperties()
    {
        _package.Workbook.Properties.Title = "Multi";
        _package.Workbook.Properties.Author = "Author";
        _package.Workbook.Properties.Company = "Company";

        _package.Workbook.Properties.Title.Should().Be("Multi");
        _package.Workbook.Properties.Author.Should().Be("Author");
        _package.Workbook.Properties.Company.Should().Be("Company");
    }
}
