using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace WordTemplateBlobWebApi.Utilities;

public class WordTemplateGenerator
{
    public static void GenerateTemplate(string outputPath)
    {
        using (var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            uint bookmarkId = 0;

            // Title
            AddParagraphWithText(body, "PROPOSAL DOCUMENT", isBold: true, fontSize: 28);
            AddParagraphWithText(body, string.Empty);

            // Customer Name
            AddParagraphWithText(body, "Customer Name:");
            body.AppendChild(CreateParagraphWithBookmark("CustomerName", "John Doe", bookmarkId++));
            AddParagraphWithText(body, string.Empty);

            // Project Name
            AddParagraphWithText(body, "Project Name:");
            body.AppendChild(CreateParagraphWithBookmark("ProjectName", "Sample Project", bookmarkId++));
            AddParagraphWithText(body, string.Empty);

            // Prepared By
            AddParagraphWithText(body, "Prepared By:");
            body.AppendChild(CreateParagraphWithBookmark("PreparedBy", "Proposal Team", bookmarkId++));
            AddParagraphWithText(body, string.Empty);

            // Generated Date
            AddParagraphWithText(body, "Generated Date:");
            body.AppendChild(CreateParagraphWithBookmark("GeneratedDate", DateTime.UtcNow.ToString("yyyy-MM-dd"), bookmarkId++));
            AddParagraphWithText(body, string.Empty);

            // Separator
            AddParagraphWithText(body, new string('_', 50));
            AddParagraphWithText(body, string.Empty);

            // Line Items Table
            AddParagraphWithText(body, "LINE ITEMS:");
            var table = CreateLineItemsTable(bookmarkId);
            body.AppendChild(table);

            // Summary
            AddParagraphWithText(body, string.Empty);
            AddParagraphWithText(body, "Thank you for your consideration.", isItalic: true);

            mainPart.Document.Save();
        }

        Console.WriteLine($"Word template created successfully at: {outputPath}");
    }

    private static void AddParagraphWithText(Body body, string text, bool isBold = false, bool isItalic = false, int fontSize = 22)
    {
        var paragraph = new Paragraph();
        var run = new Run();

        if (isBold || isItalic || fontSize != 22)
        {
            var runProperties = new RunProperties();
            if (isBold)
                runProperties.AppendChild(new Bold());
            if (isItalic)
                runProperties.AppendChild(new Italic());
            if (fontSize != 22)
                runProperties.AppendChild(new FontSize { Val = fontSize.ToString() });

            run.AppendChild(runProperties);
        }

        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        paragraph.AppendChild(run);
        body.AppendChild(paragraph);
    }

    private static Paragraph CreateParagraphWithBookmark(string bookmarkName, string content, uint bookmarkId)
    {
        var paragraph = new Paragraph();

        var bookmarkStart = new BookmarkStart { Name = bookmarkName, Id = bookmarkId.ToString() };
        paragraph.AppendChild(bookmarkStart);

        var run = new Run();
        run.AppendChild(new Text(content));
        paragraph.AppendChild(run);

        var bookmarkEnd = new BookmarkEnd { Id = bookmarkId.ToString() };
        paragraph.AppendChild(bookmarkEnd);

        return paragraph;
    }

    private static Table CreateLineItemsTable(uint startingBookmarkId)
    {
        var table = new Table();

        // Table properties
        var tblPr = new TableProperties();
        var tblW = new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct };
        tblPr.AppendChild(tblW);
        table.AppendChild(tblPr);

        // Header row
        var headerRow = new TableRow();
        headerRow.AppendChild(CreateTableCell("Item Name", isBold: true));
        headerRow.AppendChild(CreateTableCell("Description", isBold: true));
        headerRow.AppendChild(CreateTableCell("Amount", isBold: true));
        table.AppendChild(headerRow);

        // Template row with bookmarks (this will be cloned for each item)
        var templateRow = new TableRow();
        templateRow.AppendChild(CreateTableCellWithBookmark("ItemName", "Sample Item", startingBookmarkId));
        templateRow.AppendChild(CreateTableCellWithBookmark("ItemDescription", "Item description goes here", startingBookmarkId + 1));
        templateRow.AppendChild(CreateTableCellWithBookmark("ItemAmount", "/bin/zsh.00", startingBookmarkId + 2));
        table.AppendChild(templateRow);

        return table;
    }

    private static TableCell CreateTableCell(string text, bool isBold = false)
    {
        var cell = new TableCell();
        var cellProperties = new TableCellProperties();
        cellProperties.AppendChild(new TableCellWidth { Width = "auto", Type = TableWidthUnitValues.Auto });
        cell.AppendChild(cellProperties);

        var paragraph = new Paragraph();
        var run = new Run();

        if (isBold)
        {
            var runProperties = new RunProperties();
            runProperties.AppendChild(new Bold());
            run.AppendChild(runProperties);
        }

        run.AppendChild(new Text(text));
        paragraph.AppendChild(run);
        cell.AppendChild(paragraph);

        return cell;
    }

    private static TableCell CreateTableCellWithBookmark(string bookmarkName, string content, uint bookmarkId)
    {
        var cell = new TableCell();
        var cellProperties = new TableCellProperties();
        cellProperties.AppendChild(new TableCellWidth { Width = "auto", Type = TableWidthUnitValues.Auto });
        cell.AppendChild(cellProperties);

        var paragraph = new Paragraph();
        var bookmarkStart = new BookmarkStart { Name = bookmarkName, Id = bookmarkId.ToString() };
        paragraph.AppendChild(bookmarkStart);

        var run = new Run();
        run.AppendChild(new Text(content));
        paragraph.AppendChild(run);

        var bookmarkEnd = new BookmarkEnd { Id = bookmarkId.ToString() };
        paragraph.AppendChild(bookmarkEnd);

        cell.AppendChild(paragraph);
        return cell;
    }
}
