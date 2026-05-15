using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace WordTemplateBlobWebApi.Utilities;

public class WordTemplateGeneratorSdt
{
    public static void GenerateTemplate(string outputPath)
    {
        using var doc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);

        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = mainPart.Document.AppendChild(new Body());

        // Title
        AddParagraphWithText(body, "PROPOSAL DOCUMENT", isBold: true, fontSize: 28);
        AddParagraphWithText(body, string.Empty);

        // Customer Name
        AddParagraphWithText(body, "Customer Name:");
        body.AppendChild(CreatePlainTextContentControl("CustomerName", "Customer Name", "John Doe"));
        AddParagraphWithText(body, string.Empty);

        // Project Name
        AddParagraphWithText(body, "Project Name:");
        body.AppendChild(CreatePlainTextContentControl("ProjectName", "Project Name", "Sample Project"));
        AddParagraphWithText(body, string.Empty);

        // Prepared By
        AddParagraphWithText(body, "Prepared By:");
        body.AppendChild(CreatePlainTextContentControl("PreparedBy", "Prepared By", "Proposal Team"));
        AddParagraphWithText(body, string.Empty);

        // Generated Date
        AddParagraphWithText(body, "Generated Date:");
        body.AppendChild(CreatePlainTextContentControl("GeneratedDate", "Generated Date", DateTime.UtcNow.ToString("yyyy-MM-dd")));
        AddParagraphWithText(body, string.Empty);

        // Separator
        AddParagraphWithText(body, new string('_', 50));
        AddParagraphWithText(body, string.Empty);

        // Line Items Table
        AddParagraphWithText(body, "LINE ITEMS:");
        body.AppendChild(CreateLineItemsTable());

        // Summary
        AddParagraphWithText(body, string.Empty);
        AddParagraphWithText(body, "Thank you for your consideration.", isItalic: true);

        mainPart.Document.Save();

        Console.WriteLine($"Word template created successfully at: {outputPath}");
    }

    /// <summary>
    /// Creates a block-level plain text content control (SdtBlock) wrapping a paragraph.
    /// The Tag is used for programmatic lookup; the Alias is the human-readable label shown in Word.
    /// </summary>
    private static SdtBlock CreatePlainTextContentControl(string tag, string alias, string defaultText)
    {
        var sdtBlock = new SdtBlock();

        // Properties
        var sdtPr = new SdtProperties();
        sdtPr.AppendChild(new SdtAlias { Val = alias });
        sdtPr.AppendChild(new Tag { Val = tag });
        sdtPr.AppendChild(new SdtId { Val = Math.Abs(tag.GetHashCode()) });
        sdtPr.AppendChild(new ShowingPlaceholder());

        sdtPr.AppendChild(new SdtPlaceholder());

        sdtBlock.AppendChild(sdtPr);

        // Content
        var sdtContent = new SdtContentBlock();
        var paragraph = new Paragraph();
        var run = new Run();
        run.AppendChild(new Text(defaultText) { Space = SpaceProcessingModeValues.Preserve });
        paragraph.AppendChild(run);
        sdtContent.AppendChild(paragraph);
        sdtBlock.AppendChild(sdtContent);

        return sdtBlock;
    }

    /// <summary>
    /// Creates an inline plain text content control (SdtRun) inside a table cell paragraph.
    /// </summary>
    private static TableCell CreateTableCellWithContentControl(string tag, string alias, string defaultText)
    {
        var cell = new TableCell();
        var cellProperties = new TableCellProperties();
        cellProperties.AppendChild(new TableCellWidth { Width = "auto", Type = TableWidthUnitValues.Auto });
        cell.AppendChild(cellProperties);

        var paragraph = new Paragraph();

        var sdtRun = new SdtRun();

        var sdtPr = new SdtProperties();
        sdtPr.AppendChild(new SdtAlias { Val = alias });
        sdtPr.AppendChild(new Tag { Val = tag });
        sdtPr.AppendChild(new SdtId { Val = Math.Abs(tag.GetHashCode()) });
        sdtRun.AppendChild(sdtPr);

        var sdtContent = new SdtContentRun();
        var run = new Run();
        run.AppendChild(new Text(defaultText) { Space = SpaceProcessingModeValues.Preserve });
        sdtContent.AppendChild(run);
        sdtRun.AppendChild(sdtContent);

        paragraph.AppendChild(sdtRun);
        cell.AppendChild(paragraph);
        return cell;
    }

    private static Table CreateLineItemsTable()
    {
        var table = new Table();

        var tblPr = new TableProperties();
        tblPr.AppendChild(new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct });
        table.AppendChild(tblPr);

        // Header row
        var headerRow = new TableRow();
        headerRow.AppendChild(CreateTableCell("Item Name", isBold: true));
        headerRow.AppendChild(CreateTableCell("Description", isBold: true));
        headerRow.AppendChild(CreateTableCell("Amount", isBold: true));
        table.AppendChild(headerRow);

        // Template row with content controls
        var templateRow = new TableRow();
        templateRow.AppendChild(CreateTableCellWithContentControl("ItemName", "Item Name", "Sample Item"));
        templateRow.AppendChild(CreateTableCellWithContentControl("ItemDescription", "Item Description", "Item description goes here"));
        templateRow.AppendChild(CreateTableCellWithContentControl("ItemAmount", "Item Amount", "0.00"));
        table.AppendChild(templateRow);

        return table;
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
}
