using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace GlobomanticsElectricCompany.BillProcessor.Builder
{
    public class BillSummaryTableBuilder
    {
        public static void Build(Document doc, DocumentBuilder builder)
        {
            var summaryTextBox = new Shape(doc, ShapeType.TextBox)
            {
                Width = 182,
                Height = 65,
                Stroked = false,
                WrapType = WrapType.None,
                RelativeHorizontalPosition = RelativeHorizontalPosition.Margin,
                RelativeVerticalPosition = RelativeVerticalPosition.Margin,
                Top = 35,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            var table = new Table(doc);
            var chargeSummary = HelperMethods.CreateChargeSummary();
            var customerInfo = HelperMethods.CreateCustomerContactInfo();

            doc.FirstSection.Body.AppendChild(table);

            var accountNumberRow = new Row(doc);

            var accountNumberLabelCell = new Cell(doc);
            var accountNumberLabelParagraph = new Paragraph(doc);
            var accountNumberLabelRun = new Run(doc, "Account Number");

            accountNumberLabelParagraph.AppendChild(accountNumberLabelRun);
            accountNumberLabelCell.AppendChild(accountNumberLabelParagraph);
            accountNumberRow.AppendChild(accountNumberLabelCell);

            var accountNumberValueCell = new Cell(doc);
            var accountNumberValueParagraph = new Paragraph(doc);
            var accountNumberValueRun = new Run(doc, customerInfo.AccountNumber.ToString());

            accountNumberValueParagraph.AppendChild(accountNumberValueRun);
            accountNumberValueCell.AppendChild(accountNumberValueParagraph);
            accountNumberRow.AppendChild(accountNumberValueCell);

            var amountDueRow = new Row(doc);

            var amountDueLabelCell = new Cell(doc);
            var amountDueLabelParagraph = new Paragraph(doc);
            var amountDueLabelRun = new Run(doc, "Amount Due");

            amountDueLabelParagraph.AppendChild(amountDueLabelRun);
            amountDueLabelCell.AppendChild(amountDueLabelParagraph);
            amountDueRow.AppendChild(amountDueLabelCell);

            var amountDueValueCell = new Cell(doc);
            var amountDueValueParagraph = new Paragraph(doc);
            var amountDueValueRun = new Run(doc, chargeSummary.AmountDue.ToString("C"));

            amountDueValueParagraph.AppendChild(amountDueValueRun);
            amountDueValueCell.AppendChild(amountDueValueParagraph);
            amountDueRow.AppendChild(amountDueValueCell);

            var dueDateRow = new Row(doc);

            var dueDateLabelCell = new Cell(doc);
            var dueDateLabelParagraph = new Paragraph(doc);
            var dueDateLabelRun = new Run(doc, "Due Date");

            dueDateLabelParagraph.AppendChild(dueDateLabelRun);
            dueDateLabelCell.AppendChild(dueDateLabelParagraph);
            dueDateRow.AppendChild(dueDateLabelCell);

            var dueDateValueCell = new Cell(doc);
            var dueDateValueParagraph = new Paragraph(doc);
            var dueDateValueRun = new Run(doc, chargeSummary.DueDate.ToShortDateString());

            dueDateValueParagraph.AppendChild(dueDateValueRun);
            dueDateValueCell.AppendChild(dueDateValueParagraph);
            dueDateRow.AppendChild(dueDateValueCell);

            table.AppendChild(accountNumberRow);
            table.AppendChild(amountDueRow);
            table.AppendChild(dueDateRow);

            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
                    var runs = cell.GetChildNodes(NodeType.Run, true);
                    foreach (Run run in runs)
                    {
                        run.Font.Name = GlobalDocumentSettings.FontName;
                        run.Font.Size = GlobalDocumentSettings.FontSize;
                    }

                    //Center text horizontally
                    foreach (Paragraph paragraph in cell.Paragraphs)
                    {
                        paragraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    }
                }
            }

            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            summaryTextBox.AppendChild(table);

            builder.InsertNode(summaryTextBox);
        }
    }
}