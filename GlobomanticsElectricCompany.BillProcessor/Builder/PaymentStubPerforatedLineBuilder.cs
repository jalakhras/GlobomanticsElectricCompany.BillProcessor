using Aspose.Words;
using Aspose.Words.Drawing;

namespace GlobomanticsElectricCompany.BillProcessor.Builder
{
    public class PaymentStubPerforatedLineBuilder
    {
        public static void Build(DocumentBuilder builder)
        {
            var line = new Shape(builder.Document, ShapeType.Line)
            {
                RelativeVerticalPosition = RelativeVerticalPosition.Margin,
                RelativeHorizontalPosition = RelativeHorizontalPosition.Margin,
                Top = 530,
                Width = 575,
                HorizontalAlignment = HorizontalAlignment.Center,
                StrokeWeight = 2,
                Stroked = true,
                Stroke = { DashStyle = DashStyle.ShortDot }
            };

            builder.InsertNode(line);

            var lineMessageTextBox = new Shape(builder.Document, ShapeType.TextBox)
            {
                RelativeHorizontalPosition = RelativeHorizontalPosition.Margin,
                RelativeVerticalPosition = RelativeVerticalPosition.Margin,
                Top = 510,
                HorizontalAlignment = HorizontalAlignment.Center,
                Stroked = false,
                Width = 450,
                Height = 20
            };

            var lineMessagePargraph = new Paragraph(builder.Document);
            var lineMessageRun = new Run(builder.Document, "PLEASE DETACH LOWER PORTION" +" AND RETURN WITH PAYMENT");
            lineMessageRun.Font.Name = GlobalDocumentSettings.FontName;
            lineMessageRun.Font.Size = GlobalDocumentSettings.FontSize;
            lineMessagePargraph.AppendChild(lineMessageRun);
            lineMessagePargraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            lineMessageTextBox.AppendChild(lineMessagePargraph);
            builder.Document.FirstSection.Body.FirstParagraph.AppendChild(lineMessageTextBox);
        }
    }
}