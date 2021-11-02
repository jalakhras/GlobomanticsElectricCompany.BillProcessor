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
        }
    }
}