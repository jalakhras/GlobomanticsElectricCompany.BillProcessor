using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Data;
using System.Drawing;

namespace GlobomanticsElectricCompany.BillProcessor.Builder
{
    public class PaymentStubBuilder
    {
        public static void Build(DocumentBuilder builder)
        {
            var chargeSummaryTextbox = new Shape(builder.Document, ShapeType.TextBox)
            {
                RelativeHorizontalPosition = RelativeHorizontalPosition.Margin,
                RelativeVerticalPosition = RelativeVerticalPosition.Margin,
                Top = 550,
                Stroked = false,
                Width = 300,
                Height = 80,
                                HorizontalAlignment = HorizontalAlignment.Right

            };
            chargeSummaryTextbox.Left = (builder.PageSetup.PageWidth -
                                     (chargeSummaryTextbox.Width));

            var chargeSummaryDataTable = HelperMethods.CreateChargeSummaryDataTable();

            var table = builder.StartTable();

            foreach (DataRow row in chargeSummaryDataTable.Rows)
            {
                foreach (var item in row.ItemArray)
                {
                    builder.InsertCell();
                    builder.Write(item.ToString());
                }

                builder.EndRow();
            }
            table.AutoFit(AutoFitBehavior.AutoFitToContents);
            table.ClearBorders();
            table.SetBorder(BorderType.Left, LineStyle.Thick, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Right, LineStyle.Thick, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Top, LineStyle.Thick, 1.5, Color.Black, true);
            table.SetBorder(BorderType.Bottom, LineStyle.Thick, 1.5, Color.Black, true);

            builder.EndTable();
            chargeSummaryTextbox.AppendChild(table);

            builder.Document.FirstSection.Body.FirstParagraph.AppendChild(chargeSummaryTextbox);
        }
    }
}
