using Aspose.Words;
using System.Drawing;

namespace GlobomanticsElectricCompany.BillProcessor
{
    public class GlobalDocumentSettings
    {
        //Global font properties
        public static string FontName => "Gotham Book";
        public static int FontSize => 12;

        //File information
        public static string Filename => "E:\\Courses\\Paths\\Aspose.Words for .NET Creating Dynamic Documents\\_Project\\GlobomanticsElectricCompany.BillProcessor\\BillDemo.docx";

        //Set page margins
        public static void SetPageMargins(Document doc)
        {
            foreach (Section section in doc.Sections)
            {
                section.PageSetup.TopMargin = 36;
                section.PageSetup.LeftMargin = 36;
                section.PageSetup.BottomMargin = 36;
                section.PageSetup.RightMargin = 36;
            }
        }
    }
}
