using Aspose.Words;

namespace GlobomanticsElectricCompany.BillProcessor.Builder
{
    public class CompanyContactInfoBuilder
    {
        public static void Build(DocumentBuilder builder)
        {
            builder.Writeln();
            builder.Writeln(HelperMethods.CreateCompanyContactInfoText(true));
        }
    }
}
