using GlobomanticsElectricCompany.BillProcessor.Builder;

namespace GlobomanticsElectricCompany.BillProcessor
{
    public class BillProcessor
    {
        public static void Main()
        {            
            //Initialize document object and document builder object
            

            //Create company contact info
            CompanyContactInfoBuilder.Build();

            //Create bill summary table inside a text box
            BillSummaryTableBuilder.Build();

            //Create bill details table using Document Builder
            BillDetailsTableBuilder.Build();

            //Create perforated line for payment stub
            PaymentStubPerforatedLineBuilder.Build();

            //Create payment stub
            PaymentStubBuilder.Build();

            //Set global page margins
            GlobalDocumentSettings.SetPageMargins();

            //Save document

        }
    }
}
