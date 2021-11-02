﻿using Aspose.Words;
using GlobomanticsElectricCompany.BillProcessor.Builder;

namespace GlobomanticsElectricCompany.BillProcessor
{
    public class BillProcessor
    {
        public static void Main()
        {
            //Initialize document object and document builder object
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            //Create company contact info
            CompanyContactInfoBuilder.Build(builder);

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
            doc.Save(GlobalDocumentSettings.Filename);

        }
    }
}
