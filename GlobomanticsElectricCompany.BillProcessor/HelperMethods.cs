using System;
using System.Data;
using System.Text;
using GlobomanticsElectricCompany.BillProcessor.Models;

namespace GlobomanticsElectricCompany.BillProcessor
{
    public class HelperMethods
    {
        public static CompanyContactInfo CreateCompanyContactInfo()
        {
            return new CompanyContactInfo
            {
                CompanyName = "Globomantics Electric Company",
                Line1 = "1234 Power Grid Way",
                City = "Electri City",
                State = "PA",
                Zip = "98765-4321",
                Phone = "1-800-POWER-UP",
                Email = "help@Globomantics.kwh"
            };
        }

        public static string CreateCompanyContactInfoText(bool includeAllContactInfo)
        {
            var contactInfo = CreateCompanyContactInfo();

            var sb = new StringBuilder();

            if (includeAllContactInfo)
            {
                sb.AppendLine(contactInfo.Line1);
                sb.AppendLine(contactInfo.CityStateZip);
                sb.AppendLine(contactInfo.Phone);
                sb.AppendLine(contactInfo.Email);
            }
            else
            {
                sb.AppendLine(contactInfo.CompanyName);
                sb.AppendLine(contactInfo.Line1);
                sb.AppendLine(contactInfo.CityStateZip);
            }
            return sb.ToString();
        }

        public static CustomerContactInfo CreateCustomerContactInfo()
        {
            return new CustomerContactInfo
            {
                AccountNumber = 12345,
                FirstName = "Happy",
                LastName = "Customer",
                Line1 = "9876 North Main Street",
                City = "Suburbia",
                State = "PA",
                Zip = "98764-3214",
                Phone = "555-555-4321",
                Email = "hcustomer@fakemail.net",
                IsAutoPay = true
            };
        }

        public static string CreateCustomerContactInfoText()
        {
            var contactInfo = CreateCustomerContactInfo();

            var sb = new StringBuilder();

            sb.AppendLine($"{contactInfo.FirstName} {contactInfo.LastName}");
            sb.AppendLine(contactInfo.Line1);
            sb.AppendLine(contactInfo.CityStateZip);

            return sb.ToString();
        }

        public static ChargeSummary CreateChargeSummary()
        {
            return new ChargeSummary
            {
                PriorCharge = 78.75M,
                PriorPaymentDate = new DateTime(2016, 01, 05),
                PriorPaymentAmount = 78.75M,
                CurrentChargeAmount = 87.43M,
                SalesTaxRate = 1.00M,
                DueDate = new DateTime(2016, 02, 05),
                BillCreatedDate = new DateTime(2016, 01, 21),
                IsAutoPay = true
            };
        }

        public static DataTable CreateChargeSummaryDataTable()
        {
            var chargeSummary = CreateChargeSummary();
            var accountInfo = CreateCustomerContactInfo();

            var dt = new DataTable();

            dt.Columns.Add();
            dt.Columns.Add();

            dt.Rows.Add("Account Number", accountInfo.AccountNumber.ToString());
            dt.Rows.Add("Amount Due", chargeSummary.AmountDue.ToString("C"));
            dt.Rows.Add("Payment Due Date",
                accountInfo.IsAutoPay ? "AUTOPAY" : chargeSummary.DueDate.ToShortDateString());

            return dt;
        }
    }
}