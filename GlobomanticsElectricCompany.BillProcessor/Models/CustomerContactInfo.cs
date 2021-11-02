namespace GlobomanticsElectricCompany.BillProcessor.Models
{
    public class CustomerContactInfo : ContactInfo
    {
        public int AccountNumber { get;set; }

        public string FirstName { get; set; }
        public string LastName { get; set; }

        public bool IsAutoPay { get; set; }
    }
}
