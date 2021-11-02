namespace GlobomanticsElectricCompany.BillProcessor.Models
{
    public class ContactInfo
    {
        public string Line1 { get; set; }
        public string Line2 { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string Zip { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string CityStateZip => $"{City}, {State} {Zip}";
    }
}
