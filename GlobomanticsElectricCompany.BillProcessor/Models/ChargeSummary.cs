using System;

namespace GlobomanticsElectricCompany.BillProcessor.Models
{
    public class ChargeSummary
    {
        public decimal PriorCharge { get; set; }
        public decimal PriorPaymentAmount { get; set; }
        public decimal CurrentChargeAmount { get; set; }
        public decimal SalesTaxRate { get; set; }
        public decimal SalesTaxValue => (SalesTaxRate / 100) * CurrentChargeAmount;
        public decimal AmountDue => (PriorCharge - PriorPaymentAmount) + CurrentChargeAmount + SalesTaxValue;

        public DateTime PriorPaymentDate { get; set; }
        public DateTime DueDate { get; set; }
        public DateTime BillCreatedDate { get; set; }

        public bool IsAutoPay { get; set; }
        public bool IsPastDue => PriorPaymentAmount < PriorCharge;
    }
}
