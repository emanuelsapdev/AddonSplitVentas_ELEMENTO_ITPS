namespace Addon_AutoDivSalesOrd.Models
{
    public class CreditNoteFormModel
    {
        public int DocEntry { get; set; }
        public string Canceled { get; set; }
        public string Importado { get; set; }
        public decimal ImportedPercentage { get; set; } = 0m;
    }
}
