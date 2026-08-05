using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Models
{
    public class GerentePickingForm
    {
        public List<GerentePickingFormRow> Rows {  get; set; }

        public GerentePickingForm()
        {
            Rows = new List<GerentePickingFormRow>();
        }
    }

    public class GerentePickingFormRow
    {
        public int DocEntry { get; set; }
        public int LineNum { get; set; }
        public int LineIdRdr1 { get; set; }
        public string ItemCode { get; set; }
        public decimal AvailableStock { get; set; }
        public string WhsCode { get; set; }
        public int RelatedOrd { get; set; }
        public decimal SplitPercentage { get; set; }
        public string Tipo { get; set; }
        public decimal QtyOpen { get; set; }
        public bool QuantityExceedsAvailable { get; set; }
        public decimal ToRelease { get; set; }
        public decimal ItemsPerUnit { get; set; }
        public decimal AvailableForRelease { get; set; }

    }

    //public class  DataSplit
    //{
    //    public int SplitPercA { get; set; }
    //    public int SplitPercB { get; set; }
    //    public decimal OpenQtyA { get; set; }
    //    public decimal OpenQtyB { get; set; }
    //    public decimal TotalQty { get; set; }
    //    public decimal AvailableStock { get; set; }
    //}
}
