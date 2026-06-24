using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Models
{
    public class SalesOrderFormModel
    {
        public string CardCode { get; set; }
        public decimal SplitPercentage { get; set; } = 0m;
        public decimal ImportedPercentage { get; set; } = 0m;
        public string AssignedEntity { get; set; }
        public string RelatedOrder { get; set; } = "0";
        public string CategoryClient { get; set; }
        public string DocDate { get; set; }
        public string DocDueDate { get; set; }
        public string TaxDate { get; set; }
        public string Comments { get; set; }
        public int DocEntry { get; set; }

        public bool IsImportOrder { get; set; } = false;
        public decimal ImportLineDiscount { get; set; } = 0m;
        public decimal TotalDiscountPercent { get; set; }

        public List<SalesOrderLineModel> Lines { get; set; } = new List<SalesOrderLineModel>();
    }

    public class SalesOrderLineModel
    {
        public string ItemCode { get; set; }
        public decimal Quantity { get; set; }
        public decimal UnitPrice { get; set; }
        public string UomCode { get; set; }
        public string UomEntry { get; set; }
        public string WhsCode { get; set; }
        public decimal Discount { get; set; }
        public string TaxCode { get; set; }
        public int LineNum { get; set; } = -1;
        public int LineId { get; set; } = -1;
        public string AgrNo { get; set; } = string.Empty;
        public string LineStatus { get; set; }
    }
}
