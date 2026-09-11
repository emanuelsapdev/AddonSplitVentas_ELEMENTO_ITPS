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
        public int PaymentGroupCode { get; set; }
        public int GlobalAgreement { get; set; }
        public string AgreementPriceList { get; set; } = string.Empty;
        public string Address2 { get; set; }
        public string ShipToCode { get; set; }
        public string Address { get; set; }
        public string PayToCode { get; set; }

        // Direcciones ----------------------
        public string StreetS { get; set; }
        public string StreetNoS { get; set; }
        public string BlockS { get; set; }
        public string CityS { get; set; }
        public string ZipCodeS { get; set; }
        public string CountyS { get; set; }
        public string StateS { get; set; }
        public string CountryS { get; set; }
        public string BuildingS { get; set; }
        public string Address2S { get; set; }
        public string Address3S { get; set; }
        public string GlbLocNumS { get; set; }
        public string TransportistaS { get; set; }
        public string DeliveryZoneS { get; set; }

        public string StreetB { get; set; }
        public string StreetNoB { get; set; }
        public string BlockB { get; set; }
        public string CityB { get; set; }
        public string ZipCodeB { get; set; }
        public string CountyB { get; set; }
        public string StateB { get; set; }
        public string CountryB { get; set; }
        public string BuildingB { get; set; }
        public string Address2B { get; set; }
        public string Address3B { get; set; }
        public string GlbLocNumB { get; set; }
        public string TransportistaB { get; set; }
        public string DeliveryZoneB { get; set; }

        // Lineas -------------------------
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
