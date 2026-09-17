using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Forms.SalesOrder
{
    public partial class SalesOrderFrm
    {
        public SalesOrderFormModel GetDataFromFormOrder(SAPbouiCOM.Form oForm)
        {
            var mdl = new SalesOrderFormModel();

            DBDataSource db = oForm.DataSources.DBDataSources.Item(Constants.DataSources.SalesOrderHeader);
            DBDataSource db1 = oForm.DataSources.DBDataSources.Item(Constants.DataSources.SalesOrderLines);
            DBDataSource db12 = oForm.DataSources.DBDataSources.Item(Constants.DataSources.SalesOrderAddresses);

            // HEADER
            var oCardCode = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_CardCode).Specific;
            //var oSplitPercentage = (SAPbouiCOM.ComboBox)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_SplitPercentage).Specific;
            //var oImportedPercentage = (SAPbouiCOM.ComboBox)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_ImportedPercentage).Specific;
            var oRelatedOrder = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_RelatedOrder).Specific;
            var oDocDate = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_DocDate).Specific;
            var oDocDueDate = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_DocDueDate).Specific;
            var oTaxDate = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_TaxDate).Specific;
            var oComments = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Comments).Specific;
            var oCategoryClient = (ComboBox)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_CategoryClient).Specific;
            var oDiscPrcnt = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_DiscPrcnt).Specific;
            var oPaymentGroupCode = (ComboBox)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_PaymentGroupCode).Specific;
            var oAddress2 = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Address2).Specific;
            var oShipToCode = (ComboBox)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_ShipToCode).Specific;
            var oAddress = (EditText)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Address).Specific;
            var oPayToCode = (ComboBox)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_PayToDate).Specific;
            var oMtx = (Matrix)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Matrix).Specific;

            // Dirección Entrega
            var oStreetS = db12.GetValue("StreetS", 0);
            var oStreetNoS = db12.GetValue("StreetNoS", 0);
            var oBlockS = db12.GetValue("BlockS", 0);
            var oCityS = db12.GetValue("CityS", 0);
            var oZipCodeS = db12.GetValue("ZipCodeS", 0);
            var oCountyS = db12.GetValue("CountyS", 0);
            var oStateS = db12.GetValue("StateS", 0);
            var oCountryS = db12.GetValue("CountryS", 0);
            var oBuildingS = db12.GetValue("BuildingS", 0);
            var oAddress2S = db12.GetValue("Address2S", 0);
            var oAddress3S = db12.GetValue("Address3S", 0);
            var oGlbLocNumS = db12.GetValue("GlbLocNumS", 0);
            var oTransportistaS = db12.GetValue("U_ITPS_TransportistaS", 0);
            var oDeliveryZoneS = db12.GetValue("U_ITPS_DeliveryZoneS", 0);

            mdl.StreetS = oStreetS;
            mdl.StreetNoS = oStreetNoS;
            mdl.BlockS = oBlockS;
            mdl.CityS = oCityS;
            mdl.ZipCodeS = oZipCodeS;
            mdl.CountyS = oCountyS;
            mdl.StateS = oStateS;
            mdl.CountryS = oCountryS;
            mdl.BuildingS = oBuildingS;
            mdl.Address2S = oAddress2S;
            mdl.Address3S = oAddress3S;
            mdl.GlbLocNumS = oGlbLocNumS;
            mdl.TransportistaS = oTransportistaS;
            mdl.DeliveryZoneS = oDeliveryZoneS;

            
            // Dirección Facturación
            var oStreetB = db12.GetValue("StreetB", 0);
            var oStreetNoB = db12.GetValue("StreetNoB", 0);
            var oBlockB = db12.GetValue("BlockB", 0);
            var oCityB = db12.GetValue("CityB", 0);
            var oZipCodeB = db12.GetValue("ZipCodeB", 0);
            var oCountyB = db12.GetValue("CountyB", 0);
            var oStateB = db12.GetValue("StateB", 0);
            var oCountryB = db12.GetValue("CountryB", 0);
            var oBuildingB = db12.GetValue("BuildingB", 0);
            var oAddress2B = db12.GetValue("Address2B", 0);
            var oAddress3B = db12.GetValue("Address3B", 0);
            var oGlbLocNumB = db12.GetValue("GlbLocNumB", 0);
            var oTransportistaB = db12.GetValue("U_ITPS_TransportistaB", 0);
            var oDeliveryZoneB = db12.GetValue("U_ITPS_DeliveryZoneB", 0);

            mdl.StreetB = oStreetB;
            mdl.StreetNoB = oStreetNoB;
            mdl.BlockB = oBlockB;
            mdl.CityB = oCityB;
            mdl.ZipCodeB = oZipCodeB;
            mdl.CountyB = oCountyB;
            mdl.StateB = oStateB;
            mdl.CountryB = oCountryB;
            mdl.BuildingB = oBuildingB;
            mdl.Address2B = oAddress2B;
            mdl.Address3B = oAddress3B;
            mdl.GlbLocNumB = oGlbLocNumB;
            mdl.TransportistaB = oTransportistaB;
            mdl.DeliveryZoneB = oDeliveryZoneB;

            mdl.CardCode = oCardCode.Value;

            string splitPercentageValue = db.GetValue(Constants.SalesOrder_Fields.Head_SplitPercentage, 0).Trim();
            mdl.SplitPercentage = string.IsNullOrEmpty(splitPercentageValue) ? 0m : ConverterService.GetDecimalFromStringSAP(splitPercentageValue);


            string importedPercentageValue = db.GetValue(Constants.SalesOrder_Fields.Head_ImportedPercentage, 0).Trim();
            mdl.ImportedPercentage = string.IsNullOrEmpty(importedPercentageValue) ? 0m : ConverterService.GetDecimalFromStringSAP(importedPercentageValue);

            mdl.RelatedOrder = oRelatedOrder.Value;

            string assignedEntityValue = db.GetValue(Constants.SalesOrder_Fields.Head_AssignedEntity, 0).Trim();
            mdl.AssignedEntity = string.IsNullOrEmpty(assignedEntityValue) ? string.Empty : assignedEntityValue;

            string firstLineGlobalAgree = string.Empty;
            if (oMtx.RowCount >= 1)
            {
                var oFirstLineAgrNo = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_AgrNo, 1);
                firstLineGlobalAgree = oFirstLineAgrNo.Value?.Trim();
            }

            string globalAgree = db.GetValue(Constants.SalesOrder_Fields.Head_GlobalAgree, 0).Trim();
            mdl.GlobalAgreement = int.TryParse(globalAgree, out int agreement) ? agreement : 0;
            
            mdl.AgreementPriceList = db.GetValue(Constants.SalesOrder_Fields.Head_AgreementPriceList, 0).Trim();

            mdl.DocDate = oDocDate.Value;
            mdl.DocDueDate = oDocDueDate.Value;
            mdl.TaxDate = oTaxDate.Value;
            mdl.Comments = oComments.Value;
            mdl.TotalDiscountPercent = decimal.Parse(oDiscPrcnt.Value, CultureInfo.InvariantCulture);
            mdl.CategoryClient = oCategoryClient.Value;
            mdl.PaymentGroupCode = Convert.ToInt32(oPaymentGroupCode.Value);
            mdl.Address2 = oAddress2.Value;
            mdl.ShipToCode = oShipToCode.Value;
            mdl.Address = oAddress.Value;
            mdl.PayToCode = oPayToCode.Value;

            string docEntry = db.GetValue(Constants.SalesOrder_Fields.Head_DocEntry, 0);

            if (!string.IsNullOrEmpty(docEntry))
            {
                mdl.DocEntry = Convert.ToInt32(docEntry);
            }

            

            // LINES
            for (int i = 1; i <= oMtx.RowCount; i++)
            {
                var oItemCode = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_ItemCode, i);
                var sItemCode = oItemCode.Value;
                if (string.IsNullOrWhiteSpace(sItemCode))
                    continue;

                var oQuantity = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_Quantity, i);
                var oUomCode = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_UomCode, i);
                var oUomEntry = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_UomEntry, i);
                var oWhsCode = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_WhsCode, i);
                var oDiscount = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_Discount, i);
                var oTaxCode = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_TaxCode, i);
                var oPrice = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_Price, i);
                var oLineNum = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_LineNum, i);
                var oLineId = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_LineId, i);
                var oLineStatus = (ComboBox)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_LineStatus, i);
                var oGlblAgrmt = (EditText)oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_AgrNo, i);

                var mdlLine = new SalesOrderLineModel();

                mdlLine.ItemCode = sItemCode;
                mdlLine.Quantity = ConverterService.GetDecimalFromStringSAP(oQuantity.Value);
                mdlLine.UomCode = oUomCode.Value;
                mdlLine.UomEntry = oUomEntry.Value;

                mdlLine.WhsCode = oWhsCode.Value;
                mdlLine.Discount = ConverterService.GetDecimalFromStringSAP(oDiscount.Value);
                mdlLine.UnitPrice = ConverterService.GetDecimalFromCurrencyStringSAP(oPrice.Value);
                mdlLine.TaxCode = oTaxCode.Value;
                mdlLine.LineNum = Convert.ToInt32(oLineNum.Value);
                mdlLine.LineId = Convert.ToInt32(oLineId.Value);
                mdlLine.LineStatus = oLineStatus.Value;

                string glblAgrmtValue = oGlblAgrmt.Value;
                mdlLine.AgrNo = string.IsNullOrEmpty(glblAgrmtValue) ? string.Empty : glblAgrmtValue;

                mdl.Lines.Add(mdlLine);

            }

            return mdl;
        }

        
    }
}
