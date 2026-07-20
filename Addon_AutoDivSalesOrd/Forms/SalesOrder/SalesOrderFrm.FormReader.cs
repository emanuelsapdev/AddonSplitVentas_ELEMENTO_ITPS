using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
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
            var oMtx = (Matrix)oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Matrix).Specific;

            mdl.CardCode = oCardCode.Value;

            string splitPercentageValue = db.GetValue(Constants.SalesOrder_Fields.Head_SplitPercentage, 0).Trim();
            mdl.SplitPercentage = string.IsNullOrEmpty(splitPercentageValue) ? 0m : ConverterService.GetDecimalFromStringSAP(splitPercentageValue);


            string importedPercentageValue = db.GetValue(Constants.SalesOrder_Fields.Head_ImportedPercentage, 0).Trim();
            mdl.ImportedPercentage = string.IsNullOrEmpty(importedPercentageValue) ? 0m : ConverterService.GetDecimalFromStringSAP(importedPercentageValue);

            mdl.RelatedOrder = oRelatedOrder.Value;

            string assignedEntityValue = db.GetValue(Constants.SalesOrder_Fields.Head_AssignedEntity, 0).Trim();
            mdl.AssignedEntity = string.IsNullOrEmpty(assignedEntityValue) ? string.Empty : assignedEntityValue;

            string globalAgree = db.GetValue(Constants.SalesOrder_Fields.Head_GlobalAgree, 0).Trim();
            mdl.GlobalAgreement = int.TryParse(globalAgree, out int agreement) ? agreement : 0;

            mdl.DocDate = oDocDate.Value;
            mdl.DocDueDate = oDocDueDate.Value;
            mdl.TaxDate = oTaxDate.Value;
            mdl.Comments = oComments.Value;
            mdl.TotalDiscountPercent = decimal.Parse(oDiscPrcnt.Value, CultureInfo.InvariantCulture);
            mdl.CategoryClient = oCategoryClient.Value;
            mdl.PaymentGroupCode = Convert.ToInt32(oPaymentGroupCode.Value);

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
