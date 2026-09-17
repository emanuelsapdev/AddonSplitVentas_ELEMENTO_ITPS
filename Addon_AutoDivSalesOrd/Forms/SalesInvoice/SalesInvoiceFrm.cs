using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Forms.SalesInvoice
{
    /// <summary>
    /// Manejador de eventos del formulario de Factura de Venta (FormType 133).
    /// Detecta facturas originadas en Órdenes de Artículos Importados y genera
    /// el asiento contable correspondiente en empresa B.
    /// </summary>
    public partial class SalesInvoiceFrm : IFormEventHandler
    {
        public const string FormType = Constants.FormTypes.SalesInvoice;

        #region Implementación de IFormEventHandler

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // 
            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ActionSuccess
                     && pVal.ItemUID == Constants.Invoice_FieldsUIDs.Head_ImportedPercentage

                     )
            {
                SAPbouiCOM.Form oForm = null;
                SAPbobsCOM.Recordset oRec = null;

                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    oRec = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                    SAPbouiCOM.ComboBox etImportedPerc = oForm.Items.Item(Constants.Invoice_FieldsUIDs.Head_ImportedPercentage).Specific;
                    string vImportedPerc = etImportedPerc.Value;
                    decimal.TryParse(vImportedPerc, out decimal importedPerc);

                    //if (importedPerc == 100m) return;

                    Matrix oMtx = oForm.Items.Item(Constants.Invoice_FieldsUIDs.Head_Matrix).Specific;

                    oForm.Freeze(true);

                    for (int row = 1; row <= oMtx.RowCount; row++)
                    {
                        SAPbouiCOM.EditText oDiscount = null;
                        try
                        {

                            oDiscount = oMtx.Columns.Item(Constants.Invoice_FieldsUIDs.Det_Discount).Cells.Item(row).Specific;
                            string itemCode = (oMtx.GetCellSpecific(Constants.Invoice_FieldsUIDs.Det_ItemCode, row)).Value;

                            string q = $@"SELECT 1 FROM OITM WHERE ""ItemCode"" = '{itemCode}' AND ""QryGroup1"" = 'Y'";
                            oRec.DoQuery(q);
                            if (oRec.RecordCount == 0) continue;


                            decimal.TryParse(oDiscount.Value.Replace(".", ","), out decimal discountSap);
                            if (string.IsNullOrEmpty(itemCode) || discountSap == importedPerc) continue;
                            oDiscount.Value = importedPerc != 100m ? vImportedPerc.Replace(",", ".") : "0.00";
                        }
                        finally
                        {
                            if (oDiscount != null) Marshal.ReleaseComObject(oDiscount);
                        }
                    }

                    oForm.ActiveItem = Constants.Invoice_FieldsUIDs.Head_ImportedPercentage;
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null)
                    {
                        oForm.Freeze(false);
                        Marshal.ReleaseComObject(oForm);
                    }

                    if (oRec != null) Marshal.ReleaseComObject(oRec);

                }
            }
        }

        public void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (boi.EventType == BoEventTypes.et_FORM_DATA_ADD && boi.ActionSuccess)
            {
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.DBDataSource oDS = null;
                SAPbobsCOM.Documents oInvoice = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(boi.FormUID);
                    oDS = oForm.DataSources.DBDataSources.Item("OINV");
                    oInvoice = ConnectionSDK.DIAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

                    string rawDocEntry = oDS.GetValue("DocEntry", 0);
                    string CANCELED = oDS.GetValue("CANCELED", 0);
                    if (CANCELED != "N") return;
                    
                    if (!int.TryParse(rawDocEntry, out int docEntry) || docEntry <= 0) return;

                    if (oInvoice.GetByKey(docEntry))
                    {
                        int transId = ProcessImportJournalEntry(docEntry);
                        if(transId != -1)
                        {
                            oInvoice.DocumentReferences.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_JournalEntry;
                            oInvoice.DocumentReferences.ReferencedDocEntry = transId;
                            oInvoice.DocumentReferences.Add();

                            var ret = oInvoice.Update();
                            if (ret != 0)
                            {
                                ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                                throw new Exception($"Error al actualizar la factura de importados con el asiento referenciado. {errCode} - {errMsg}");
                            }
                            string transIdStr = transId.ToString();
                            ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_JournalPosting, null, transIdStr);
                        }

                    }


                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
                finally
                {
                    if (oForm != null)
                        Marshal.ReleaseComObject(oForm);

                    if (oDS != null)
                        Marshal.ReleaseComObject(oDS);
                }

            }

        }

        public void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion
    }
}
