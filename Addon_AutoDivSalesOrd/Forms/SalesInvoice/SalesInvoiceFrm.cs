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

        private const decimal RetryJournalImportedPerc = 50m;

        #region Implementación de IFormEventHandler

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
            {
                try
                {
                    AddBtnRetryJournal(FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error($"Error al agregar el botón de reintento de asiento. {ex.Message}");
                }
                return;
            }

            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess
                     && pVal.ItemUID == Constants.Invoice_FieldsUIDs.Head_BtnRetryJournal)
            {
                try
                {
                    RetryImportJournalEntry(FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
                return;
            }

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

            if ((boi.EventType == BoEventTypes.et_FORM_DATA_LOAD || boi.EventType == BoEventTypes.et_FORM_DATA_UPDATE)
                     && boi.ActionSuccess)
            {
                try
                {
                    UpdateBtnRetryJournalState(boi.FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error($"Error al actualizar el estado del botón de asiento de importados. {ex.Message}");
                }
                return;
            }

            if (boi.EventType == BoEventTypes.et_FORM_DATA_ADD && boi.ActionSuccess)
            {
                try
                {
                    GenerateImportJournalEntryOnInvoiceAdd(boi.FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
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
