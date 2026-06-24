using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Services;
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
