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

        private const string BtnRetryJournalUID = "btnRetryJE";
        private const decimal RetryJournalImportedPerc = 50m;

        #region Implementación de IFormEventHandler

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
            {
                AddBtnRetryJournal(FormUID);
                return;
            }

            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess
                     && pVal.ItemUID == BtnRetryJournalUID)
            {
                RetryImportJournalEntry(FormUID);
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

        #region Reintento de asiento de importados

        /// <summary>
        /// Agrega el botón "Generar Asiento Imp." junto al botón Cancelar.
        /// Solo queda habilitado en modo OK (factura existente).
        /// </summary>
        private void AddBtnRetryJournal(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                for (int i = 0; i < oForm.Items.Count; i++)
                    if (oForm.Items.Item(i).UniqueID == BtnRetryJournalUID) return;

                SAPbouiCOM.Item oItemBtnCancel = oForm.Items.Item("2");
                SAPbouiCOM.Item oItemBtn = oForm.Items.Add(BtnRetryJournalUID, BoFormItemTypes.it_BUTTON);

                oItemBtn.Top = oItemBtnCancel.Top;
                oItemBtn.Left = oItemBtnCancel.Left + oItemBtnCancel.Width + 5;
                oItemBtn.Width = oItemBtnCancel.Width * 2;
                oItemBtn.Height = oItemBtnCancel.Height;
                oItemBtn.AffectsFormMode = false;
                oItemBtn.LinkTo = "2";

                oItemBtn.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)BoAutoFormMode.afm_Add, BoModeVisualBehavior.mvb_False);
                oItemBtn.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)BoAutoFormMode.afm_Find, BoModeVisualBehavior.mvb_False);
                oItemBtn.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)BoAutoFormMode.afm_Ok, BoModeVisualBehavior.mvb_True);

                ((SAPbouiCOM.Button)oItemBtn.Specific).Caption = "Generar Asiento Imp.";
            }
            catch (Exception ex)
            {
                NotificationService.Error($"Error al agregar el botón de reintento de asiento. {ex.Message}");
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }

        /// <summary>
        /// Reintenta generar el asiento de importados de la factura abierta,
        /// solo si no existe ya un asiento vigente asociado por OJDT."U_ITPS_RelatedInvoice".
        /// </summary>
        private void RetryImportJournalEntry(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.DBDataSource oDS = null;
            SAPbobsCOM.Documents oInvoice = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                if (oForm.Mode != BoFormMode.fm_OK_MODE)
                {
                    NotificationService.Warn("Guarde o descarte los cambios de la factura antes de generar el asiento.");
                    return;
                }

                oDS = oForm.DataSources.DBDataSources.Item("OINV");

                string rawDocEntry = oDS.GetValue("DocEntry", 0).Trim();
                string CANCELED = oDS.GetValue("CANCELED", 0).Trim();

                if (!int.TryParse(rawDocEntry, out int docEntry) || docEntry <= 0)
                {
                    NotificationService.Warn("No se pudo determinar la factura abierta.");
                    return;
                }

                if (CANCELED != "N")
                {
                    NotificationService.Warn("La factura está cancelada, no se genera el asiento.");
                    return;
                }

                (int importDocEntry, decimal importedPerc) = GetImportOrderDocEntryAndImportedPercFromInvoice(docEntry);
                if (importDocEntry == -1)
                {
                    NotificationService.Warn("La factura no es de importados (U_Importado distinto de 'Y').");
                    return;
                }

                if (importedPerc != RetryJournalImportedPerc)
                {
                    NotificationService.Warn($"El porcentaje de importados de la factura es {importedPerc:0.##}%. Solo se genera asiento para {RetryJournalImportedPerc:0.##}%.");
                    return;
                }

                int existingTransId = GetActiveJournalEntryForInvoice(docEntry);
                if (existingTransId != -1)
                {
                    NotificationService.Warn($"La factura ya tiene el asiento N° {existingTransId} asociado. No se genera uno nuevo.");
                    ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_JournalPosting, null, existingTransId.ToString());
                    return;
                }

                string msg = "La factura no tiene asiento de importados asociado. ¿Desea generarlo?";
                if (ConnectionSDK.UIAPI.MessageBox(msg, 2, "Confirmar", "Cancelar") != 1) return;

                int transId = ProcessImportJournalEntry(docEntry);
                if (transId == -1)
                {
                    NotificationService.Error("No se generó el asiento: no se encontró el cliente o el importe de descuento de importados es cero.");
                    return;
                }

                oInvoice = ConnectionSDK.DIAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                if (!oInvoice.GetByKey(docEntry))
                    throw new Exception($"Se creó el asiento N° {transId}, pero no se pudo leer la factura {docEntry} para referenciarlo.");

                SAPbobsCOM.Document_References oRefs = oInvoice.DocumentReferences;
                if (oRefs.Count > 0)
                {
                    oRefs.SetCurrentLine(oRefs.Count - 1);
                    if (oRefs.ReferencedDocEntry != 0) oRefs.Add();
                }

                oRefs.ReferencedObjectType = SAPbobsCOM.ReferencedObjectTypeEnum.rot_JournalEntry;
                oRefs.ReferencedDocEntry = transId;

                var ret = oInvoice.Update();
                if (ret != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Se creó el asiento N° {transId}, pero hubo un error al referenciarlo en la factura. {errCode} - {errMsg}");
                }

                // Refresca la factura en pantalla para que no quede desactualizada respecto de la DI API.
                try { ConnectionSDK.UIAPI.ActivateMenuItem(Constants.MenusUID.RowsRefresh); } catch { }

                NotificationService.Success($"Asiento N° {transId} generado correctamente.");
                ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_JournalPosting, null, transId.ToString());
            }
            catch (Exception ex)
            {
                NotificationService.Error(ex.Message);
            }
            finally
            {
                if (oInvoice != null) Marshal.ReleaseComObject(oInvoice);
                if (oDS != null) Marshal.ReleaseComObject(oDS);
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }

        #endregion
    }
}
