using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Configuration;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.SalesInvoice
{
    public partial class SalesInvoiceFrm
    {
        /// <summary>
        /// Detecta si la factura es de importados y, de ser así, genera el asiento
        /// contable en empresa B: Débito SN / Crédito cuenta de ingresos B.
        /// </summary>
        private int ProcessImportJournalEntry(int invoiceDocEntry)
        {
            (int docEntry, decimal importedPerc) = GetImportOrderDocEntryAndImportedPercFromInvoice(invoiceDocEntry);
            if (docEntry == -1 || importedPerc == 100m) return -1;

            var (cardCode, totalDiscount) = GetInvoiceImportDiscountData(invoiceDocEntry, importedPerc);

            if (string.IsNullOrEmpty(cardCode) || totalDiscount <= 0) return -1;

            string incomeAccount = AppConfig.Get(Constants.ConfigProps.ImportIncomeAccountB);
            if (string.IsNullOrEmpty(incomeAccount))
                throw new Exception($"Factura de Importados: no se encontró la configuración '{Constants.ConfigProps.ImportIncomeAccountB}' en @CONFIGS_DEVELOPMENT.");

            return CreateImportJournalEntry(cardCode, totalDiscount, incomeAccount, invoiceDocEntry);
        }

        /// <summary>
        /// Crea el asiento contable para la deuda de empresa B:
        ///   Línea 1 – Débito al SN (cliente)
        ///   Línea 2 – Crédito a cuenta de ingresos empresa B
        /// </summary>
        private int CreateImportJournalEntry(string cardCode, double amount, string incomeAccount, int invoiceDocEntry)
        {
            if(!ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.StartTransaction();

            JournalEntries je = null;
            try
            {
                je = (JournalEntries)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oJournalEntries);

                je.Memo = $"Deuda empresa B - Factura Importados #{invoiceDocEntry}";
                je.ReferenceDate = DateTime.Today;
                je.TaxDate = DateTime.Now;
                je.DueDate = DateTime.Now;
                je.UserFields.Fields.Item("U_Tipo").Value = "PRESUPUESTO";
                je.UserFields.Fields.Item("U_ITPS_RelatedInvoice").Value = invoiceDocEntry.ToString();

                // Línea 1: Débito al SN
                je.Lines.ShortName = cardCode;
                je.Lines.Debit = amount;
                je.Lines.Add();

                // Línea 2: Crédito a cuenta de ingresos empresa B
                je.Lines.AccountCode = incomeAccount;
                je.Lines.Credit = amount;
                je.Lines.Add();

                if (je.Add() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al crear el asiento de importados. {errCode} - {errMsg}");
                }

                if (ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_Commit);
                
                string newKey = ConnectionSDK.DIAPI.GetNewObjectKey();
                return int.TryParse(newKey, out int transId) ? transId : -1;
            }
            catch
            {
                if (ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_RollBack);
                throw;
            }
            finally
            {

                if (je != null) Marshal.ReleaseComObject(je);
            }
        }

        /// <summary>
        /// Reintenta generar el asiento de importados de la factura abierta,
        /// solo si no existe ya un asiento asociado por OJDT."U_ITPS_RelatedInvoice".
        /// </summary>
        private void RetryImportJournalEntry(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    NotificationService.Warn("Guarde o descarte los cambios de la factura antes de generar el asiento.");
                    return;
                }

                var data = GetDataFromFormInvoice(oForm);

                if (data.DocEntry <= 0)
                {
                    NotificationService.Warn("No se pudo determinar la factura abierta.");
                    return;
                }

                if (data.Canceled != "N")
                {
                    NotificationService.Warn("La factura está cancelada, no se genera el asiento.");
                    return;
                }

                (int importDocEntry, decimal importedPerc) = GetImportOrderDocEntryAndImportedPercFromInvoice(data.DocEntry);
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

                int existingTransId = GetJournalEntryForInvoice(data.DocEntry);
                if (existingTransId != -1)
                {
                    NotificationService.Warn($"La factura ya tiene el asiento N° {existingTransId} asociado. No se genera uno nuevo.");
                    ConnectionSDK.UIAPI.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, null, existingTransId.ToString());
                    return;
                }

                string msg = "La factura no tiene asiento de importados asociado. ¿Desea generarlo?";
                if (ConnectionSDK.UIAPI.MessageBox(msg, 2, "Confirmar", "Cancelar") != 1) return;

                int transId = ProcessImportJournalEntry(data.DocEntry);
                if (transId == -1)
                {
                    NotificationService.Error("No se generó el asiento: no se encontró el cliente o el importe de descuento de importados es cero.");
                    return;
                }

                LinkJournalEntryToInvoice(data.DocEntry, transId);

                // Refresca la factura en pantalla para que no quede desactualizada respecto de la DI API.
                try { ConnectionSDK.UIAPI.ActivateMenuItem(Constants.MenusUID.RowsRefresh); } catch { }

                NotificationService.Success($"Asiento N° {transId} generado correctamente.");
                ConnectionSDK.UIAPI.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, null, transId.ToString());
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }

        /// <summary>
        /// Agrega el asiento como documento referenciado en la factura.
        /// </summary>
        private void LinkJournalEntryToInvoice(int invoiceDocEntry, int transId)
        {
            Documents oInvoice = null;
            try
            {
                oInvoice = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oInvoices);
                if (!oInvoice.GetByKey(invoiceDocEntry))
                    throw new Exception($"Se creó el asiento N° {transId}, pero no se pudo leer la factura {invoiceDocEntry} para referenciarlo.");

                oInvoice.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_JournalEntry;
                oInvoice.DocumentReferences.ReferencedDocEntry = transId;
                oInvoice.DocumentReferences.Add();

                if (oInvoice.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Se creó el asiento N° {transId}, pero hubo un error al referenciarlo en la factura. {errCode} - {errMsg}");
                }
            }
            finally
            {
                if (oInvoice != null) Marshal.ReleaseComObject(oInvoice);
            }
        }
    }
}
