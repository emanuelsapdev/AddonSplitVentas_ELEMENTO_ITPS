using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Configuration;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.CreditNote
{
    public partial class CreditNoteFrm
    {
        /// <summary>
        /// Detecta si la nota de crédito es de importados y, de ser así, genera el asiento
        /// contable que revierte la deuda de empresa B: Débito cuenta de ingresos B / Crédito SN.
        /// </summary>
        private int ProcessImportJournalEntry(int creditNoteDocEntry)
        {
            (int docEntry, decimal importedPerc) = GetImportedDocEntryAndImportedPercFromCreditNote(creditNoteDocEntry);
            if (docEntry == -1 || importedPerc == 100m) return -1;

            var (cardCode, totalDiscount) = GetCreditNoteImportDiscountData(creditNoteDocEntry, importedPerc);

            if (string.IsNullOrEmpty(cardCode) || totalDiscount <= 0) return -1;

            string incomeAccount = AppConfig.Get(Constants.ConfigProps.ImportIncomeAccountB);
            if (string.IsNullOrEmpty(incomeAccount))
                throw new Exception($"Nota de Crédito de Importados: no se encontró la configuración '{Constants.ConfigProps.ImportIncomeAccountB}' en @CONFIGS_DEVELOPMENT.");

            return CreateImportJournalEntry(cardCode, totalDiscount, incomeAccount, creditNoteDocEntry);
        }

        /// <summary>
        /// Crea el asiento contable que revierte la deuda de empresa B (inverso al de la factura):
        ///   Línea 1 – Débito a cuenta de ingresos empresa B
        ///   Línea 2 – Crédito al SN (cliente)
        /// </summary>
        private int CreateImportJournalEntry(string cardCode, double amount, string incomeAccount, int creditNoteDocEntry)
        {
            if (!ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.StartTransaction();

            JournalEntries je = null;
            try
            {
                je = (JournalEntries)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oJournalEntries);

                je.Memo = $"Reversión deuda empresa B - NC Importados #{creditNoteDocEntry}";
                je.ReferenceDate = DateTime.Today;
                je.TaxDate = DateTime.Now;
                je.DueDate = DateTime.Now;
                je.UserFields.Fields.Item("U_Tipo").Value = "PRESUPUESTO";
                je.UserFields.Fields.Item("U_ITPS_RelatedCreditNote").Value = creditNoteDocEntry.ToString();

                // Línea 1: Débito a cuenta de ingresos empresa B
                je.Lines.AccountCode = incomeAccount;
                je.Lines.Debit = amount;
                je.Lines.Add();

                // Línea 2: Crédito al SN
                je.Lines.ShortName = cardCode;
                je.Lines.Credit = amount;
                je.Lines.Add();

                if (je.Add() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al crear el asiento de importados de la nota de crédito. {errCode} - {errMsg}");
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
        /// Al crear la nota de crédito, genera el asiento de importados y lo referencia en ella.
        /// </summary>
        private void GenerateImportJournalEntryOnCreditNoteAdd(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                var data = GetDataFromFormCreditNote(oForm);
                if (data.Canceled != "N" || data.DocEntry <= 0) return;

                int transId = CreateAndLinkImportJournalEntry(data.DocEntry);
                if (transId != -1)
                    ConnectionSDK.UIAPI.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, null, transId.ToString());
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }

        /// <summary>
        /// Genera el asiento de importados de la nota de crédito y, si se creó, lo referencia en ella.
        /// </summary>
        /// <returns>TransId del asiento creado, o -1 si la nota de crédito no corresponde.</returns>
        private int CreateAndLinkImportJournalEntry(int creditNoteDocEntry)
        {
            int transId = ProcessImportJournalEntry(creditNoteDocEntry);
            if (transId == -1) return -1;

            LinkJournalEntryToCreditNote(creditNoteDocEntry, transId);
            return transId;
        }

        /// <summary>
        /// Reintenta generar el asiento de importados de la nota de crédito abierta,
        /// solo si no existe ya un asiento asociado por OJDT."U_ITPS_RelatedCreditNote".
        /// </summary>
        private void RetryImportJournalEntry(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    NotificationService.Warn("Guarde o descarte los cambios de la nota de crédito antes de generar el asiento.");
                    return;
                }

                var data = GetDataFromFormCreditNote(oForm);

                if (data.DocEntry <= 0)
                {
                    NotificationService.Warn("No se pudo determinar la nota de crédito abierta.");
                    return;
                }

                if (data.Canceled != "N")
                {
                    NotificationService.Warn("La nota de crédito está cancelada, no se genera el asiento.");
                    return;
                }

                (int importDocEntry, decimal importedPerc) = GetImportedDocEntryAndImportedPercFromCreditNote(data.DocEntry);
                if (importDocEntry == -1)
                {
                    NotificationService.Warn("La nota de crédito no es de importados (U_Importado distinto de 'Y').");
                    return;
                }

                if (importedPerc != RetryJournalImportedPerc)
                {
                    NotificationService.Warn($"El porcentaje de importados de la nota de crédito es {importedPerc:0.##}%. Solo se genera asiento para {RetryJournalImportedPerc:0.##}%.");
                    return;
                }

                int existingTransId = GetJournalEntryForCreditNote(data.DocEntry);
                if (existingTransId != -1)
                {
                    NotificationService.Warn($"La nota de crédito ya tiene el asiento N° {existingTransId} asociado. No se genera uno nuevo.");
                    ConnectionSDK.UIAPI.OpenForm(SAPbouiCOM.BoFormObjectEnum.fo_JournalPosting, null, existingTransId.ToString());
                    return;
                }

                string msg = "La nota de crédito no tiene asiento de importados asociado. ¿Desea generarlo?";
                if (ConnectionSDK.UIAPI.MessageBox(msg, 2, "Confirmar", "Cancelar") != 1) return;

                int transId = CreateAndLinkImportJournalEntry(data.DocEntry);
                if (transId == -1)
                {
                    NotificationService.Error("No se generó el asiento: no se encontró el cliente o el importe de descuento de importados es cero.");
                    return;
                }

                // Refresca la nota de crédito en pantalla para que no quede desactualizada respecto de la DI API.
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
        /// Agrega el asiento como documento referenciado en la nota de crédito.
        /// </summary>
        private void LinkJournalEntryToCreditNote(int creditNoteDocEntry, int transId)
        {
            Documents oCreditNote = null;
            try
            {
                oCreditNote = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oCreditNotes);
                if (!oCreditNote.GetByKey(creditNoteDocEntry))
                    throw new Exception($"Se creó el asiento N° {transId}, pero no se pudo leer la nota de crédito {creditNoteDocEntry} para referenciarlo.");

                oCreditNote.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_JournalEntry;
                oCreditNote.DocumentReferences.ReferencedDocEntry = transId;
                oCreditNote.DocumentReferences.Add();

                if (oCreditNote.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Se creó el asiento N° {transId}, pero hubo un error al referenciarlo en la nota de crédito. {errCode} - {errMsg}");
                }
            }
            finally
            {
                if (oCreditNote != null) Marshal.ReleaseComObject(oCreditNote);
            }
        }
    }
}
