using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Configuration;
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
            (int orderDocEntry, decimal importedPerc) = GetImportOrderDocEntryAndImportedPercFromInvoice(invoiceDocEntry);
            if (orderDocEntry == -1 || importedPerc == 100m) return -1;

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
                return -1;
            }
            finally
            {

                if (je != null) Marshal.ReleaseComObject(je);
            }
        }
    }
}
