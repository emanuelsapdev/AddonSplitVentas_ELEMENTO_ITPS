using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using SAPbouiCOM;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.SalesInvoice
{
    public partial class SalesInvoiceFrm
    {
        private Item TryGetItem(SAPbouiCOM.Form form, string itemUid)
        {
            try
            {
                return form.Items.Item(itemUid);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Una factura lleva asiento de importados si U_Importado = 'Y',
        /// U_ITPS_ImportedPercentage = 50 y no está cancelada.
        /// </summary>
        private bool IsImportedInvoiceForJournal(SalesInvoiceFormModel data)
        {
            return data.Importado == "Y"
                && data.ImportedPercentage == RetryJournalImportedPerc
                && data.Canceled == "N";
        }

        /// <summary>
        /// Habilita el botón "Generar Asiento Imp." solo si la factura cargada es de importados
        /// (U_Importado = 'Y', U_ITPS_ImportedPercentage = 50, no cancelada) y todavía no tiene
        /// un asiento asociado por OJDT."U_ITPS_RelatedInvoice". Se evalúa en cada carga de
        /// la factura (apertura, navegación, refresco) y luego de actualizarla.
        /// </summary>
        private void UpdateBtnRetryJournalState(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                var data = GetDataFromFormInvoice(oForm);

                bool enabled = data.DocEntry > 0
                    && IsImportedInvoiceForJournal(data)
                    && GetJournalEntryForInvoice(data.DocEntry) == -1;

                ToggleEnableBtnRetryJournal(oForm, enabled);
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }
    }
}
