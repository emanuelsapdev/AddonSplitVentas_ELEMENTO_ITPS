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
        /// Habilita el botón "Generar Asiento Imp." solo si la factura cargada es de importados.
        /// </summary>
        private void UpdateBtnRetryJournalState(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                var data = GetDataFromFormInvoice(oForm);
                ToggleEnableBtnRetryJournal(oForm, IsImportedInvoiceForJournal(data));
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }
    }
}
