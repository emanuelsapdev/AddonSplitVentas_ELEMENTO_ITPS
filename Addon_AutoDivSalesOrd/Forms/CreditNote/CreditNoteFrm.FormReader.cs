using Addon_AutoDivSalesOrd.Models;
using System.Globalization;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.CreditNote
{
    public partial class CreditNoteFrm
    {
        /// <summary>
        /// Lee de la cabecera (ORIN) de la nota de crédito abierta los datos necesarios
        /// para evaluar el asiento de importados.
        /// </summary>
        private CreditNoteFormModel GetDataFromFormCreditNote(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DBDataSource oDS = null;
            try
            {
                oDS = oForm.DataSources.DBDataSources.Item("ORIN");

                int.TryParse(oDS.GetValue("DocEntry", 0).Trim(), out int docEntry);
                decimal.TryParse(oDS.GetValue("U_ITPS_ImportedPercentage", 0).Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal importedPerc);

                return new CreditNoteFormModel
                {
                    DocEntry = docEntry,
                    Canceled = oDS.GetValue("CANCELED", 0).Trim(),
                    Importado = oDS.GetValue("U_Importado", 0).Trim(),
                    ImportedPercentage = importedPerc
                };
            }
            finally
            {
                if (oDS != null) Marshal.ReleaseComObject(oDS);
            }
        }
    }
}
