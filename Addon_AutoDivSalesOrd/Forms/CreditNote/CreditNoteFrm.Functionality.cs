using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using SAPbouiCOM;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.CreditNote
{
    public partial class CreditNoteFrm
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
        /// Una nota de crédito lleva asiento de importados si U_Importado = 'Y',
        /// U_ITPS_ImportedPercentage = 50 y no está cancelada.
        /// </summary>
        private bool IsImportedCreditNoteForJournal(CreditNoteFormModel data)
        {
            return data.Importado == "Y"
                && data.ImportedPercentage == RetryJournalImportedPerc
                && data.Canceled == "N";
        }

        /// <summary>
        /// Habilita el botón "Generar Asiento Imp." solo si la nota de crédito cargada es de importados
        /// (U_Importado = 'Y', U_ITPS_ImportedPercentage = 50, no cancelada) y todavía no tiene
        /// un asiento asociado por OJDT."U_ITPS_RelatedCreditNote". Se evalúa en cada carga de
        /// la nota de crédito (apertura, navegación, refresco) y luego de actualizarla.
        /// </summary>
        private void UpdateBtnRetryJournalState(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                var data = GetDataFromFormCreditNote(oForm);

                bool enabled = data.DocEntry > 0
                    && IsImportedCreditNoteForJournal(data)
                    && GetJournalEntryForCreditNote(data.DocEntry) == -1;

                ToggleEnableBtnRetryJournal(oForm, enabled);
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }

        /// <summary>
        /// Al seleccionar el porcentaje de importados, aplica ese porcentaje como descuento
        /// en las líneas cuyo artículo tiene la propiedad de importado (OITM.QryGroup1 = 'Y').
        /// Con 100% el descuento de esas líneas vuelve a 0.
        /// </summary>
        private void ApplyImportedPercDiscountToLines(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbobsCOM.Recordset oRec = null;

            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                oRec = (SAPbobsCOM.Recordset)ConnectionSDK.DIAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbouiCOM.ComboBox etImportedPerc = oForm.Items.Item(Constants.CreditNote_FieldsUIDs.Head_ImportedPercentage).Specific;
                string vImportedPerc = etImportedPerc.Value;
                decimal.TryParse(vImportedPerc, out decimal importedPerc);

                Matrix oMtx = oForm.Items.Item(Constants.CreditNote_FieldsUIDs.Head_Matrix).Specific;

                oForm.Freeze(true);

                for (int row = 1; row <= oMtx.RowCount; row++)
                {
                    SAPbouiCOM.EditText oDiscount = null;
                    try
                    {
                        oDiscount = oMtx.Columns.Item(Constants.CreditNote_FieldsUIDs.Det_Discount).Cells.Item(row).Specific;
                        string itemCode = (oMtx.GetCellSpecific(Constants.CreditNote_FieldsUIDs.Det_ItemCode, row)).Value;

                        string q = $@"SELECT 1 FROM OITM WHERE ""ItemCode"" = '{itemCode}' AND ""{Constants.ItemImport.OitmImportProperty}"" = 'Y'";
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

                oForm.ActiveItem = Constants.CreditNote_FieldsUIDs.Head_ImportedPercentage;
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
}
