using Addon_AutoDivSalesOrd.Common;
using SAPbouiCOM;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.SalesInvoice
{
    public partial class SalesInvoiceFrm
    {
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

                if (TryGetItem(oForm, Constants.Invoice_FieldsUIDs.Head_BtnRetryJournal) != null) return;

                SAPbouiCOM.Item oItemBtnCancel = oForm.Items.Item(Constants.Invoice_FieldsUIDs.Head_BtnCancel);
                SAPbouiCOM.Item oItemBtn = oForm.Items.Add(Constants.Invoice_FieldsUIDs.Head_BtnRetryJournal, BoFormItemTypes.it_BUTTON);

                oItemBtn.Top = oItemBtnCancel.Top;
                oItemBtn.Left = oItemBtnCancel.Left + oItemBtnCancel.Width + 5;
                oItemBtn.Width = oItemBtnCancel.Width * 2;
                oItemBtn.Height = oItemBtnCancel.Height;
                oItemBtn.AffectsFormMode = false;
                oItemBtn.LinkTo = Constants.Invoice_FieldsUIDs.Head_BtnCancel;

                oItemBtn.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)BoAutoFormMode.afm_Add, BoModeVisualBehavior.mvb_False);
                oItemBtn.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)BoAutoFormMode.afm_Find, BoModeVisualBehavior.mvb_False);
                oItemBtn.SetAutoManagedAttribute(BoAutoManagedAttr.ama_Editable, (int)BoAutoFormMode.afm_Ok, BoModeVisualBehavior.mvb_True);

                ((SAPbouiCOM.Button)oItemBtn.Specific).Caption = "Generar Asiento Imp.";
            }
            finally
            {
                if (oForm != null) Marshal.ReleaseComObject(oForm);
            }
        }

        private void ToggleEnableBtnRetryJournal(SAPbouiCOM.Form oForm, bool enabled)
        {
            SAPbouiCOM.Item oItemBtn = TryGetItem(oForm, Constants.Invoice_FieldsUIDs.Head_BtnRetryJournal);
            if (oItemBtn == null) return;

            oItemBtn.Enabled = enabled;
        }
    }
}
