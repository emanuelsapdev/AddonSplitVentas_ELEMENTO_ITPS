using Addon_AutoDivSalesOrd.Addons.Tools;
using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Tools;
using SAPbouiCOM;
using System;

namespace Addon_AutoDivSalesOrd.Forms.GerentePicking
{
    public partial class GerentePickingFrm
    {
        private static SAPbouiCOM.Button AddBtnProcess(string FormUID)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                SAPbouiCOM.Button oBtnItem = oForm.Items.Add(Constants.GerentePicking_FieldsUIDs.Head_BtnProcess, BoFormItemTypes.it_BUTTON).Specific;

                SAPbouiCOM.Item oItemBtnCancel = oForm.Items.Item(Constants.GerentePicking_FieldsUIDs.Head_BtnCancel);

                oBtnItem.Item.Top = oItemBtnCancel.Top;
                oBtnItem.Item.Left = oItemBtnCancel.Left + oItemBtnCancel.Width + 10;
                oBtnItem.Item.Width = oItemBtnCancel.Width * 2;
                oBtnItem.Item.Height = oItemBtnCancel.Height;

                oBtnItem.Caption = "Procesar Split Stock";

                return oBtnItem;
            }
            catch { throw; }
            finally
            {
                if (oForm != null) MarshalGC.ReleaseComObject(oForm);
            }
        }

        private static void ToggleEnableBtnProcess(bool enabled)
        {
            SAPbouiCOM.Form oForm = null;
            try
            {
                oForm = ConnectionSDK.UIAPI.Forms.ActiveForm;
                
                try
                {
                    SAPbouiCOM.Button oBtnItem = oForm.Items.Item(Constants.GerentePicking_FieldsUIDs.Head_BtnProcess).Specific;
                    oBtnItem.Item.Enabled = enabled;
                    
                } catch
                {
                    

                } 
            }
            finally
            {
                if (oForm != null) MarshalGC.ReleaseComObject(oForm);
            }
        }

        public static void ClearAllToReleaseValues(Matrix oMtx)
        {
            for (int r = 1; r <= oMtx.RowCount; r++)
            {
                oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_ToRelease, r).Value = "";
            }
        }
    }
}
