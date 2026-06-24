using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Forms.SalesQuote
{
    partial class SalesQuoteFrm : IFormEventHandler
    {
        public const string FormType = "149";
        public const string MenuUID = "2049";

        #region Implementación de IFormEventHandler

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.ActionSuccess)
            {
                SAPbouiCOM.Form oForm = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    oForm.Freeze(true);
                    SAPbouiCOM.ComboBox oCbTipo = oForm.Items.Item(FieldsUIDs.Head_AssignedEntity).Specific;
                    oCbTipo.Select(Constants.FixedValues.EntityA);

                    EditText oDocDueDate = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_DocDueDate).Specific;
                    oDocDueDate.Value = DateTime.Now.AddDays(7).ToString("yyyyMMdd");

                    SAPbouiCOM.EditText oEtCardCode = oForm.Items.Item(FieldsUIDs.Head_CardCode).Specific;
                    oEtCardCode.Item.Click();

                    oCbTipo.Item.Enabled = false;
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
                   
                }

            }

        }

        public void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (boi.EventType == BoEventTypes.et_FORM_LOAD && boi.ActionSuccess)
            {
                SAPbouiCOM.Form oForm = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(boi.FormUID);
                    
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
                }

            }

        }



        public void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = null;

            try
            {

                switch (pVal.MenuUID)
                {
                    case MenuUID:
                    case Constants.MenusUID.ModeAdd:
                    case Constants.MenusUID.Duplicate:
                    case Constants.MenusUID.RowInitial:
                    case Constants.MenusUID.RowBefore:
                    case Constants.MenusUID.RowAfter:
                    case Constants.MenusUID.RowEnd:
                    case Constants.MenusUID.RowsRefresh:
                    case Constants.MenusUID.Cancel:
                    case Constants.MenusUID.Close:
                        break;
                }
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
            }

        }

        #endregion
    }
}
