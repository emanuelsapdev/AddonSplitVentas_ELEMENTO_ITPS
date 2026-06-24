using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Configuration;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Addon_AutoDivSalesOrd.Forms.SalesOrder
{
    public partial class SalesOrderFrm
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

 
        private void CloseAndReopen(SAPbouiCOM.Form oForm, int entryPrimary, int entrySecondary)
        {
            oForm.Mode = BoFormMode.fm_OK_MODE;
            oForm.Close();

            bool bothOpen = entryPrimary > 0 && entrySecondary > 0;

            if (entryPrimary > 0)
            {
                ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, entryPrimary.ToString());

                if (bothOpen)
                {
                    var fPrimary = ConnectionSDK.UIAPI.Forms.ActiveForm;
                    fPrimary.Left = 0;
                    fPrimary.Top = 0;

                    if (entrySecondary > 0)
                    {
                        ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, entrySecondary.ToString());
                        var fSecondary = ConnectionSDK.UIAPI.Forms.ActiveForm;
                        fSecondary.Left = fPrimary.Left + fPrimary.Width + 4;
                        fSecondary.Top = fPrimary.Top;
                    }

                    return;
                }
            }

            if (entrySecondary > 0)
                ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, entrySecondary.ToString());
        }

        private void EnsureNoActiveTransaction()
        {
            if (ConnectionSDK.DIAPI.InTransaction)
                ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_RollBack);
        }

        private void RollbackIfNeeded()
        {
            if (ConnectionSDK.DIAPI.InTransaction)
                ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_RollBack);
        }


    }
}
