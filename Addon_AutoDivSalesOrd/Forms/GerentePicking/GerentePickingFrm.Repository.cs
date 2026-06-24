using Addon_AutoDivSalesOrd.Addons.Tools;
using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Forms.GerentePicking
{
    public partial class GerentePickingFrm
    {
        public decimal GetAvailableStock(string itemCode, string whsCode)
        {
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                oRec = (SAPbobsCOM.Recordset)ConnectionSDK.DIAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                
                string q = string.Format(@"SELECT ""AvailableStock_Unidades_Base"" FROM {0} WHERE ""ItemCode"" = '{1}' AND ""WhsCode"" = '{2}';", Constants.DbViews.StockSplitVta, itemCode, whsCode);

                oRec.DoQuery(q);

                if (oRec.EoF)
                {
                    return 0m;
                }

                return Convert.ToDecimal(oRec.Fields.Item(0).Value);

            }
            catch (Exception ex)
            {
                NotificationService.Error(ex.Message);
                return 0m;
            }
            finally
            {
                if (oRec != null)
                {
                    MarshalGC.ReleaseComObject(oRec);
                }
            }
        }

       

        public decimal GetQtyDozenPerPackage(string itemCode)
        {
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                oRec = (SAPbobsCOM.Recordset)ConnectionSDK.DIAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string q = string.Format(@"
                    SELECT TO_TINYINT(REPLACE(IFNULL(T1.""Name"", '1'), ',', '.')) AS ""QtyXPack""
                    FROM OITM T0
                    LEFT JOIN ""@ITPS_DOC_BULTO"" T1 ON T1.""Code"" = T0.U_DOC_BULTO
                    WHERE T0.""ItemCode"" = '{0}'", itemCode);

                oRec.DoQuery(q);

                if (oRec.EoF)
                    return 0m;

                return Convert.ToDecimal(oRec.Fields.Item(0).Value);
            }
            catch (Exception ex)
            {
                NotificationService.Error(ex.Message);
                return 0m;
            }
            finally
            {
                if (oRec != null) MarshalGC.ReleaseComObject(oRec);
            }
        }
    }
}
