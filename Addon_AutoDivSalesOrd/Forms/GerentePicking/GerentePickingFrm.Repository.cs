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

        public decimal GetItemsPerUnit(string itemCode, string uomCode)
        {
            SAPbobsCOM.Recordset oRec = null;
            try
            {
                oRec = (SAPbobsCOM.Recordset)ConnectionSDK.DIAPI.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string article = (itemCode ?? string.Empty).Replace("'", "''");
                string unitCode = (uomCode ?? string.Empty).Replace("'", "''");

                string q = $@"
                                SELECT
                                    T0.""ItemCode"",
                                    T1.""UgpEntry"",
                                    T1.""UgpCode"",
                                    T1.""UgpName"",
                                    T2.""UomCode""  AS ""UnidadBase"",
                                    T2.""UomName""  AS ""NombreUnidadBase"",
                                    T4.""UomCode""  AS ""UnidadConsultada"",
                                    CASE
                                        WHEN T4.""UomEntry"" = T1.""BaseUom"" THEN 1
                                        ELSE T3.""BaseQty"" / T3.""AltQty""
                                    END AS ""FactorConversionAUnidadBase""
                                FROM ""OITM"" T0
                                INNER JOIN ""OUGP"" T1 ON T1.""UgpEntry"" = T0.""UgpEntry""
                                INNER JOIN ""OUOM"" T2 ON T2.""UomEntry"" = T1.""BaseUom""
                                INNER JOIN ""OUOM"" T4 ON T4.""UomCode"" = '{unitCode}'
                                LEFT JOIN  ""UGP1"" T3
                                       ON T3.""UgpEntry"" = T1.""UgpEntry""
                                      AND T3.""UomEntry"" = T4.""UomEntry""
                                WHERE T0.""ItemCode"" = '{article}'";

                oRec.DoQuery(q);

                if (oRec.EoF)
                    return 0m;

                return Convert.ToDecimal(oRec.Fields.Item("FactorConversionAUnidadBase").Value);
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
