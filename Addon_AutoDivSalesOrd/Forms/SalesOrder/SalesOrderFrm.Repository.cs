using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Forms.SalesOrder
{
    public partial class SalesOrderFrm
    {
        private Dictionary<string, bool> GetItemsImportedFlags(List<string> itemCodes)
        {
            var result = new Dictionary<string, bool>();
            if (itemCodes == null || itemCodes.Count == 0) return result;

            Recordset rs = null;
            try
            {
                rs = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                string inClause = string.Join(",", itemCodes.Select(c => $"'{c.Replace("'", "''")}'"));
                string q = $@"SELECT ""ItemCode"", ""{Constants.ItemImport.OitmImportProperty}"" FROM OITM WHERE ""ItemCode"" IN ({inClause})";
                rs.DoQuery(q);
                while (!rs.EoF)
                {
                    string code = rs.Fields.Item(0).Value?.ToString();
                    string prop = rs.Fields.Item(1).Value?.ToString();
                    if (!string.IsNullOrEmpty(code))
                        result[code] = prop == Constants.FixedValues.Yes;
                    rs.MoveNext();
                }
            }
            finally
            {
                if (rs != null) Marshal.ReleaseComObject(rs);
            }
            return result;
        }

        private int GetRelatedOrder(int entryPrincipalOrd)
        {
            Recordset rs = null;
            try
            {
                rs = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                string q = $@"SELECT ""DocEntry"" FROM ORDR WHERE ""U_ITPS_RelatedOrder"" = '{entryPrincipalOrd}'";
                rs.DoQuery(q);

                int entrySecondaryOrd = -1;
                while (!rs.EoF)
                {
                    entrySecondaryOrd = rs.Fields.Item(0).Value;
                    break;
                }

                // if (entrySecondaryOrd == -1) throw new Exception("No existe una orden relacionada al split. Imposible modificar un pedido secundario, modificar el principal");

                return entrySecondaryOrd;
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        private Dictionary<int, double> GetPrevQuantities(SalesOrderFormModel data)
        {
            var quantities = new Dictionary<int, double>();
            Recordset rs = null;
            try
            {
                rs = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                string q = $@"SELECT ""LineNum"", ""Quantity"" FROM RDR1 WHERE ""DocEntry"" = '{data.DocEntry}'";
                rs.DoQuery(q);
                while (!rs.EoF)
                {
                    int lineNum = rs.Fields.Item(0).Value;
                    double qty = rs.Fields.Item(1).Value;
                    quantities.Add(lineNum, qty);
                    rs.MoveNext();
                }
                return quantities;
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        private decimal? GetLinePriceByAgreement(string itemCode, string agreementNumber)
        {
            if (string.IsNullOrWhiteSpace(itemCode) || string.IsNullOrWhiteSpace(agreementNumber))
                return null;

            Recordset rs = null;
            try
            {
                rs = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string safeItemCode = itemCode.Replace("'", "''");
                string safeAgreementNumber = agreementNumber.Replace("'", "''");

                string q = $@"SELECT T2.""Price""
                             FROM OOAT T0
                             INNER JOIN ITM1 T2
                                 ON T2.""PriceList"" = T0.""ListNum""
                                AND T2.""ItemCode"" = '{safeItemCode}'
                             WHERE TO_VARCHAR(T0.""AbsID"") = '{safeAgreementNumber}'
                               AND T0.""BpType"" = 'C'";

                rs.DoQuery(q);

                if (rs.EoF || rs.Fields.Item(0).Value == null)
                    return null;

                return Convert.ToDecimal(rs.Fields.Item(0).Value);
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        private string GetAgreementPriceListName(int globalAgreement)
        {
            if (globalAgreement <= 0)
                return string.Empty;

            Recordset rs = null;
            try
            {
                rs = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                string q = $@"SELECT T1.""ListName""
                             FROM OOAT T0
                             INNER JOIN OPLN T1 ON T1.""ListNum"" = T0.""ListNum""
                             WHERE TO_VARCHAR(T0.""AbsID"") = '{globalAgreement}'
                               AND T0.""BpType"" = 'C'";
                rs.DoQuery(q);

                if (rs.EoF || rs.Fields.Item(0).Value == null)
                    return string.Empty;

                return rs.Fields.Item(0).Value.ToString().Trim();
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        private void SyncDeliveryZone(int entryPrimary, int entrySecondary)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string update = $@"
                        UPDATE ORDR T0
                        SET T0.""U_ITPS_DeliveryZone"" = T1.""U_ITPS_DeliveryZoneS""
                        FROM ORDR T0
                        INNER JOIN RDR12 T1 ON T1.""DocEntry"" = T0.""DocEntry""
                        WHERE T0.""DocEntry"" IN ('{entryPrimary}', '{entrySecondary}')";

                rs.DoQuery(update);
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }

        private (decimal commonPerc, decimal importedPerc, string categoryCli) GetPercentagesAndCategoryBySN(string cardCode)
        {
            Recordset rs = null;

            decimal commonPerc = -1m;
            decimal importedPerc = -1m;
            string categoryCli = string.Empty;
            try
            {

                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                string q = $@"SELECT IFNULL(""U_ITPS_PRIORIDAD"", -1) AS ""U_ITPS_PRIORIDAD"", IFNULL(""U_ITPS_PRIORIDAD_IMPORTADOS"", 0) AS ""U_ITPS_PRIORIDAD_IMPORTADOS"", IFNULL(""U_CATEGORIA_CLIENTE"", '') AS ""U_CATEGORIA_CLIENTE"" FROM OCRD WHERE ""CardCode"" = '{cardCode}'";
                rs.DoQuery(q);
                if (!rs.EoF)
                {
                    commonPerc = Convert.ToDecimal(rs.Fields.Item(0).Value);
                    importedPerc = Convert.ToDecimal(rs.Fields.Item(1).Value);
                    categoryCli = Convert.ToString(rs.Fields.Item(2).Value);
                }
                return (commonPerc, importedPerc, categoryCli);
            }
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
        }
    }
}
