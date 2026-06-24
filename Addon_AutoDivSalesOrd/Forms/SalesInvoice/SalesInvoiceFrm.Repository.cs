using Addon_AutoDivSalesOrd.Common;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.SalesInvoice
{
    public partial class SalesInvoiceFrm
    {
        /// <summary>
        /// Determina si una Factura de Venta proviene de una Orden de Artículos Importados.
        /// Traza la cadena OINV → INV1 → ODLN → DLN1 → ORDR → RDR1 → OITM
        /// y verifica que al menos una línea tenga QryGroup1 = 'Y'.
        /// </summary>
        /// <returns>DocEntry de la ORDR de importados, o -1 si no aplica.</returns>
        private (int docEntry, decimal importedPerc) GetImportOrderDocEntryAndImportedPercFromInvoice(int invoiceDocEntry)
        {
            Recordset rs = null;
            int docEntry = -1;
            decimal importedPerc = 0m;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Traza: OINV → INV1 (BaseType=15=ODLN) → DLN1 (BaseType=17=ORDR) → RDR1 → OITM
                // Devuelve el DocEntry de la ORDR si todos sus artículos son importados (QryGroup1='Y')
                string q = $@"
                    SELECT TOP 1 T3.""DocEntry"", T3.""U_ITPS_ImportedPercentage""
                    FROM OINV T0
                    INNER JOIN INV1  T1 ON T1.""DocEntry"" = T0.""DocEntry""  AND T1.""BaseType"" = 15
                    INNER JOIN DLN1  T2 ON T2.""DocEntry"" = T1.""BaseEntry"" AND T2.""LineNum"" = T1.""BaseLine"" AND T2.""BaseType"" = 17
                    INNER JOIN ORDR  T3 ON T3.""DocEntry"" = T2.""BaseEntry""
                    INNER JOIN RDR1  T4 ON T4.""DocEntry"" = T3.""DocEntry""
                    INNER JOIN OITM  T5 ON T5.""ItemCode"" = T4.""ItemCode""
                    WHERE T0.""DocEntry"" = {invoiceDocEntry}
                    GROUP BY T3.""DocEntry"", T3.""U_ITPS_ImportedPercentage""
                    HAVING COUNT(T4.""LineNum"") = SUM(CASE WHEN T5.""{Constants.ItemImport.OitmImportProperty}"" = '{Constants.FixedValues.Yes}' THEN 1 ELSE 0 END)";

                rs.DoQuery(q);
                
                if (rs.EoF) ;

                docEntry = Convert.ToInt32(rs.Fields.Item(0).Value);
                importedPerc = Convert.ToDecimal(rs.Fields.Item(1).Value);
                return (docEntry, importedPerc);
            }
            finally
            {
                if (rs != null) Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Obtiene el importe total de descuento de importados de la Factura:
        /// sumatoria de (PriceAfterVAT × DiscountPercent / 100 × Quantity) por línea.
        /// También retorna el CardCode del cliente.
        /// </summary>
        private (string cardCode, double totalDiscountAmount) GetInvoiceImportDiscountData(int invoiceDocEntry, decimal importedPerc)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string q = $@"
                    WITH LineAmounts AS (
                    SELECT
                        T0.""CardCode"",
                        T0.""DiscPrcnt""                                                     AS ""HeaderDisc"",
                        (T1.""PriceBefDi"" * {importedPerc} / 100) * T1.""Quantity""       AS ""LineTot""
                    FROM OINV T0
                    INNER JOIN INV1 T1 ON T1.""DocEntry"" = T0.""DocEntry""
                    WHERE T0.""DocEntry"" = {invoiceDocEntry}
                )
                SELECT
                    ""CardCode"",
                    SUM(""LineTot"")
                        - SUM(""LineTot"") * (CASE WHEN ""HeaderDisc"" > 0 THEN ""HeaderDisc"" ELSE 0 END) / 100
                        AS ""DiscountTotal""
                FROM LineAmounts
                GROUP BY ""CardCode"", ""HeaderDisc"";";

                rs.DoQuery(q);

                if (rs.EoF) return (null, 0);

                string cardCode = rs.Fields.Item(0).Value?.ToString();
                double total = Convert.ToDouble(rs.Fields.Item(1).Value);
                return (cardCode, total);
            }
            finally
            {
                if (rs != null) Marshal.ReleaseComObject(rs);
            }
        }
    }
}
