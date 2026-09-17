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
        /// Detecta si una Factura de Venta es de importados.
        /// Obtiene los campos U_Importado y U_ITPS_ImportedPercentage directamente de la tabla OINV.
        /// </summary>
        /// <returns>DocEntry de la factura (-1 si no es importada) y el porcentaje de importados.</returns>
        private (int docEntry, decimal importedPerc) GetImportOrderDocEntryAndImportedPercFromInvoice(int invoiceDocEntry)
        {
            Recordset rs = null;
            int docEntry = -1;
            decimal importedPerc = 0m;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Solo obtiene de OINV los campos de bandera de importados y su porcentaje
                string q = $@"
                    SELECT T0.""DocEntry"", T0.""U_ITPS_ImportedPercentage""
                    FROM OINV T0
                    WHERE T0.""DocEntry"" = {invoiceDocEntry} AND T0.""U_Importado"" = 'Y'";

                rs.DoQuery(q);
                
                if (rs.EoF) return (docEntry, importedPerc);

                docEntry = Convert.ToInt32(rs.Fields.Item(0).Value);
                importedPerc = Convert.ToDecimal(rs.Fields.Item(1).Value ?? 0);
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
