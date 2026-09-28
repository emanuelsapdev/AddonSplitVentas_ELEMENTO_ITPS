using Addon_AutoDivSalesOrd.Common;
using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Forms.CreditNote
{
    public partial class CreditNoteFrm
    {
        /// <summary>
        /// Detecta si una Nota de Crédito de Venta es de importados.
        /// Obtiene los campos U_Importado y U_ITPS_ImportedPercentage directamente de la tabla ORIN.
        /// </summary>
        /// <returns>DocEntry de la nota de crédito (-1 si no es importada) y el porcentaje de importados.</returns>
        private (int docEntry, decimal importedPerc) GetImportedDocEntryAndImportedPercFromCreditNote(int creditNoteDocEntry)
        {
            Recordset rs = null;
            int docEntry = -1;
            decimal importedPerc = 0m;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Solo obtiene de ORIN los campos de bandera de importados y su porcentaje
                string q = $@"
                    SELECT T0.""DocEntry"", T0.""U_ITPS_ImportedPercentage""
                    FROM ORIN T0
                    WHERE T0.""DocEntry"" = {creditNoteDocEntry} AND T0.""U_Importado"" = 'Y'";

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
        /// Busca un asiento asociado a la nota de crédito por medio de OJDT."U_ITPS_RelatedCreditNote".
        /// </summary>
        /// <returns>TransId del asiento, o -1 si no existe.</returns>
        private int GetJournalEntryForCreditNote(int creditNoteDocEntry)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string q = $@"
                    SELECT TOP 1 T0.""TransId""
                    FROM OJDT T0
                    WHERE T0.""U_ITPS_RelatedCreditNote"" = {creditNoteDocEntry}
                    ORDER BY T0.""TransId"" DESC";

                rs.DoQuery(q);

                if (rs.EoF) return -1;

                return Convert.ToInt32(rs.Fields.Item(0).Value);
            }
            finally
            {
                if (rs != null) Marshal.ReleaseComObject(rs);
            }
        }

        /// <summary>
        /// Obtiene el importe total de descuento de importados de la Nota de Crédito:
        /// sumatoria de (PriceBefDi × importedPerc / 100 × Quantity) por línea,
        /// menos el descuento de cabecera. También retorna el CardCode del cliente.
        /// </summary>
        private (string cardCode, double totalDiscountAmount) GetCreditNoteImportDiscountData(int creditNoteDocEntry, decimal importedPerc)
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
                    FROM ORIN T0
                    INNER JOIN RIN1 T1 ON T1.""DocEntry"" = T0.""DocEntry""
                    WHERE T0.""DocEntry"" = {creditNoteDocEntry}
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
