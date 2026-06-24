using Addon_AutoDivSalesOrd.Common;
using SAPbobsCOM;
using System;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Repositories
{
    /// <summary>
    /// Repositorio para consultas de socios de negocio (OCRD) y condiciones fiscales.
    /// Todos los Recordset se liberan correctamente con Marshal.ReleaseComObject.
    /// </summary>
    public static class BusinessPartnerRepository
    {
        /// <summary>
        /// Obtiene la categoría de IVA y si es extranjero para un socio de negocio.
        /// </summary>
        /// <param name="cardCode">Código del socio de negocio.</param>
        /// <returns>Tupla con categoría IVA y flag de extranjero.</returns>
        public static (string Category, bool IsForeigner) GetIvaCategory(string cardCode)
        {
            if (string.IsNullOrEmpty(cardCode))
                return (null, false);

            Recordset oRec = null;
            try
            {
                oRec = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = $@"
                    SELECT ""U_B1SYS_VATCtg"", ""VatStatus"" 
                    FROM OCRD 
                    WHERE ""CardCode"" = '{EscapeSql(cardCode)}'";

                oRec.DoQuery(query);

                if (oRec.RecordCount == 0)
                    return (null, false);

                string ctg = oRec.Fields.Item(0).Value?.ToString();
                bool isForeigner = oRec.Fields.Item(1).Value?.ToString() == "N";
                return (ctg, isForeigner);
            }
            finally
            {
                if (oRec != null) Marshal.ReleaseComObject(oRec);
            }
        }

        /// <summary>
        /// Obtiene el código de condición AFIP para un socio de negocio
        /// desde la tabla VTL_CODCONDAFIP cruzada con la categoría IVA en OCRD.
        /// </summary>
        /// <param name="cardCode">Código del socio de negocio.</param>
        /// <returns>CodCond o 0 si no existe.</returns>
        public static int GetAfipConditionCode(string cardCode)
        {
            if (string.IsNullOrEmpty(cardCode))
                return 0;

            Recordset oRec = null;
            try
            {
                oRec = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = $@"
                    SELECT CAST(V0.""CodCond"" AS INT)
                    FROM VTL_CODCONDAFIP V0
                    INNER JOIN OCRD T1 
                        ON T1.""CardCode"" = '{EscapeSql(cardCode)}'
                       AND T1.""U_B1SYS_VATCtg"" = V0.""U_B1SYS_VATCtg""";

                oRec.DoQuery(query);

                if (oRec.RecordCount == 0)
                    return 0;

                int.TryParse(oRec.Fields.Item(0).Value?.ToString(), out int result);
                return result;
            }
            finally
            {
                if (oRec != null) Marshal.ReleaseComObject(oRec);
            }
        }

        /// <summary>
        /// Determina la letra de documento fiscal (A, B, E, etc.)
        /// según la categoría IVA del cliente.
        /// </summary>
        /// <param name="cardCode">Código del socio de negocio.</param>
        /// <returns>Letra del comprobante fiscal o string vacío.</returns>
        public static string GetDocumentLetter(string cardCode)
        {
            var (ivaCategory, isForeigner) = GetIvaCategory(cardCode);
            if (string.IsNullOrEmpty(ivaCategory))
                return string.Empty;

            ivaCategory = ivaCategory.ToUpperInvariant();

            // Extranjeros exentos → Letra E
            if (isForeigner && ivaCategory == "EX") return "E";

            switch (ivaCategory)
            {
                case "MT": // Monotributo
                case "RI": // Responsable Inscripto
                    return "A";

                case "CF":  // Consumidor Final
                case "RNI": // Responsable No Inscripto
                case "NG":  // No Gravado
                case "NA":  // No Alcanzado
                case "NC":  // No Categorizado
                case "EX":  // Exento (no extranjero)
                    return "B";

                default:
                    return string.Empty;
            }
        }

        private static string EscapeSql(string s) => (s ?? string.Empty).Replace("'", "''");

        /// <summary>
        /// Obtiene el porcentaje de split definido en la ficha del socio de negocio.
        /// </summary>
        public static decimal GetSplitPercentage(string cardCode)
        {
            if (string.IsNullOrEmpty(cardCode))
                return 0m;

            Recordset oRec = null;
            try
            {
                oRec = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = $@"SELECT ""{Constants.BusinessPartner_Fields.Head_SplitPercentage}"" FROM OCRD WHERE ""CardCode"" = '{EscapeSql(cardCode)}'";
                oRec.DoQuery(query);
                if (oRec.EoF) return 0m;
                var val = oRec.Fields.Item(0).Value;
                return val == null ? 0m : Convert.ToDecimal(val);
            }
            finally
            {
                if (oRec != null) Marshal.ReleaseComObject(oRec);
            }
        }
    }
}
