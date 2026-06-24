using Addon_AutoDivSalesOrd.Common;
using SAPbobsCOM;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Addon_AutoDivSalesOrd.Configuration
{
    /// <summary>
    /// Administra las configuraciones de desarrollo cargadas desde @CONFIGS_DEVELOPMENT.
    /// Se carga una sola vez al iniciar el addon y se puede recargar bajo demanda.
    /// Reemplaza a GlobalConfigsDevelopments + Utils.GetConfigsDevelopment().
    /// </summary>
    public static class AppConfig
    {
        private static readonly Dictionary<string, string> _configs = new Dictionary<string, string>();
        private static bool _loaded = false;

        /// <summary>
        /// Carga las configuraciones desde la tabla @CONFIGS_DEVELOPMENT.
        /// Si ya están cargadas, no hace nada (usar Reload para forzar).
        /// </summary>
        public static void Load()
        {
            if (_loaded) return;

            Recordset oRec = null;
            try
            {
                oRec = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                oRec.DoQuery(@"SELECT ""U_ITPS_Prop"", ""U_ITPS_Value"" FROM ""@CONFIGS_DEVELOPMENT""");

                while (!oRec.EoF)
                {
                    string prop = oRec.Fields.Item(0).Value?.ToString();
                    string val = oRec.Fields.Item(1).Value?.ToString();

                    if (!string.IsNullOrEmpty(prop))
                        _configs[prop] = val;

                    oRec.MoveNext();
                }

                _loaded = true;
            }
            finally
            {
                if (oRec != null) Marshal.ReleaseComObject(oRec);
            }
        }

        /// <summary>
        /// Obtiene el valor de una configuración por nombre de propiedad.
        /// Si las configuraciones no están cargadas, las carga automáticamente.
        /// </summary>
        /// <param name="propertyName">Nombre de la propiedad (U_ITPS_Prop).</param>
        /// <returns>Valor de la configuración o null si no existe.</returns>
        public static string Get(string propertyName)
        {
            if (!_loaded)
                Load();

            _configs.TryGetValue(propertyName, out string value);
            return value;
        }

        /// <summary>
        /// Fuerza la recarga de todas las configuraciones desde la base de datos.
        /// Útil cuando se abren formularios que pueden tener configuraciones actualizadas.
        /// </summary>
        public static void Reload()
        {
            _loaded = false;
            _configs.Clear();
            Load();
        }
    }
}
