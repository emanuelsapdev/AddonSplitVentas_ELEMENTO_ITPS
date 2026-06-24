using Addon_AutoDivSalesOrd.Configuration;

namespace Addon_AutoDivSalesOrd.Common
{
    /// <summary>
    /// Ejecuta tareas de inicialización al arranque del addon:
    /// creación de infraestructura (tablas UDT, campos UDF) y carga de configuraciones.
    /// </summary>
    public class ExecutionsApp
    {
        public ExecutionsApp()
        {
            // Crear infraestructura (tablas y campos UDF) si no existen.
            // Descomentar en primera ejecución o cuando se agreguen nuevos campos:
            new ConfigsDevelopmentInfra().ProcessInfrastructure();
            new SalesOrderSplitInfra().ProcessInfrastructure();

            // Cargar configuraciones de desarrollo al inicio (una sola vez)
            //AppConfig.Load();
        }
    }
}
