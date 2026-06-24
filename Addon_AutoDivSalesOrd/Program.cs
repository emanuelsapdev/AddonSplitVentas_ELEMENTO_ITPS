using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Services;
using System;

namespace Addon_AutoDivSalesOrd
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                // 1. Conectar a SAP Business One (UI API + DI API)
                ConnectionSDK.Singlenton();

                if (!ConnectionSDK.Connected)
                    return;

                NotificationService.Success(Constants.Messages.ConnectedOk);

                // 2. Ejecutar inicializaciones (infraestructura + carga de configuraciones)
                var executions = new ExecutionsApp();

                // 3. Registrar enrutador de eventos de UI API
                var eventRouter = new EventRouter();

                // 4. Mantener referencias vivas para evitar recolección por GC
                GC.KeepAlive(eventRouter);
                GC.KeepAlive(executions);

                // 5. Iniciar bucle de mensajes (mantiene el addon activo)
                System.Windows.Forms.Application.Run();
            }
            catch (Exception ex)
            {
                ConnectionSDK.UIAPI?.MessageBox(Constants.Messages.FatalErrorPrefix + ex.Message);
            }
        }
    }
}
