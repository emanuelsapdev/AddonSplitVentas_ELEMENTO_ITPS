using Addon_AutoDivSalesOrd.Forms.GerentePicking;
using Addon_AutoDivSalesOrd.Forms.SalesInvoice;
using Addon_AutoDivSalesOrd.Forms.SalesOrder;
using Addon_AutoDivSalesOrd.Forms.SalesQuote;
using Addon_AutoDivSalesOrd.Services;
using SAPbouiCOM;
using System;
using System.Collections.Generic;

namespace Addon_AutoDivSalesOrd.Common
{
    /// <summary>
    /// Enruta los eventos de UI API a los manejadores de formulario correspondientes.
    /// Usa un patrón de registro: cada formulario se instancia una vez y se reutiliza
    /// durante todo el ciclo de vida del addon.
    /// Reemplaza a la clase Events anterior.
    /// </summary>
    public class EventRouter
    {
        // Registro de manejadores: clave = FormTypeEx, valor = handler singleton
        private readonly Dictionary<string, IFormEventHandler> _handlers;

        public EventRouter()
        {
            // Registrar todos los formularios del addon
            _handlers = new Dictionary<string, IFormEventHandler>
            {
                { SalesQuoteFrm.FormType, new SalesQuoteFrm() },
                { SalesOrderFrm.FormType, new SalesOrderFrm() },
                { GerentePickingFrm.FormType, new GerentePickingFrm() },
                { SalesInvoiceFrm.FormType, new SalesInvoiceFrm() }
            };

            try
            {
                ConnectionSDK.UIAPI.AppEvent += OnAppEvent;
                ConnectionSDK.UIAPI.MenuEvent += OnMenuEvent;
                ConnectionSDK.UIAPI.ItemEvent += OnItemEvent;
                ConnectionSDK.UIAPI.FormDataEvent += OnFormDataEvent;
            }
            catch (Exception ex)
            {
                NotificationService.Error(Constants.Messages.EventRouterInitErrorPrefix + ex.Message);
            }
        }

   

        #region Delegación de eventos

        private void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (_handlers.TryGetValue(pVal.FormTypeEx, out var handler))
                    handler.OnItemEvent(FormUID, ref pVal, out BubbleEvent);
            }
            catch (Exception ex)
            {
                NotificationService.Error(Constants.Messages.ItemEventErrorPrefix + ex.Message);
            }
        }

        

        private void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (_handlers.TryGetValue(boi.FormTypeEx, out var handler))
                    handler.OnFormDataEvent(ref boi, out BubbleEvent);
            }
            catch (Exception ex)
            {
                NotificationService.Error(Constants.Messages.FormDataEventErrorPrefix + ex.Message);
            }
        }

        private void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                // Los eventos de menú se propagan a todos los handlers registrados
                foreach (var handler in _handlers.Values)
                    handler.OnMenuEvent(ref pVal, out BubbleEvent);
            }
            catch (Exception ex)
            {
                NotificationService.Error(Constants.Messages.MenuEventErrorPrefix + ex.Message);
            }
        }

        private void OnAppEvent(BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    // Cierre de la aplicación SAP
                    case BoAppEventTypes.aet_ShutDown:
                    case BoAppEventTypes.aet_ServerTerminition:
                    case BoAppEventTypes.aet_CompanyChanged:
                        ConnectionSDK.UIAPI?.StatusBar.SetText(Constants.Messages.FinalizingAddon);

                        System.Windows.Forms.Application.Exit();
                        break;

                    // Cambio de fuente o idioma: reiniciar addon
                    case BoAppEventTypes.aet_FontChanged:
                    case BoAppEventTypes.aet_LanguageChanged:
                        ConnectionSDK.UIAPI?.StatusBar.SetText(Constants.Messages.RestartingAddon);

                        System.Windows.Forms.Application.Restart();
                        break;
                }
            }
            catch (Exception ex)
            {
                NotificationService.Error(Constants.Messages.AppEventErrorPrefix + ex.Message);
            }
        }

        #endregion
    }
}
