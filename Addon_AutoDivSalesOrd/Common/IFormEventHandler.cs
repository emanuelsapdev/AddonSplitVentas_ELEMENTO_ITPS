using SAPbouiCOM;

namespace Addon_AutoDivSalesOrd.Common
{
    /// <summary>
    /// Contrato para manejadores de eventos de formularios SAP B1.
    /// Cada formulario personalizado debe implementar esta interfaz
    /// para ser registrado en el EventRouter.
    /// </summary>
    public interface IFormEventHandler
    {
        void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent);
        void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent);
        void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent);
    }
}
