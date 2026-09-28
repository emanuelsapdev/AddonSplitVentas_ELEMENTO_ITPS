using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Services;
using SAPbouiCOM;
using System;

namespace Addon_AutoDivSalesOrd.Forms.CreditNote
{
    /// <summary>
    /// Manejador de eventos del formulario de Nota de Crédito de Venta (FormType 179).
    /// Detecta notas de crédito de artículos importados y genera el asiento contable
    /// que revierte la deuda de empresa B.
    /// </summary>
    public partial class CreditNoteFrm : IFormEventHandler
    {
        public const string FormType = Constants.FormTypes.CreditNote;

        private const decimal RetryJournalImportedPerc = 50m;

        #region Implementación de IFormEventHandler

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && !pVal.BeforeAction)
            {
                try
                {
                    AddBtnRetryJournal(FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error($"Error al agregar el botón de reintento de asiento. {ex.Message}");
                }
                return;
            }

            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ActionSuccess
                     && pVal.ItemUID == Constants.CreditNote_FieldsUIDs.Head_BtnRetryJournal)
            {
                try
                {
                    RetryImportJournalEntry(FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
                return;
            }

            if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ActionSuccess
                     && pVal.ItemUID == Constants.CreditNote_FieldsUIDs.Head_ImportedPercentage)
            {
                try
                {
                    ApplyImportedPercDiscountToLines(FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
            }
        }

        public void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if ((boi.EventType == BoEventTypes.et_FORM_DATA_LOAD || boi.EventType == BoEventTypes.et_FORM_DATA_UPDATE)
                     && boi.ActionSuccess)
            {
                try
                {
                    UpdateBtnRetryJournalState(boi.FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error($"Error al actualizar el estado del botón de asiento de importados. {ex.Message}");
                }
                return;
            }

            if (boi.EventType == BoEventTypes.et_FORM_DATA_ADD && boi.ActionSuccess)
            {
                try
                {
                    GenerateImportJournalEntryOnCreditNoteAdd(boi.FormUID);
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
            }
        }

        public void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        #endregion
    }
}
