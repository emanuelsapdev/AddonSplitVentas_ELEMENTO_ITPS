using Addon_AutoDivSalesOrd.Addons.Tools;
using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Configuration;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Forms.SalesOrder
{
    /// <summary>
    /// Manejador de eventos del formulario de Registro de Cheques (FormType 607).
    /// Solo contiene lógica de UI y orquestación: delega acceso a datos y
    /// lógica de negocio a Repositories y Services respectivamente.
    /// </summary>
    public partial class SalesOrderFrm : IFormEventHandler
    {
        /// <summary>
        /// FormTypeEx del formulario de Orden de Venta en SAP B1.
        /// </summary>
        public const string FormType = Constants.FormTypes.SalesOrder;

        #region Implementación de IFormEventHandler

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            // ANTES DE CONCRETAR EL PRESIONAR EL BOTON DE CREAR
            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && !pVal.ActionSuccess
                     && pVal.ItemUID == Constants.ItemUids.AddOrUpdateButton && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {

                BubbleEvent = false;
                SAPbouiCOM.Form oForm = null;
                SAPbobsCOM.Recordset oRec = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                    // Eliminar líneas vacías de la matriz antes de procesar
                    //RemoveEmptyLinesFromMatrix(oForm);

                    // ExportMatrixColumnsDetailToTxt(oForm, "C:\\columnsData.txt");  // solo para testear

                    var data = GetDataFromFormOrder(oForm);

                    ApplyImportedOrderRules(data);

                    ValidationService.ValidatePorcentageAtCreate(data);
                    ValidationService.ValidateQuantitiesAtCreate(data);
                    //ValidationService.ValidateStockAtCreate(data);

                    EnsureNoActiveTransaction();
                    ConnectionSDK.DIAPI.StartTransaction();

                    var (entryPrimary, entrySecondary) = ExecuteCreate(data);

                    ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_Commit);

                    SyncDeliveryZone(entryPrimary, entrySecondary);

                    CloseAndReopen(oForm, entryPrimary, entrySecondary);
                }
                catch (OperationCanceledException)
                {
                    RollbackIfNeeded();
                    BubbleEvent = false;
                }
                catch (Exception ex)
                {
                    RollbackIfNeeded();
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null) Marshal.ReleaseComObject(oForm);
                    if (oRec != null) Marshal.ReleaseComObject(oRec);
                }

            }

            // ANTES DE CONCRETAR EL PRESIONAR EL BOTON DE ACTUALIZAR
            else if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && !pVal.ActionSuccess
                     && pVal.ItemUID == Constants.ItemUids.AddOrUpdateButton && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                BubbleEvent = false;
                SAPbouiCOM.Form oForm = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    var data = GetDataFromFormOrder(oForm);
                    var prevQties = GetPrevQuantities(data);

                    ValidationService.ValidateEntityCompanyAtUpdate(data);
                    ValidationService.ValidatePorcentageAtUpdate(data);
                    ValidationService.ValidateQuantitiesAtUpdate(data, prevQties);
                    //ValidationService.ValidateStockAtUpdate(data);

                    EnsureNoActiveTransaction();
                    ConnectionSDK.DIAPI.StartTransaction();

                    var (entryPrimary, entrySecondary) = ExecuteUpdate(data);

                    ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_Commit);

                    CloseAndReopen(oForm, entryPrimary, entrySecondary);
                }
                catch (Exception ex)
                {
                    RollbackIfNeeded();
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null) Marshal.ReleaseComObject(oForm);
                }
            }

            // SALIR DEL FOCO DEL CAMPO "CardCode"
            else if(pVal.EventType == BoEventTypes.et_LOST_FOCUS && !pVal.BeforeAction
                     && (pVal.ItemUID == Constants.SalesOrder_FieldsUIDs.Head_CardCode || pVal.ItemUID == Constants.SalesOrder_FieldsUIDs.Head_CardName) 
                     && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {

                SAPbouiCOM.Form oForm = null;
                SAPbobsCOM.Recordset oRec = null;

                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    oRec = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                    EditText oDocDueDate = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_DocDueDate).Specific;
                    oDocDueDate.Value = DateTime.Now.AddDays(5).ToString("yyyyMMdd");

                    EditText oCardCode = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_CardCode).Specific;
                    string cardCode = oCardCode.Value;

                    if(string.IsNullOrEmpty(cardCode)) return;

                    (decimal commonPerc, decimal importedPerc, string categoryClient) = GetPercentagesAndCategoryBySN(cardCode);

                    oForm.Freeze(true);
                    if (Regex.IsMatch(commonPerc.ToString(), "[100|50|25|0]"))
                    {
                        SAPbouiCOM.ComboBox oSplitPerc = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_SplitPercentage).Specific;
                        oSplitPerc.Select(commonPerc + ",00", BoSearchKey.psk_ByValue);
                    }

                    if (Regex.IsMatch(importedPerc.ToString(), "[100|50|25|0]"))
                    {
                        SAPbouiCOM.ComboBox oImportedPerc = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_ImportedPercentage).Specific;
                        oImportedPerc.Select(importedPerc + ",00", BoSearchKey.psk_ByValue);
                    }

                    if (Regex.IsMatch(categoryClient, "[AA|A|B|C]"))
                    {
                        SAPbouiCOM.ComboBox oCategoryCli = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_CategoryClient).Specific;
                        oCategoryCli.Select(categoryClient, BoSearchKey.psk_ByValue);
                    }
                }
                catch (Exception ex)
                {
                    
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null)
                    {
                        oForm.Freeze(false);
                        Marshal.ReleaseComObject(oForm);
                    }

                    if(oRec !=null) Marshal.ReleaseComObject(oRec);
                }
               
            }

            // SALIR DEL FOCO DE LA COLUMNA "ItemCode"
            else if(pVal.EventType == BoEventTypes.et_LOST_FOCUS && pVal.ActionSuccess
                     && pVal.ItemUID == Constants.SalesOrder_FieldsUIDs.Head_Matrix 
                     && pVal.ColUID == Constants.SalesOrder_FieldsUIDs.Det_ItemCode 
                     && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                SAPbouiCOM.Form oForm = null;
                SAPbobsCOM.Recordset oRec = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    oRec = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                    SAPbouiCOM.ComboBox etImportedPerc = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_ImportedPercentage).Specific;
                    string vImportedPerc = etImportedPerc.Value;

                    Matrix oMtx = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Matrix).Specific;
                    oForm.Freeze(true);

                    string itemCode = (oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_ItemCode, pVal.Row)).Value;
                    if (string.IsNullOrEmpty(itemCode)) return;

                    string q = $@"SELECT 1 FROM OITM WHERE ""ItemCode"" = '{itemCode}' AND ""QryGroup1"" = 'Y'";
                    oRec.DoQuery(q);
                    if (oRec.RecordCount == 0) return;

                    EditText oDiscount = oMtx.Columns.Item(Constants.SalesOrder_FieldsUIDs.Det_Discount).Cells.Item(pVal.Row).Specific;
                    decimal.TryParse(vImportedPerc, out decimal importedPerc);
                    decimal.TryParse(oDiscount.Value.Replace(".", ","), out decimal discountSap);
                    if (string.IsNullOrEmpty(itemCode) || discountSap == importedPerc) return;
                    oDiscount.Value = vImportedPerc.Replace(",", ".");

                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null)
                    {
                        oForm.Freeze(false);
                        Marshal.ReleaseComObject(oForm);
                    }

                    if (oRec != null) Marshal.ReleaseComObject(oRec);
                    
                }
            }

            // AL CAMBIAR EL VALOR DEL SELECTOR PORCENTAJE DE IMPORTADOS
            else if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ActionSuccess
                     && pVal.ItemUID == Constants.SalesOrder_FieldsUIDs.Head_ImportedPercentage
                     // && pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE
                     )
            {
                SAPbouiCOM.Form oForm = null;
                SAPbobsCOM.Recordset oRec = null;

                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    oRec = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                    SAPbouiCOM.ComboBox etImportedPerc = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_ImportedPercentage).Specific;
                    string vImportedPerc = etImportedPerc.Value;
                    decimal.TryParse(vImportedPerc, out decimal importedPerc);

                    //if (importedPerc == 100m) return;

                    Matrix oMtx = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Matrix).Specific;

                    oForm.Freeze(true);

                    for (int row = 1; row <= oMtx.RowCount; row++)
                    {
                        SAPbouiCOM.EditText oDiscount = null;
                        try
                        {

                            oDiscount = oMtx.Columns.Item(Constants.SalesOrder_FieldsUIDs.Det_Discount).Cells.Item(row).Specific;
                            string itemCode = (oMtx.GetCellSpecific(Constants.SalesOrder_FieldsUIDs.Det_ItemCode, row)).Value;

                            string q = $@"SELECT 1 FROM OITM WHERE ""ItemCode"" = '{itemCode}' AND ""QryGroup1"" = 'Y'";
                            oRec.DoQuery(q);
                            if (oRec.RecordCount == 0) continue;

                            
                            decimal.TryParse(oDiscount.Value.Replace(".", ","), out decimal discountSap);
                            if (string.IsNullOrEmpty(itemCode) || discountSap == importedPerc) continue;
                            oDiscount.Value = importedPerc != 100m ? vImportedPerc.Replace(",", ".") : "0.00";
                        }
                        finally
                        {
                            if (oDiscount != null) Marshal.ReleaseComObject(oDiscount);
                        }
                    }

                    oForm.ActiveItem = Constants.SalesOrder_FieldsUIDs.Head_ImportedPercentage;
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null)
                    {
                        oForm.Freeze(false);
                        Marshal.ReleaseComObject(oForm);
                    }

                    if (oRec != null) Marshal.ReleaseComObject(oRec);

                }
            }

        }

        public void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (boi.EventType == BoEventTypes.et_FORM_DATA_LOAD && boi.ActionSuccess)
            {
                SAPbouiCOM.Form oForm = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(boi.FormUID);
                    //oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_SplitPercentage).Enabled = enableItem;
                    oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_Money).Enabled = true;

                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                    BubbleEvent = false;
                }
                finally
                {
                    if (oForm != null)
                        Marshal.ReleaseComObject(oForm);
                }

            }

        }



        public void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = null;
            var baseData = new SalesOrderFormModel();
            Documents primaryOrd = null;
            Documents secondaryOrd = null;
            if (ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_RollBack);
            ConnectionSDK.DIAPI.StartTransaction();

            try
            {
                switch (pVal.MenuUID)
                {
                    case Constants.MenusUID.Cancel:

                        oForm = ConnectionSDK.UIAPI.Forms.ActiveForm;
                        if (oForm.TypeEx != SalesOrderFrm.FormType) return;
                        BubbleEvent = false;

                        baseData = GetDataFromFormOrder(oForm);
                        
                        int entryPrimary = baseData.DocEntry;
                        int entrySecondary = GetRelatedOrder(entryPrimary);

                        string msg = "¿Estas seguro que quieres cancelar este documento?";
                        if (entrySecondary != -1)
                        {
                            msg = "Al cancelar este documento tambien cancelara su orden split relacionada.";
                        }
                        bool confirm = ConnectionSDK.UIAPI.MessageBox(msg, 2, "Confirmar", "Cancelar") == 1;
                        if (!confirm) return;

                        primaryOrd = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                        primaryOrd.GetByKey(entryPrimary);
                        bool hasPCancel = primaryOrd.Cancel() == 0;

                        if (!hasPCancel)
                        {
                            ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                            throw new Exception($"Error al cancelar las ordenes. {errCode} - {errMsg}");
                        }

                        secondaryOrd = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                        if (entrySecondary != -1)
                        {
                            secondaryOrd.GetByKey(entrySecondary);
                            bool hasSCancel = secondaryOrd.Cancel() == 0;

                            if (!hasSCancel)
                            {
                                ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                                throw new Exception($"Error al cancelar las ordenes. {errCode} - {errMsg}");
                            }

                        }

                        if (ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_Commit);
                        oForm.Mode = BoFormMode.fm_OK_MODE;
                        oForm.Close();

                        bool bothOpen = entryPrimary > 0 && entrySecondary > 0;

                        ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, baseData.DocEntry.ToString());

                        if (bothOpen)
                        {
                            var fPrimary = ConnectionSDK.UIAPI.Forms.ActiveForm;
                            fPrimary.Left = 0;
                            fPrimary.Top = 0;

                            if (entrySecondary > 0)
                            {
                                ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, entrySecondary.ToString());
                                var fSecondary = ConnectionSDK.UIAPI.Forms.ActiveForm;
                                fSecondary.Left = fPrimary.Left + fPrimary.Width + 4;
                                fSecondary.Top = fPrimary.Top;
                            }

                        }

                        break;

                    case Constants.MenusUID.Close:

                        oForm = ConnectionSDK.UIAPI.Forms.ActiveForm;
                        if (oForm.TypeEx != SalesOrderFrm.FormType) return;
                        BubbleEvent = false;

                        baseData = GetDataFromFormOrder(oForm);
                        

                        entryPrimary = baseData.DocEntry;
                        entrySecondary = GetRelatedOrder(entryPrimary);

                        msg = "¿Estas seguro que quieres cerrar este documento?";
                        if (entrySecondary != -1)
                        {
                            msg = "Al cerrar este documento tambien cerraras su orden split relacionada.";
                        }
                        confirm = ConnectionSDK.UIAPI.MessageBox(msg, 2, "Confirmar", "Cancelar") == 1;
                        if (!confirm) return;

                        primaryOrd = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                        primaryOrd.GetByKey(entryPrimary);
                        bool hasPClose = primaryOrd.Close() == 0;

                        if (!hasPClose)
                        {
                            ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                            throw new Exception($"Error al cerrar las ordenes. {errCode} - {errMsg}");
                        }

                        secondaryOrd = ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                        if (entrySecondary != -1)
                        {
                            secondaryOrd.GetByKey(entrySecondary);
                            bool hasSClose = secondaryOrd.Close() == 0;

                            if (!hasSClose)
                            {
                                ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                                throw new Exception($"Error al cerrar las ordenes. {errCode} - {errMsg}");
                            }

                        }

                        if (ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_Commit);
                        oForm.Mode = BoFormMode.fm_OK_MODE;
                        oForm.Close();

                        bothOpen = entryPrimary > 0 && entrySecondary > 0;

                        ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, baseData.DocEntry.ToString());

                        if (bothOpen)
                        {
                            var fPrimary = ConnectionSDK.UIAPI.Forms.ActiveForm;
                            fPrimary.Left = 0;
                            fPrimary.Top = 0;

                            if (entrySecondary > 0)
                            {
                                ConnectionSDK.UIAPI.OpenForm(BoFormObjectEnum.fo_Order, null, entrySecondary.ToString());
                                var fSecondary = ConnectionSDK.UIAPI.Forms.ActiveForm;
                                fSecondary.Left = fPrimary.Left + fPrimary.Width + 4;
                                fSecondary.Top = fPrimary.Top;
                            }

                        }

                        break;

                    case Constants.MenusUID.ModeAdd:
                    case Constants.MenusUID.OrdenVenta:
                    case Constants.MenusUID.Duplicate:

                        try
                        {
                            if (pVal.BeforeAction) break;
                            oForm = ConnectionSDK.UIAPI.Forms.ActiveForm;
                            if (oForm.TypeEx != SalesOrderFrm.FormType) return;

                            SAPbouiCOM.ComboBox oCbTipo = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_AssignedEntity).Specific;
                            oForm.Freeze(true);
                            oCbTipo.Select("NORMAL");
                            oForm.ActiveItem = Constants.SalesOrder_FieldsUIDs.Head_CardCode;
                            oCbTipo.Item.Enabled = false;
                        }
                        catch { }
                        finally
                        {
                            if (oForm != null)
                            {
                                oForm.Freeze(false);
                                Marshal.ReleaseComObject(oForm);
                            }
                        }
                        break;

                    case Constants.MenusUID.RowInitial:
                    case Constants.MenusUID.RowBefore:
                    case Constants.MenusUID.RowAfter:
                    case Constants.MenusUID.RowEnd:
                    case Constants.MenusUID.RowsRefresh:
                        try
                        {
                            if (pVal.BeforeAction) break;
                            oForm = ConnectionSDK.UIAPI.Forms.ActiveForm;
                            if (oForm.TypeEx != SalesOrderFrm.FormType) return;
                            SAPbouiCOM.ComboBox oCbTipo = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_AssignedEntity).Specific;
                            SAPbouiCOM.ComboBox oCbSplitPerc = oForm.Items.Item(Constants.SalesOrder_FieldsUIDs.Head_SplitPercentage).Specific;
                            oForm.Freeze(true);
                            oForm.ActiveItem = Constants.SalesOrder_FieldsUIDs.Head_CardCode;
                            oCbTipo.Item.Enabled = false;
                            oCbSplitPerc.Item.Enabled = false;
                        }
                        catch { }
                        finally
                        {
                            if (oForm != null)
                            {
                                oForm.Freeze(false);
                                Marshal.ReleaseComObject(oForm);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                if (ConnectionSDK.DIAPI.InTransaction) ConnectionSDK.DIAPI.EndTransaction(BoWfTransOpt.wf_RollBack);
                NotificationService.Error(ex.Message);
                BubbleEvent = false;
            }
            finally
            {
                if (oForm != null)
                    Marshal.ReleaseComObject(oForm);

                if (primaryOrd != null)
                    Marshal.ReleaseComObject(primaryOrd);

                if (secondaryOrd != null)
                    Marshal.ReleaseComObject(secondaryOrd);
            }

        }

        #endregion


    }
}