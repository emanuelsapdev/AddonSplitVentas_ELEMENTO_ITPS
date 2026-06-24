using Addon_AutoDivSalesOrd.Addons.Tools;
using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using Addon_AutoDivSalesOrd.Tools;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Forms.GerentePicking
{
    public partial class GerentePickingFrm : IFormEventHandler
    {
        /// <summary>
        /// FormTypeEx del formulario de Orden de Venta en SAP B1.
        /// </summary>
        public const string FormType = Constants.FormTypes.GerentePicking;

        public void OnFormDataEvent(ref BusinessObjectInfo boi, out bool BubbleEvent)
        {
            throw new NotImplementedException();
        }

        public void OnItemEvent(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;


            //if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
            //{
            //    var oMtx = (Matrix)ConnectionSDK.UIAPI.Forms.Item(FormUID).Items.Item(Constants.GerentePicking_FieldsUIDs.Head_MatrixOpen).Specific;
            //    IUDataTest.ExportMatrixColumnsToTxt(oMtx, @"C:\GerentePicking_MatrixColumns.txt");
            //}

            if (pVal.EventType == BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
            {
                try
                {
                    AddBtnProcess(FormUID);                   
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
            }

            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction && 
                (pVal.ItemUID == Constants.GerentePicking_FieldsUIDs.Head_FolderOpen || 
                pVal.ItemUID == Constants.GerentePicking_FieldsUIDs.Head_FolderRelease ||
                pVal.ItemUID == Constants.GerentePicking_FieldsUIDs.Head_FolderPickingPerformed
                )
                )
            {
                SAPbouiCOM.Form oForm = null;
                try
                {
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);
                    bool enableBtn = false;
                    if (oForm.PaneLevel == Constants.GerentePicking_LevelPanel.OpenLevelPane)
                        enableBtn = true;
                    
                    ToggleEnableBtnProcess(enableBtn);
                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
            }


            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction && pVal.ItemUID == Constants.GerentePicking_FieldsUIDs.Head_BtnProcess)
            {
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.ProgressBar oProgress = null;
                try
                {
                    oProgress = ConnectionSDK.UIAPI.StatusBar.CreateProgressBar("Comenzando el proceso", 100, false);
                    oForm = ConnectionSDK.UIAPI.Forms.Item(FormUID);

                    if (oForm.PaneLevel != Constants.GerentePicking_LevelPanel.OpenLevelPane) return;

                    var oMtx = (Matrix)oForm.Items.Item(Constants.GerentePicking_FieldsUIDs.Head_MatrixOpen).Specific;

                    oForm.Freeze(true);

                    oProgress.Text = "Analizando registros";
                    oProgress.Value = 30;

                    ClearAllToReleaseValues(oMtx);

                    var dataMtx = ReadRows(oMtx, GetAvailableStock);

                    oProgress.Text = "Realizando Calculos";
                    oProgress.Value = 50;

                    var groups = GroupBySplit(dataMtx);
                    CalculateToRelease(groups);

                    oProgress.Text = "Aplicando Cantidades";
                    foreach (var splitGroups in groups.Values)
                        foreach (var splitGroup in splitGroups)
                            foreach (var row in splitGroup)
                            {
                                oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_ToRelease, row.LineNum).Value = row.ToRelease.ToString();
                                oProgress.Value += 1;
                            }

                    oProgress.Value = 100;
                    oProgress.Stop();

                    NotificationService.Warn("Tarea completada");

                }
                catch (Exception ex)
                {
                    NotificationService.Error(ex.Message);
                }
                finally
                {
                    if(oForm != null)
                    {
                        oForm.Freeze(false);
                        MarshalGC.ReleaseComObject(oForm);
                    }

                    if(oProgress != null)
                    {
                        oProgress.Stop();
                        MarshalGC.ReleaseComObject(oProgress);
                    }
                }
            }
        }

        public void OnMenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
        }
    }
}
