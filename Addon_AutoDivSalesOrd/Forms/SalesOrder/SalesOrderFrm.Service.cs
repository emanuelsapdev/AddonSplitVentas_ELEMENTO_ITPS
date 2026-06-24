using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Configuration;
using Addon_AutoDivSalesOrd.Models;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Forms.SalesOrder
{
    public partial class SalesOrderFrm
    {
        public void ApplyImportedOrderRules(SalesOrderFormModel data)
        {
            var itemCodes = data.Lines.Select(l => l.ItemCode).Where(c => !string.IsNullOrWhiteSpace(c)).Distinct().ToList();
            if (itemCodes.Count == 0) return;

            var importedFlags = GetItemsImportedFlags(itemCodes);
            bool anyImported = importedFlags.Values.Any(v => v);
            if (!anyImported) return;

            ValidationService.ValidateImportedOrderItems(importedFlags);

            // Orden de Importados: split al 100% NORMAL y descuento de línea = % split del SN
            //decimal bpSplitPerc = Repositories.BusinessPartnerRepository.GetSplitPercentage(data.CardCode);

            data.IsImportOrder = true;
            data.ImportLineDiscount = data.ImportedPercentage;  
            data.SplitPercentage = 100m;
            data.AssignedEntity = Constants.FixedValues.EntityA;
        }

        public (decimal qtyPrincipal, decimal qtySecondary) CalculateQuantities(decimal baseQuantity, decimal percentage)
        {
            var qtyPrincipal = decimal.Round(baseQuantity * (percentage / 100m), 2, MidpointRounding.AwayFromZero);
            var qtySecondary = decimal.Round(baseQuantity - qtyPrincipal, 2, MidpointRounding.AwayFromZero);
            return (qtyPrincipal, qtySecondary);
        }

        public void ShowMessageConfirm(SalesOrderFormModel data)
        {
            string msg;

            if (data.SplitPercentage == 0m)
            {
                msg = "Porcentaje: 0%\nSe creará la orden al 100% para PRESUPUESTO.\n¿Desea continuar?";
            }
            else if (data.SplitPercentage == 100m)
            {
                msg = "Porcentaje: 100%\nSe creará la orden al 100% para NORMAL.\n¿Desea continuar?";
            }
            else
            {
                decimal percB = 100m - data.SplitPercentage;
                msg = $"Porcentaje: {data.SplitPercentage}%\nSe crearán dos órdenes:\n  - NORMAL: {data.SplitPercentage}%\n  - PRESUPUESTO: {percB}%\n¿Desea continuar?";
            }

            int confirm = ConnectionSDK.UIAPI.MessageBox(msg, 2, "Confirmar", "Cancelar");
            if (confirm != 1)
                throw new OperationCanceledException("Creación cancelada por el usuario.");
        }

        /// <summary>
        /// Crea las órdenes según el porcentaje de split.
        /// 0% → solo principal | 100% → solo secundaria | else → ambas vinculadas
        /// </summary>
        private (int entryPrimary, int entrySecondary) ExecuteCreate(SalesOrderFormModel data)
        {
            ShowMessageConfirm(data);

            if (data.SplitPercentage == 0m)
            {
                int secondary = CreateSecondaryOrder(data);
                return (-1, secondary);
            }

            if (data.SplitPercentage == 100m)
            {
                int primary = CreatePrincipalOrder(data);
                return (primary, -1);
            }

            // Split parcial: crear ambas y vincularlas
            int entryPrimary = CreatePrincipalOrder(data);
            data.RelatedOrder = entryPrimary.ToString();
            int entrySecondary = CreateSecondaryOrder(data);

            // Actualizar RelatedOrder en la orden primaria
            UpdateRelatedOrderInPrimaryOrder(entryPrimary, entrySecondary);

            return (entryPrimary, entrySecondary);
        }

        /// <summary>
        /// Actualiza según entidad asignada y porcentaje.
        /// Solo-principal, solo-secundaria, o split completo.
        /// </summary>
        private (int entryPrimary, int entrySecondary) ExecuteUpdate(SalesOrderFormModel data)
        {
            bool isOnlyPrincipal = data.SplitPercentage == 100m
                                && data.AssignedEntity == Constants.FixedValues.EntityA;

            bool isOnlySecondary = data.SplitPercentage == 100m
                                && data.AssignedEntity == Constants.FixedValues.EntityB;

            if (isOnlyPrincipal)
            {
                UpdatePrincipalOrder(data);
                return (data.DocEntry, -1);
            }

            if (isOnlySecondary)
            {
                string taxCodeSecondary = AppConfig.Get(Constants.ConfigProps.TaxCodeSecondaryOrder);
                data.Lines.ForEach(l => l.TaxCode = taxCodeSecondary);
                UpdatePrincipalOrder(data);
                return (-1, data.DocEntry);
            }

            // Split completo: actualizar principal y su secundaria vinculada
            UpdatePrincipalOrder(data);

            int entryPrimary = data.DocEntry;
            int entrySecondary = GetRelatedOrder(entryPrimary);

            data.DocEntry = entrySecondary;
            data.RelatedOrder = entryPrimary.ToString();
            UpdateSecondaryOrder(data);

            return (entryPrimary, entrySecondary);
        }

        private int CreatePrincipalOrder(SalesOrderFormModel data)
        {
            Documents primaryOrder = null;
            int docEntry = 0;
            try
            {
                primaryOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                primaryOrder.CardCode = data.CardCode;

                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_SplitPercentage).Value = (double)data.SplitPercentage;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ImportedPercentage).Value = (double)data.ImportedPercentage;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_CategoryClient).Value = data.CategoryClient;

                var docDateText = data.DocDate;
                var docDueDateText = data.DocDueDate;
                var taxDateText = data.TaxDate;

                if (!string.IsNullOrWhiteSpace(docDateText))
                    primaryOrder.DocDate = ConverterService.GetDateTimeFromStringSAP(docDateText);
                if (!string.IsNullOrWhiteSpace(docDueDateText))
                    primaryOrder.DocDueDate = ConverterService.GetDateTimeFromStringSAP(docDueDateText);
                if (!string.IsNullOrWhiteSpace(taxDateText))
                    primaryOrder.TaxDate = ConverterService.GetDateTimeFromStringSAP(taxDateText);

                primaryOrder.Comments = data.Comments;
                primaryOrder.DiscountPercent = (double)data.TotalDiscountPercent;

                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_AssignedEntity).Value = data.AssignedEntity;

                var validLines = data.Lines
                    .Where(l => !string.IsNullOrWhiteSpace(l.ItemCode))
                    .ToList();

                for (int i = 0; i < validLines.Count; i++)
                {

                    var line = validLines[i];

                    decimal baseQty = line.Quantity;
                    var (qtyPrincipal, _) = CalculateQuantities(baseQty, data.SplitPercentage);

                    primaryOrder.Lines.ItemCode = line.ItemCode;
                    primaryOrder.Lines.Quantity = (double)qtyPrincipal;

                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        primaryOrder.Lines.WarehouseCode = whsCode;

                    primaryOrder.Lines.UnitPrice = (double)line.UnitPrice;

                    if (!string.IsNullOrWhiteSpace(line.UomEntry))
                        primaryOrder.Lines.UoMEntry = Convert.ToInt32(line.UomEntry);

                    primaryOrder.Lines.TaxCode = line.TaxCode;

                    if (data.IsImportOrder)
                        primaryOrder.Lines.DiscountPercent = (double)data.ImportLineDiscount;

                    if (!string.IsNullOrEmpty(line.AgrNo))
                        primaryOrder.Lines.AgreementNo = Convert.ToInt32(line.AgrNo);

                    if (i < validLines.Count)
                        primaryOrder.Lines.Add();
                }

                if (primaryOrder.Add() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al crear la Orden Principal. {errCode} - {errMsg}");
                }

                string newKey = ConnectionSDK.DIAPI.GetNewObjectKey();
                docEntry = int.Parse(newKey);
            }
            finally
            {
                if (primaryOrder != null)
                    Marshal.ReleaseComObject(primaryOrder);
            }
            return docEntry;
        }

        private int CreateSecondaryOrder(SalesOrderFormModel data)
        {
            Documents secondaryOrder = null;
            int docEntry = 0;
            try
            {
                secondaryOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);

                secondaryOrder.CardCode = data.CardCode;

                double splitPercentage = (double)(100m - data.SplitPercentage);
                //if(splitPercentage != 75)
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_SplitPercentage).Value = splitPercentage;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ImportedPercentage).Value = (double)data.ImportedPercentage;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_CategoryClient).Value = data.CategoryClient;


                var docDateText = data.DocDate;
                var docDueDateText = data.DocDueDate;
                var taxDateText = data.TaxDate;

                if (!string.IsNullOrWhiteSpace(docDateText))
                    secondaryOrder.DocDate = ConverterService.GetDateTimeFromStringSAP(docDateText);
                if (!string.IsNullOrWhiteSpace(docDueDateText))
                    secondaryOrder.DocDueDate = ConverterService.GetDateTimeFromStringSAP(docDueDateText);
                if (!string.IsNullOrWhiteSpace(taxDateText))
                    secondaryOrder.TaxDate = ConverterService.GetDateTimeFromStringSAP(taxDateText);

                secondaryOrder.Comments = data.Comments;
                secondaryOrder.DiscountPercent = (double)data.TotalDiscountPercent;

                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_AssignedEntity).Value = data.AssignedEntity == Constants.FixedValues.EntityA ? Constants.FixedValues.EntityB : data.AssignedEntity;

                if (!string.IsNullOrEmpty(data.RelatedOrder))
                {
                    secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_RelatedOrder).Value = data.RelatedOrder;
                    secondaryOrder.DocumentReferences.ReferencedDocEntry = Convert.ToInt32(data.RelatedOrder);
                    secondaryOrder.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_SalesOrder;
                }

                string taxCode = AppConfig.Get(Constants.ConfigProps.TaxCodeSecondaryOrder);
                foreach (var line in data.Lines)
                {
                    if (string.IsNullOrWhiteSpace(line.ItemCode))
                        continue;

                    decimal baseQty = line.Quantity;
                    var (_, qtySecondary) = CalculateQuantities(baseQty, data.SplitPercentage);

                    secondaryOrder.Lines.ItemCode = line.ItemCode;
                    secondaryOrder.Lines.Quantity = (double)qtySecondary;

                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        secondaryOrder.Lines.WarehouseCode = whsCode;

                    secondaryOrder.Lines.Price = (double)line.UnitPrice;
                    secondaryOrder.Lines.UoMEntry = Convert.ToInt32(line.UomEntry);

                    secondaryOrder.Lines.TaxCode = taxCode;

                    if (!string.IsNullOrEmpty(line.AgrNo))
                        secondaryOrder.Lines.AgreementNo = Convert.ToInt32(line.AgrNo);

                    secondaryOrder.Lines.Add();

                }

                if (secondaryOrder.Add() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al crear la Orden Secundaria. {errCode} - {errMsg}");
                }

                string newKey = ConnectionSDK.DIAPI.GetNewObjectKey();
                docEntry = int.Parse(newKey);
            }
            finally
            {
                if (secondaryOrder != null)
                    Marshal.ReleaseComObject(secondaryOrder);
            }
            return docEntry;
        }

        private void UpdatePrincipalOrder(SalesOrderFormModel data)
        {
            Documents primaryOrder = null;

            if (data.DocEntry == 0) return;

            try
            {
                primaryOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);

                primaryOrder.GetByKey(data.DocEntry);

                var docDateText = data.DocDate;
                var docDueDateText = data.DocDueDate;
                var taxDateText = data.TaxDate;

                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ImportedPercentage).Value = (double)data.ImportedPercentage;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_CategoryClient).Value = data.CategoryClient;


                if (!string.IsNullOrWhiteSpace(docDateText))
                    primaryOrder.DocDate = ConverterService.GetDateTimeFromStringSAP(docDateText);
                if (!string.IsNullOrWhiteSpace(docDueDateText))
                    primaryOrder.DocDueDate = ConverterService.GetDateTimeFromStringSAP(docDueDateText);
                if (!string.IsNullOrWhiteSpace(taxDateText))
                    primaryOrder.TaxDate = ConverterService.GetDateTimeFromStringSAP(taxDateText);

                primaryOrder.Comments = data.Comments;
                primaryOrder.DiscountPercent = (double)data.TotalDiscountPercent;

                // ELIMINACION DE LINEAS - FUNCIONA BIEN
                var currLines = data.Lines.Select(l => l.LineId);
                for (int i = 0; i < primaryOrder.Lines.Count; i++)
                {
                    primaryOrder.Lines.SetCurrentLine(i);

                    if (!currLines.Contains(primaryOrder.Lines.LineNum))
                    {
                        primaryOrder.Lines.Delete();
                    }
                }

                // ACTUALIZACION DE LINEAS - FUNCIONA BIEN
                for (int i = 0; i < primaryOrder.Lines.Count; i++)
                {
                    primaryOrder.Lines.SetCurrentLine(i);

                    if (currLines.Contains(primaryOrder.Lines.LineNum))
                    {
                        var line = data.Lines[i];
                        if (string.IsNullOrWhiteSpace(line.ItemCode))
                            continue;

                        primaryOrder.Lines.ItemCode = line.ItemCode;

                        double prevQty = primaryOrder.Lines.Quantity;
                        decimal currQty = line.Quantity;
                        if (prevQty != (double)currQty)
                        {
                            var (qtyPrincipal, _) = CalculateQuantities(currQty, data.SplitPercentage);
                            primaryOrder.Lines.Quantity = (double)qtyPrincipal;
                        }

                        var whsCode = line.WhsCode;
                        if (!string.IsNullOrWhiteSpace(whsCode))
                            primaryOrder.Lines.WarehouseCode = whsCode;

                        primaryOrder.Lines.UnitPrice = (double)line.UnitPrice;

                        primaryOrder.Lines.TaxCode = line.TaxCode;

                        if (line.LineStatus == "C" && primaryOrder.Lines.LineStatus != BoStatus.bost_Close)
                            primaryOrder.Lines.LineStatus = BoStatus.bost_Close;
                    }
                }

                // AGREGAR NUEVAS LINEAS - FUNCIONA BIEN
                for (int i = 0; i < data.Lines.Count; i++)
                {
                    try
                    {
                        primaryOrder.Lines.SetCurrentLine(i);
                    }
                    catch
                    {
                        primaryOrder.Lines.Add();
                        primaryOrder.Lines.SetCurrentLine(primaryOrder.Lines.Count - 1);

                        var line = data.Lines[i];
                        if (string.IsNullOrWhiteSpace(line.ItemCode))
                            continue;

                        primaryOrder.Lines.ItemCode = line.ItemCode;

                        decimal currQty = line.Quantity;
                        var (qtyPrincipal, _) = CalculateQuantities(currQty, data.SplitPercentage);
                        primaryOrder.Lines.Quantity = (double)qtyPrincipal;


                        var whsCode = line.WhsCode;
                        if (!string.IsNullOrWhiteSpace(whsCode))
                            primaryOrder.Lines.WarehouseCode = whsCode;

                        primaryOrder.Lines.UnitPrice = (double)line.UnitPrice;

                        primaryOrder.Lines.TaxCode = line.TaxCode;

                        //if (line.LineStatus == "C")
                        //{
                        //    primaryOrder.Lines.LineStatus = BoStatus.bost_Close;
                        //    if (primaryOrder.Update() != 0)
                        //    {
                        //        ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                        //        throw new Exception($"Error al actualizar la linea de la Orden Principal. {errCode} - {errMsg}");
                        //    }
                        //}

                    }
                }

                if (primaryOrder.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al actualizar la Orden Principal. {errCode} - {errMsg}");
                }
            }
            finally
            {
                if (primaryOrder != null)
                    Marshal.ReleaseComObject(primaryOrder);
            }
        }

        private void UpdateSecondaryOrder(SalesOrderFormModel data)
        {
            Documents secondaryOrder = null;
            Documents primaryOrder = null;

            if (data.DocEntry == 0) return;

            try
            {
                secondaryOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                primaryOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);


                if (!secondaryOrder.GetByKey(data.DocEntry)) return;
                    //throw new Exception("No se encontró la orden secundaria a actualizar.");

                if (!string.IsNullOrEmpty(data.RelatedOrder))
                    primaryOrder.GetByKey(Convert.ToInt32(data.RelatedOrder));


                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ImportedPercentage).Value = (double)data.ImportedPercentage;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_CategoryClient).Value = data.CategoryClient;

                var docDateText = data.DocDate;
                var docDueDateText = data.DocDueDate;
                var taxDateText = data.TaxDate;


                if (!string.IsNullOrWhiteSpace(docDateText))
                    secondaryOrder.DocDate = ConverterService.GetDateTimeFromStringSAP(docDateText);
                if (!string.IsNullOrWhiteSpace(docDueDateText))
                    secondaryOrder.DocDueDate = ConverterService.GetDateTimeFromStringSAP(docDueDateText);
                if (!string.IsNullOrWhiteSpace(taxDateText))
                    secondaryOrder.TaxDate = ConverterService.GetDateTimeFromStringSAP(taxDateText);

                secondaryOrder.Comments = data.Comments;
                secondaryOrder.DiscountPercent = (double)data.TotalDiscountPercent;

                string taxCodeSecondary = AppConfig.Get(Constants.ConfigProps.TaxCodeSecondaryOrder);

                // ELIMINACION DE LINEAS - FUNCIONA BIEN
                var currLines = data.Lines.Select(l => l.LineId);
                for (int i = 0; i < secondaryOrder.Lines.Count; i++)
                {
                    secondaryOrder.Lines.SetCurrentLine(i);

                    if (!currLines.Contains(secondaryOrder.Lines.LineNum))
                    {
                        secondaryOrder.Lines.Delete();
                    }
                }

                // ACTUALIZACION DE LINEAS - FUNCIONA BIEN
                for (int i = 0; i < secondaryOrder.Lines.Count; i++)
                {
                    secondaryOrder.Lines.SetCurrentLine(i);

                    if (currLines.Contains(secondaryOrder.Lines.LineNum))
                    {
                        var line = data.Lines[i];
                        if (string.IsNullOrWhiteSpace(line.ItemCode))
                            continue;

                        secondaryOrder.Lines.ItemCode = line.ItemCode;

                        double prevQty;
                        if (!string.IsNullOrEmpty(data.RelatedOrder))
                        {

                            primaryOrder.Lines.SetCurrentLine(i);
                            prevQty = primaryOrder.Lines.Quantity;
                        }
                        else
                        {
                            prevQty = secondaryOrder.Lines.Quantity;
                        }

                        decimal currQty = line.Quantity;
                        if (prevQty != (double)currQty)
                        {
                            var (_, qtySecondary) = CalculateQuantities(currQty, data.SplitPercentage);
                            secondaryOrder.Lines.Quantity = (double)qtySecondary;
                        }

                        var whsCode = line.WhsCode;
                        if (!string.IsNullOrWhiteSpace(whsCode))
                            secondaryOrder.Lines.WarehouseCode = whsCode;

                        secondaryOrder.Lines.UnitPrice = (double)line.UnitPrice;

                        secondaryOrder.Lines.TaxCode = taxCodeSecondary;

                        //if (line.LineStatus == "C" && secondaryOrder.Lines.LineStatus != BoStatus.bost_Close)
                        //{
                        //    secondaryOrder.Lines.LineStatus = BoStatus.bost_Close;
                        //    if (secondaryOrder.Update() != 0)
                        //    {
                        //        ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                        //        throw new Exception($"Error al actualizar la linea de la Orden Secundaria. {errCode} - {errMsg}");
                        //    }
                        //}

                    }
                }

                // AGREGAR NUEVAS LINEAS - FUNCIONA BIEN
                for (int i = 0; i < data.Lines.Count; i++)
                {
                    try
                    {
                        secondaryOrder.Lines.SetCurrentLine(i);
                    }
                    catch
                    {
                        secondaryOrder.Lines.Add();
                        secondaryOrder.Lines.SetCurrentLine(secondaryOrder.Lines.Count - 1);

                        var line = data.Lines[i];
                        if (string.IsNullOrWhiteSpace(line.ItemCode))
                            continue;

                        secondaryOrder.Lines.ItemCode = line.ItemCode;

                        decimal currQty = line.Quantity;
                        var (_, qtySecondary) = CalculateQuantities(currQty, data.SplitPercentage);
                        secondaryOrder.Lines.Quantity = (double)qtySecondary;

                        var whsCode = line.WhsCode;
                        if (!string.IsNullOrWhiteSpace(whsCode))
                            secondaryOrder.Lines.WarehouseCode = whsCode;

                        secondaryOrder.Lines.UnitPrice = (double)line.UnitPrice;

                        secondaryOrder.Lines.TaxCode = taxCodeSecondary;

                        if (line.LineStatus == "C")
                            secondaryOrder.Lines.LineStatus = BoStatus.bost_Close;

                    }
                }

                if (secondaryOrder.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al actualizar la Orden Secundaria. {errCode} - {errMsg}");
                }
            }
            finally
            {
                if (secondaryOrder != null) Marshal.ReleaseComObject(secondaryOrder);
                if (primaryOrder != null) Marshal.ReleaseComObject(primaryOrder);
            }
        }

        private void UpdateRelatedOrderInPrimaryOrder(int primaryDocEntry, int secondaryDocEntry)
        {
            Documents primaryOrder = null;
            try
            {
                primaryOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);

                if (!primaryOrder.GetByKey(primaryDocEntry))
                    throw new Exception("No se encontró la orden principal para actualizar el campo RelatedOrder.");

                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_RelatedOrder).Value = secondaryDocEntry.ToString();

                if (primaryOrder.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al actualizar el RelatedOrder en la Orden Principal. {errCode} - {errMsg}");
                }
            }
            finally
            {
                if (primaryOrder != null)
                    Marshal.ReleaseComObject(primaryOrder);
            }
        }

    }
}
