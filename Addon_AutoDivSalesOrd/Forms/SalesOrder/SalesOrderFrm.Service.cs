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
                string totalDiscPerc = data.TotalDiscountPercent.ToString().Replace(",",".");
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ItpsDiscount).Value = totalDiscPerc;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_AssignedEntity).Value = Constants.FixedValues.EntityA; // data.AssignedEntity;

                if (data.GlobalAgreement > 0) 
                    primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_GlobalAgree).Value = data.GlobalAgreement;

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
                primaryOrder.PaymentGroupCode = data.PaymentGroupCode;
                primaryOrder.Address2 = data.Address2;
                primaryOrder.ShipToCode = data.ShipToCode;
                primaryOrder.Address = data.Address;
                primaryOrder.PayToCode = data.PayToCode;

                // Dirección Facturación:
                primaryOrder.AddressExtension.BillToStreet = data.StreetB;
                primaryOrder.AddressExtension.BillToStreetNo = data.StreetNoB;
                primaryOrder.AddressExtension.BillToBlock = data.BlockB;
                primaryOrder.AddressExtension.BillToCity = data.CityB;
                primaryOrder.AddressExtension.BillToZipCode = data.ZipCodeB;
                primaryOrder.AddressExtension.BillToCounty = data.CountyB;
                primaryOrder.AddressExtension.BillToState = data.StateB;
                primaryOrder.AddressExtension.BillToCountry = data.CountryB;
                primaryOrder.AddressExtension.BillToBuilding = data.BuildingB; 
                primaryOrder.AddressExtension.BillToAddress2 = data.Address2B; 
                primaryOrder.AddressExtension.BillToAddress3 = data.Address3B; 
                primaryOrder.AddressExtension.BillToGlobalLocationNumber = data.GlbLocNumB; 
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaB").Value = data.TransportistaB; 
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneB").Value = data.DeliveryZoneB;

                // Dirección Entrega:
                primaryOrder.AddressExtension.ShipToStreet = data.StreetS;
                primaryOrder.AddressExtension.ShipToStreetNo = data.StreetNoS;
                primaryOrder.AddressExtension.ShipToBlock = data.BlockS;
                primaryOrder.AddressExtension.ShipToCity = data.CityS;
                primaryOrder.AddressExtension.ShipToZipCode = data.ZipCodeS;
                primaryOrder.AddressExtension.ShipToCounty = data.CountyS;
                primaryOrder.AddressExtension.ShipToState = data.StateS;
                primaryOrder.AddressExtension.ShipToCountry = data.CountryS;
                primaryOrder.AddressExtension.ShipToBuilding = data.BuildingS;
                primaryOrder.AddressExtension.ShipToAddress2 = data.Address2S;
                primaryOrder.AddressExtension.ShipToAddress3 = data.Address3S;
                primaryOrder.AddressExtension.ShipToGlobalLocationNumber = data.GlbLocNumS;
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaS").Value = data.TransportistaS;
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneS").Value = data.DeliveryZoneS;

                

                var validLines = data.Lines
                    .Where(l => !string.IsNullOrWhiteSpace(l.ItemCode))
                    .ToList();

                if (validLines.Count == 0)
                    throw new Exception("No hay líneas válidas para crear la Orden Principal.");

                var invalidQtyPrincipal = validLines.FirstOrDefault(l => l.Quantity <= 0m);
                if (invalidQtyPrincipal != null)
                    throw new Exception($"Cantidad inválida en Orden Principal para el artículo {invalidQtyPrincipal.ItemCode}.");

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

                    if (int.TryParse(line.UomEntry, out int uomEntry) && uomEntry > 0)
                        primaryOrder.Lines.UoMEntry = uomEntry;

                    if (!string.IsNullOrWhiteSpace(line.TaxCode))
                        primaryOrder.Lines.TaxCode = line.TaxCode;

                    if (data.IsImportOrder && (double)data.ImportLineDiscount != 100)
                    {
                        primaryOrder.Lines.DiscountPercent = (double)data.ImportLineDiscount;
                    }
                    else
                    {
                        primaryOrder.Lines.DiscountPercent = (double)line.Discount;
                    }

                    if (int.TryParse(line.AgrNo, out int agreementNo) && agreementNo > 0)
                        primaryOrder.Lines.AgreementNo = agreementNo;

                    if (i < validLines.Count - 1)
                        primaryOrder.Lines.Add();
                }

                if (primaryOrder.Add() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    var detail = string.Join(" | ", validLines.Select((l, idx) =>
                        $"L{idx + 1}:Item={l.ItemCode},Qty={l.Quantity},UomEntry={l.UomEntry},Whs={l.WhsCode},Tax={l.TaxCode},Price={l.UnitPrice},Agr={l.AgrNo}"));
                    throw new Exception($"Error al crear la Orden Principal. {errCode} - {errMsg}.");
                }

                string newKey = ConnectionSDK.DIAPI.GetNewObjectKey();
                docEntry = int.Parse(newKey);
            }
            catch(Exception ex)
            {
                NotificationService.Error(ex.Message);
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
                string totalDiscPerc = data.TotalDiscountPercent.ToString().Replace(",",".");
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ItpsDiscount).Value = totalDiscPerc;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_GlobalAgree).Value = data.GlobalAgreement;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_AssignedEntity).Value = Constants.FixedValues.EntityB;

                var docDateText = data.DocDate;
                var docDueDateText = data.DocDueDate;
                var taxDateText = data.TaxDate;

                if (!string.IsNullOrWhiteSpace(docDateText))
                    secondaryOrder.DocDate = ConverterService.GetDateTimeFromStringSAP(docDateText);
                if (!string.IsNullOrWhiteSpace(docDueDateText))
                    secondaryOrder.DocDueDate = ConverterService.GetDateTimeFromStringSAP(docDueDateText);
                if (!string.IsNullOrWhiteSpace(taxDateText))
                    secondaryOrder.TaxDate = ConverterService.GetDateTimeFromStringSAP(taxDateText);

                secondaryOrder.Indicator = "NL"; // indicado no libro

                secondaryOrder.Comments = data.Comments;
                secondaryOrder.DiscountPercent = (double)data.TotalDiscountPercent;
                secondaryOrder.PaymentGroupCode = data.PaymentGroupCode;
                secondaryOrder.Address2 = data.Address2;
                secondaryOrder.ShipToCode = data.ShipToCode;
                secondaryOrder.Address = data.Address;
                secondaryOrder.PayToCode = data.PayToCode;

                // Dirección Facturación:
                secondaryOrder.AddressExtension.BillToStreet = data.StreetB;
                secondaryOrder.AddressExtension.BillToStreetNo = data.StreetNoB;
                secondaryOrder.AddressExtension.BillToBlock = data.BlockB;
                secondaryOrder.AddressExtension.BillToCity = data.CityB;
                secondaryOrder.AddressExtension.BillToZipCode = data.ZipCodeB;
                secondaryOrder.AddressExtension.BillToCounty = data.CountyB;
                secondaryOrder.AddressExtension.BillToState = data.StateB;
                secondaryOrder.AddressExtension.BillToCountry = data.CountryB;
                secondaryOrder.AddressExtension.BillToBuilding = data.BuildingB;
                secondaryOrder.AddressExtension.BillToAddress2 = data.Address2B;
                secondaryOrder.AddressExtension.BillToAddress3 = data.Address3B;
                secondaryOrder.AddressExtension.BillToGlobalLocationNumber = data.GlbLocNumB;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaB").Value = data.TransportistaB;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneB").Value = data.DeliveryZoneB;

                // Dirección Entrega:
                secondaryOrder.AddressExtension.ShipToStreet = data.StreetS;
                secondaryOrder.AddressExtension.ShipToStreetNo = data.StreetNoS;
                secondaryOrder.AddressExtension.ShipToBlock = data.BlockS;
                secondaryOrder.AddressExtension.ShipToCity = data.CityS;
                secondaryOrder.AddressExtension.ShipToZipCode = data.ZipCodeS;
                secondaryOrder.AddressExtension.ShipToCounty = data.CountyS;
                secondaryOrder.AddressExtension.ShipToState = data.StateS;
                secondaryOrder.AddressExtension.ShipToCountry = data.CountryS;
                secondaryOrder.AddressExtension.ShipToBuilding = data.BuildingS;
                secondaryOrder.AddressExtension.ShipToAddress2 = data.Address2S;
                secondaryOrder.AddressExtension.ShipToAddress3 = data.Address3S;
                secondaryOrder.AddressExtension.ShipToGlobalLocationNumber = data.GlbLocNumS;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaS").Value = data.TransportistaS;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneS").Value = data.DeliveryZoneS;



                if (!string.IsNullOrEmpty(data.RelatedOrder))
                {
                    secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_RelatedOrder).Value = data.RelatedOrder;
                    secondaryOrder.DocumentReferences.ReferencedDocEntry = Convert.ToInt32(data.RelatedOrder);
                    secondaryOrder.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_SalesOrder;
                }

                string taxCode = AppConfig.Get(Constants.ConfigProps.TaxCodeSecondaryOrder);
                var validSecondaryLines = data.Lines
                    .Where(l => !string.IsNullOrWhiteSpace(l.ItemCode))
                    .ToList();

                if (validSecondaryLines.Count == 0)
                    throw new Exception("No hay líneas válidas para crear la Orden Secundaria.");

                for (int i = 0; i < validSecondaryLines.Count; i++)
                {
                    var line = validSecondaryLines[i];

                    if (line.Quantity <= 0m)
                        throw new Exception($"Cantidad inválida en Orden Secundaria para el artículo {line.ItemCode}.");

                    decimal baseQty = line.Quantity;
                    var (_, qtySecondary) = CalculateQuantities(baseQty, data.SplitPercentage);

                    secondaryOrder.Lines.ItemCode = line.ItemCode;
                    secondaryOrder.Lines.Quantity = (double)qtySecondary;

                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        secondaryOrder.Lines.WarehouseCode = whsCode;

                   // secondaryOrder.Lines.Price = (double)line.UnitPrice;

                    secondaryOrder.Lines.UnitPrice = (double)line.UnitPrice;

                    if (int.TryParse(line.UomEntry, out int uomEntry) && uomEntry > 0)
                        secondaryOrder.Lines.UoMEntry = uomEntry;

                    secondaryOrder.Lines.TaxCode = taxCode;

                    if (int.TryParse(line.AgrNo, out int agreementNo) && agreementNo > 0)
                        secondaryOrder.Lines.AgreementNo = agreementNo;

                    if (i < validSecondaryLines.Count - 1)
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

                if (!primaryOrder.GetByKey(data.DocEntry)) return;

                var docDateText = data.DocDate;
                var docDueDateText = data.DocDueDate;
                var taxDateText = data.TaxDate;

                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_SplitPercentage).Value = (double)data.SplitPercentage;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ImportedPercentage).Value = (double)data.ImportedPercentage;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_CategoryClient).Value = data.CategoryClient;
                //primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_AssignedEntity).Value = data.AssignedEntity;
                string totalDiscPerc = data.TotalDiscountPercent.ToString().Replace(",", ".");
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ItpsDiscount).Value = totalDiscPerc;
                primaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_GlobalAgree).Value = data.GlobalAgreement;

                if (!string.IsNullOrWhiteSpace(docDateText))
                    primaryOrder.DocDate = ConverterService.GetDateTimeFromStringSAP(docDateText);
                if (!string.IsNullOrWhiteSpace(docDueDateText))
                    primaryOrder.DocDueDate = ConverterService.GetDateTimeFromStringSAP(docDueDateText);
                if (!string.IsNullOrWhiteSpace(taxDateText))
                    primaryOrder.TaxDate = ConverterService.GetDateTimeFromStringSAP(taxDateText);

                primaryOrder.Comments = data.Comments;
                primaryOrder.DiscountPercent = (double)data.TotalDiscountPercent;
                primaryOrder.PaymentGroupCode = data.PaymentGroupCode;
                primaryOrder.Address2 = data.Address2;
                primaryOrder.ShipToCode = data.ShipToCode;
                primaryOrder.Address = data.Address;
                primaryOrder.PayToCode = data.PayToCode;

                // Dirección Facturación:
                primaryOrder.AddressExtension.BillToStreet = data.StreetB;
                primaryOrder.AddressExtension.BillToStreetNo = data.StreetNoB;
                primaryOrder.AddressExtension.BillToBlock = data.BlockB;
                primaryOrder.AddressExtension.BillToCity = data.CityB;
                primaryOrder.AddressExtension.BillToZipCode = data.ZipCodeB;
                primaryOrder.AddressExtension.BillToCounty = data.CountyB;
                primaryOrder.AddressExtension.BillToState = data.StateB;
                primaryOrder.AddressExtension.BillToCountry = data.CountryB;
                primaryOrder.AddressExtension.BillToBuilding = data.BuildingB;
                primaryOrder.AddressExtension.BillToAddress2 = data.Address2B;
                primaryOrder.AddressExtension.BillToAddress3 = data.Address3B;
                primaryOrder.AddressExtension.BillToGlobalLocationNumber = data.GlbLocNumB;
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaB").Value = data.TransportistaB;
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneB").Value = data.DeliveryZoneB;

                // Dirección Entrega:
                primaryOrder.AddressExtension.ShipToStreet = data.StreetS;
                primaryOrder.AddressExtension.ShipToStreetNo = data.StreetNoS;
                primaryOrder.AddressExtension.ShipToBlock = data.BlockS;
                primaryOrder.AddressExtension.ShipToCity = data.CityS;
                primaryOrder.AddressExtension.ShipToZipCode = data.ZipCodeS;
                primaryOrder.AddressExtension.ShipToCounty = data.CountyS;
                primaryOrder.AddressExtension.ShipToState = data.StateS;
                primaryOrder.AddressExtension.ShipToCountry = data.CountryS;
                primaryOrder.AddressExtension.ShipToBuilding = data.BuildingS;
                primaryOrder.AddressExtension.ShipToAddress2 = data.Address2S;
                primaryOrder.AddressExtension.ShipToAddress3 = data.Address3S;
                primaryOrder.AddressExtension.ShipToGlobalLocationNumber = data.GlbLocNumS;
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaS").Value = data.TransportistaS;
                primaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneS").Value = data.DeliveryZoneS;

                var validLines = data.Lines
                    .Where(l => !string.IsNullOrWhiteSpace(l.ItemCode))
                    .ToList();

                var linesByLineNum = validLines
                    .GroupBy(l => l.LineId)
                    .ToDictionary(g => g.Key, g => g.First());

                var lineNumsToClose = new HashSet<int>(validLines
                    .Where(l => l.LineStatus == "C" && l.LineId >= 0)
                    .Select(l => l.LineId));

                // ELIMINAR LINEAS
                for (int i = primaryOrder.Lines.Count - 1; i >= 0; i--)
                {
                    primaryOrder.Lines.SetCurrentLine(i);
                    int lineNum = primaryOrder.Lines.LineNum;

                    if (!linesByLineNum.ContainsKey(lineNum))
                        primaryOrder.Lines.Delete();
                }

                var existingLineNums = new HashSet<int>();

                // ACTUALIZAR LINEAS
                for (int i = 0; i < primaryOrder.Lines.Count; i++)
                {
                    primaryOrder.Lines.SetCurrentLine(i);
                    int lineNum = primaryOrder.Lines.LineNum;
                    existingLineNums.Add(lineNum);

                    if (!linesByLineNum.TryGetValue(lineNum, out var line))
                        continue;

                    if (line.LineStatus == "C")
                        continue;

                    primaryOrder.Lines.ItemCode = line.ItemCode;

                    decimal currQty = line.Quantity;
                    decimal currentQtyPrincipal = decimal.Round((decimal)primaryOrder.Lines.Quantity, 2, MidpointRounding.AwayFromZero);
                    if (currentQtyPrincipal != decimal.Round(currQty, 2, MidpointRounding.AwayFromZero))
                    {
                        var (qtyPrincipal, _) = CalculateQuantities(currQty, data.SplitPercentage);
                        primaryOrder.Lines.Quantity = (double)qtyPrincipal;
                    }

                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        primaryOrder.Lines.WarehouseCode = whsCode;

                    primaryOrder.Lines.UnitPrice = (double)line.UnitPrice;
                    primaryOrder.Lines.TaxCode = line.TaxCode;

                    if (int.TryParse(line.UomEntry, out int uomEntry) && uomEntry > 0)
                        primaryOrder.Lines.UoMEntry = uomEntry;

                    if (int.TryParse(line.AgrNo, out int agreementNo) && agreementNo > 0)
                        primaryOrder.Lines.AgreementNo = agreementNo;

                    if (data.IsImportOrder && (double)data.ImportLineDiscount != 100)
                        primaryOrder.Lines.DiscountPercent = (double)data.ImportLineDiscount;
                    else
                        primaryOrder.Lines.DiscountPercent = (double)line.Discount;
                }


                // AGREGAR LINEAS
                foreach (var line in validLines)
                {
                    if (existingLineNums.Contains(line.LineId))
                        continue;

                    primaryOrder.Lines.Add();
                    primaryOrder.Lines.SetCurrentLine(primaryOrder.Lines.Count - 1);

                    primaryOrder.Lines.ItemCode = line.ItemCode;

                    var (qtyPrincipal, _) = CalculateQuantities(line.Quantity, data.SplitPercentage);
                    primaryOrder.Lines.Quantity = (double)qtyPrincipal;

                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        primaryOrder.Lines.WarehouseCode = whsCode;

                    primaryOrder.Lines.UnitPrice = (double)line.UnitPrice;
                    primaryOrder.Lines.TaxCode = line.TaxCode;

                    if (int.TryParse(line.UomEntry, out int uomEntry) && uomEntry > 0)
                        primaryOrder.Lines.UoMEntry = uomEntry;

                    if (int.TryParse(line.AgrNo, out int agreementNo) && agreementNo > 0)
                        primaryOrder.Lines.AgreementNo = agreementNo;

                    if (data.IsImportOrder && (double)data.ImportLineDiscount != 100)
                        primaryOrder.Lines.DiscountPercent = (double)data.ImportLineDiscount;
                    else
                        primaryOrder.Lines.DiscountPercent = (double)line.Discount;
                }

                if (primaryOrder.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al actualizar la Orden Principal. {errCode} - {errMsg}");
                }

                if (lineNumsToClose.Count > 0)
                {
                    if (!primaryOrder.GetByKey(data.DocEntry)) return;

                    bool needsCloseUpdate = false;

                    for (int i = 0; i < primaryOrder.Lines.Count; i++)
                    {
                        primaryOrder.Lines.SetCurrentLine(i);

                        if (!lineNumsToClose.Contains(primaryOrder.Lines.LineNum))
                            continue;

                        if (primaryOrder.Lines.LineStatus != BoStatus.bost_Close)
                        {
                            primaryOrder.Lines.LineStatus = BoStatus.bost_Close;
                            needsCloseUpdate = true;
                        }
                    }

                    if (needsCloseUpdate && primaryOrder.Update() != 0)
                    {
                        ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Error al actualizar el cierre de líneas de la Orden Principal. {errCode} - {errMsg}");
                    }
                }

                if (!string.IsNullOrEmpty(data.RelatedOrder))
                    SyncLineStatusToRelatedOrder(data.DocEntry, Convert.ToInt32(data.RelatedOrder));
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

                primaryOrder.GetByKey(Convert.ToInt32(data.RelatedOrder));


                if (!secondaryOrder.GetByKey(data.DocEntry)) return;

                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_SplitPercentage).Value = (double)(100m - data.SplitPercentage);
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ImportedPercentage).Value = (double)data.ImportedPercentage;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_CategoryClient).Value = data.CategoryClient;
                //secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_AssignedEntity).Value = data.AssignedEntity == Constants.FixedValues.EntityA ? Constants.FixedValues.EntityB : data.AssignedEntity;
                string totalDiscPerc = data.TotalDiscountPercent.ToString().Replace(",", ".");
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_ItpsDiscount).Value = totalDiscPerc;
                secondaryOrder.UserFields.Fields.Item(Constants.SalesOrder_Fields.Head_GlobalAgree).Value = data.GlobalAgreement;

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
                secondaryOrder.PaymentGroupCode = data.PaymentGroupCode;
                secondaryOrder.Address2 = data.Address2;
                secondaryOrder.ShipToCode = data.ShipToCode;
                secondaryOrder.Address = data.Address;
                secondaryOrder.PayToCode = data.PayToCode;

                // Dirección Facturación:
                secondaryOrder.AddressExtension.BillToStreet = data.StreetB;
                secondaryOrder.AddressExtension.BillToStreetNo = data.StreetNoB;
                secondaryOrder.AddressExtension.BillToBlock = data.BlockB;
                secondaryOrder.AddressExtension.BillToCity = data.CityB;
                secondaryOrder.AddressExtension.BillToZipCode = data.ZipCodeB;
                secondaryOrder.AddressExtension.BillToCounty = data.CountyB;
                secondaryOrder.AddressExtension.BillToState = data.StateB;
                secondaryOrder.AddressExtension.BillToCountry = data.CountryB;
                secondaryOrder.AddressExtension.BillToBuilding = data.BuildingB;
                secondaryOrder.AddressExtension.BillToAddress2 = data.Address2B;
                secondaryOrder.AddressExtension.BillToAddress3 = data.Address3B;
                secondaryOrder.AddressExtension.BillToGlobalLocationNumber = data.GlbLocNumB;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaB").Value = data.TransportistaB;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneB").Value = data.DeliveryZoneB;

                // Dirección Entrega:
                secondaryOrder.AddressExtension.ShipToStreet = data.StreetS;
                secondaryOrder.AddressExtension.ShipToStreetNo = data.StreetNoS;
                secondaryOrder.AddressExtension.ShipToBlock = data.BlockS;
                secondaryOrder.AddressExtension.ShipToCity = data.CityS;
                secondaryOrder.AddressExtension.ShipToZipCode = data.ZipCodeS;
                secondaryOrder.AddressExtension.ShipToCounty = data.CountyS;
                secondaryOrder.AddressExtension.ShipToState = data.StateS;
                secondaryOrder.AddressExtension.ShipToCountry = data.CountryS;
                secondaryOrder.AddressExtension.ShipToBuilding = data.BuildingS;
                secondaryOrder.AddressExtension.ShipToAddress2 = data.Address2S;
                secondaryOrder.AddressExtension.ShipToAddress3 = data.Address3S;
                secondaryOrder.AddressExtension.ShipToGlobalLocationNumber = data.GlbLocNumS;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_TransportistaS").Value = data.TransportistaS;
                secondaryOrder.AddressExtension.UserFields.Fields.Item("U_ITPS_DeliveryZoneS").Value = data.DeliveryZoneS;

                string taxCodeSecondary = AppConfig.Get(Constants.ConfigProps.TaxCodeSecondaryOrder);
                var validLines = data.Lines
                    .Where(l => !string.IsNullOrWhiteSpace(l.ItemCode))
                    .ToList();

                var linesByLineNum = validLines
                    .GroupBy(l => l.LineId)
                    .ToDictionary(g => g.Key, g => g.First());

                var lineNumsToClose = new HashSet<int>(validLines
                    .Where(l => l.LineStatus == "C" && l.LineId >= 0)
                    .Select(l => l.LineId));                

                for (int i = secondaryOrder.Lines.Count - 1; i >= 0; i--)
                {
                    secondaryOrder.Lines.SetCurrentLine(i);
                    int lineNumB = secondaryOrder.Lines.LineNum;

                    if (!linesByLineNum.ContainsKey(lineNumB))
                        secondaryOrder.Lines.Delete();
                }

                var existingLineNumsB = new HashSet<int>();

                for (int i = 0; i < secondaryOrder.Lines.Count; i++)
                {
                    secondaryOrder.Lines.SetCurrentLine(i);
                    int lineNumB = secondaryOrder.Lines.LineNum;
                    existingLineNumsB.Add(lineNumB);

                    if (!linesByLineNum.TryGetValue(lineNumB, out var line))
                        continue;

                    if (line.LineStatus == "C")
                        continue;

                    secondaryOrder.Lines.ItemCode = line.ItemCode;

                    // Verificamos si cambio la cantidad de la orden principal para aplicar o no la nueva cantidad en split B
                    for(int j = 0; j < primaryOrder.Lines.Count ; j++)
                    {
                        primaryOrder.Lines.SetCurrentLine(j);
                        if (primaryOrder.Lines.LineNum == secondaryOrder.Lines.LineNum)
                        {
                            decimal oldQtyPrimary = decimal.Round((decimal)primaryOrder.Lines.Quantity, 2, MidpointRounding.AwayFromZero);
                            decimal currQtyPrimary = decimal.Round((decimal)line.Quantity, 2, MidpointRounding.AwayFromZero);
                            if (oldQtyPrimary != currQtyPrimary)
                            {
                                var (_, qtySecondary) = CalculateQuantities(line.Quantity, data.SplitPercentage);
                                secondaryOrder.Lines.Quantity = (double)qtySecondary;
                            }
                        }
                    }


                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        secondaryOrder.Lines.WarehouseCode = whsCode;

                    secondaryOrder.Lines.UnitPrice = (double)line.UnitPrice;
                    secondaryOrder.Lines.TaxCode = taxCodeSecondary;

                    if (int.TryParse(line.UomEntry, out int uomEntry) && uomEntry > 0)
                        secondaryOrder.Lines.UoMEntry = uomEntry;

                    if (int.TryParse(line.AgrNo, out int agreementNo) && agreementNo > 0)
                        secondaryOrder.Lines.AgreementNo = agreementNo;
                }

                foreach (var line in validLines)
                {
                    if (existingLineNumsB.Contains(line.LineId))
                        continue;

                    secondaryOrder.Lines.Add();
                    secondaryOrder.Lines.SetCurrentLine(secondaryOrder.Lines.Count - 1);

                    secondaryOrder.Lines.ItemCode = line.ItemCode;

                    var (_, qtySecondary) = CalculateQuantities(line.Quantity, data.SplitPercentage);
                    secondaryOrder.Lines.Quantity = (double)qtySecondary;

                    var whsCode = line.WhsCode;
                    if (!string.IsNullOrWhiteSpace(whsCode))
                        secondaryOrder.Lines.WarehouseCode = whsCode;

                    secondaryOrder.Lines.UnitPrice = (double)line.UnitPrice;
                    secondaryOrder.Lines.TaxCode = taxCodeSecondary;

                    if (int.TryParse(line.UomEntry, out int uomEntry) && uomEntry > 0)
                        secondaryOrder.Lines.UoMEntry = uomEntry;

                    if (int.TryParse(line.AgrNo, out int agreementNo) && agreementNo > 0)
                        secondaryOrder.Lines.AgreementNo = agreementNo;
                }

                if (secondaryOrder.Update() != 0)
                {
                    ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Error al actualizar la Orden Secundaria. {errCode} - {errMsg}");
                }

                if (lineNumsToClose.Count > 0)
                {
                    if (!secondaryOrder.GetByKey(data.DocEntry)) return;

                    bool needsCloseUpdate = false;

                    for (int i = 0; i < secondaryOrder.Lines.Count; i++)
                    {
                        secondaryOrder.Lines.SetCurrentLine(i);

                        if (!lineNumsToClose.Contains(secondaryOrder.Lines.LineNum))
                            continue;

                        if (secondaryOrder.Lines.LineStatus != BoStatus.bost_Close)
                        {
                            secondaryOrder.Lines.LineStatus = BoStatus.bost_Close;
                            needsCloseUpdate = true;
                        }
                    }

                    if (needsCloseUpdate && secondaryOrder.Update() != 0)
                    {
                        ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Error al actualizar el cierre de líneas de la Orden Secundaria. {errCode} - {errMsg}");
                    }
                }

                if (!string.IsNullOrEmpty(data.RelatedOrder))
                    SyncLineStatusToRelatedOrder(Convert.ToInt32(data.RelatedOrder), data.DocEntry);
            }
            finally
            {
                if (secondaryOrder != null) Marshal.ReleaseComObject(secondaryOrder);
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

        /// <summary>
        /// Sincroniza el estado de cierre (LineStatus) de líneas entre dos órdenes relacionadas.
        /// Cuando una línea se cierra en la orden origen, se cierra también en la orden destino.
        /// </summary>
        /// <param name="sourceOrderEntry">DocEntry de la orden origen (la que fue actualizada)</param>
        /// <param name="targetOrderEntry">DocEntry de la orden destino (la que debe sincronizarse)</param>
        private void SyncLineStatusToRelatedOrder(int sourceOrderEntry, int targetOrderEntry)
        {
            Documents sourceOrder = null;
            Documents targetOrder = null;
            try
            {
                sourceOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);
                targetOrder = (Documents)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oOrders);

                if (!sourceOrder.GetByKey(sourceOrderEntry))
                    throw new Exception($"No se encontró la orden origen ({sourceOrderEntry}) para sincronizar líneas.");

                if (!targetOrder.GetByKey(targetOrderEntry))
                    throw new Exception($"No se encontró la orden destino ({targetOrderEntry}) para sincronizar líneas.");

                bool needsUpdate = false;

                // Recorrer líneas de la orden origen y aplicar cierre a la orden destino
                for (int i = 0; i < sourceOrder.Lines.Count; i++)
                {
                    sourceOrder.Lines.SetCurrentLine(i);

                    // Si la línea está cerrada en la orden origen
                    if (sourceOrder.Lines.LineStatus == BoStatus.bost_Close)
                    {
                        // Buscar la línea correspondiente en la orden destino por LineNum
                        for (int j = 0; j < targetOrder.Lines.Count; j++)
                        {
                            targetOrder.Lines.SetCurrentLine(j);
                            if (targetOrder.Lines.LineNum == sourceOrder.Lines.LineNum)
                            {
                                // Si no está cerrada, cerrarla
                                if (targetOrder.Lines.LineStatus != BoStatus.bost_Close)
                                {
                                    targetOrder.Lines.LineStatus = BoStatus.bost_Close;
                                    needsUpdate = true;
                                }
                                break;
                            }
                        }
                    }
                }

                // Si hubo cambios, actualizar la orden destino
                if (needsUpdate)
                {
                    if (targetOrder.Update() != 0)
                    {
                        ConnectionSDK.DIAPI.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Error al sincronizar el cierre de líneas en la orden relacionada ({targetOrderEntry}). {errCode} - {errMsg}");
                    }
                }
            }
            finally
            {
                if (sourceOrder != null)
                    Marshal.ReleaseComObject(sourceOrder);
                if (targetOrder != null)
                    Marshal.ReleaseComObject(targetOrder);
            }
        }

    }
}
