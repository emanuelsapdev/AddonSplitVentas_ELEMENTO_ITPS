namespace Addon_AutoDivSalesOrd.Common
{
    public static class Constants
    {
        public static class Addon
        {
            public const string Name = "Addon_AutoDivSalesOrd";
            public const string Identifier = "Addon_AutoDivSalesOrd";
        }

        public static class FormTypes
        {
            public const string SalesOrder = "139";
            public const string SalesQuote = "149";
            public const string SalesInvoice = "133";
            public const string GerentePicking = "81";
        }

        public static class MenusUID
        {
            public const string Cancel = "1284";
            public const string Close = "1286";
            public const string ModeAdd = "1282";  // BOTON PARA MODO CREACIÓN
            public const string OrdenVenta = "2050";  // ORDR
            public const string Duplicate = "1287";  // duplicar

            public const string RowInitial = "1290"; // flecha |<-
            public const string RowBefore = "1289"; // flecha <-
            public const string RowAfter = "1288"; // flecha ->
            public const string RowEnd = "1291"; // flecha ->|
            public const string RowsRefresh = "1304"; // update data
        }

        public static class ItemUids
        {
            public const string AddOrUpdateButton = "1";
        }

        public static class DataSources
        {
            public const string SalesOrderHeader = "ORDR";
            public const string SalesOrderLines = "RDR1";
            public const string SalesOrderAddresses = "RDR12";
        }

        public static class SalesOrder_Fields
        {
            public const string Head_CardCode = "CardCode";
            public const string Head_SplitPercentage = "U_ITPS_SplitPercentage";
            public const string Head_ImportedPercentage = "U_ITPS_ImportedPercentage";
            public const string Head_CategoryClient = "U_CATEGORIA_CLIENTE";
            public const string Head_AssignedEntity = "U_Tipo";
            public const string Head_GlobalAgree = "U_ITPS_NRO_AC";
            public const string Head_RelatedOrder = "U_ITPS_RelatedOrder";
            public const string Head_ItpsDiscount = "U_ITPS_DESCUENTO";
            public const string Head_DocDate = "DocDate";
            public const string Head_DocEntry = "DocEntry";
            public const string Head_DocDueDate = "DocDueDate";
            public const string Head_TaxDate = "TaxDate";
            public const string Head_Comments = "Comments";

            public const string Det_WhsCode = "WhsCode";
            public const string Det_ItemCode = "ItemCode";
            public const string Det_Quantity = "Quantity";
            public const string Det_UomCode = "UomCode";
            public const string Det_Discount = "DiscPrcnt";
            public const string Det_TaxCode = "TaxCode";
            public const string Det_LineNum = "LineNum";
            public const string Det_AgrNo = "AgrNo";
        }

        public static class SalesOrder_FieldsUIDs
        {
            public const string Head_CardCode = "4";
            public const string Head_CardName = "54";
            public const string Head_Matrix = "38";
            public const string Head_SplitPercentage = "U_ITPS_SplitPercentage";
            public const string Head_ImportedPercentage = "U_ITPS_ImportedPercentage";
            public const string Head_CategoryClient = "U_CATEGORIA_CLIENTE";
            public const string Head_AssignedEntity = "U_Tipo";
            public const string Head_RelatedOrder = "U_ITPS_RelatedOrder";
            public const string Head_ItpsDiscount = "U_ITPS_DESCUENTO";
            public const string Head_DocDate = "10";
            public const string Head_DocDueDate = "12";
            public const string Head_TaxDate = "46";
            public const string Head_Comments = "16";
            public const string Head_DiscPrcnt = "24";
            public const string Head_PaymentGroupCode = "47";
            public const string Head_Money = "70";
            public const string Head_Address2 = "92";
            public const string Head_ShipToCode = "40";
            public const string Head_Address = "6";
            public const string Head_PayToDate = "226";

            public const string Det_ItemCode = "1";
            public const string Det_Quantity = "11";
            public const string Det_Discount = "15";
            public const string Det_TaxCode = "160";
            public const string Det_UomCode = "1470002145";
            public const string Det_UomEntry = "1470002143";
            public const string Det_WhsCode = "24";
            public const string Det_Price = "14";
            public const string Det_LineNum = "0";
            public const string Det_LineId = "110";
            public const string Det_LineStatus = "40";
            public const string Det_AgrNo = "1250002129";

        }



        public static class GerentePicking_Fields
        {
            
        }

        
        public static class GerentePicking_LevelPanel
        {
            public const int OpenLevelPane = 1;
            
        }

        public static class GerentePicking_FieldsUIDs
        {
            public const string Head_BtnCancel = "2";
            public const string Head_MatrixOpen = "10";
            public const string Head_BtnProcess = "btnProcess";
            public const string Head_FolderRelease = "6";
            public const string Head_FolderOpen = "7";
            public const string Head_FolderPickingPerformed = "8";

            public const string Det_LineNum = "0";
            public const string Det_Select = "1";
            public const string Det_AvailableForRelease = "2";
            public const string Det_ToRelease = "3";
            public const string Det_QtyOpen = "4";
            public const string Det_ItemsPerUnit = "10000064";
            public const string Det_UomCode = "1470000128";
            public const string Det_WhsCode = "6";
            public const string Det_ItemCode = "8";
            public const string Det_DocEntry = "23";
            public const string Det_LineIdRdr1 = "31";
            public const string Det_RelatedOrd = "U_ITPS_RelatedOrder";
            public const string Det_SplitPercentage = "U_ITPS_SplitPercentage";
            public const string Det_Tipo = "U_Tipo";
            public const string Det_AvailableStock = "U_ITPS_AvailableStock";
        }

        public static class SalesOrder_SystemUIDs
        {
            public const string DocumentDateLbl = "86";
        }

        //public static class SalesOrder_CustomLabels
        //{
        //    public const string AssignedEntity = "Entidad empresa asignada";
        //    public const string RelatedOrder = "Pedido espejo relacionado";
        //    public const string SplitPercentage = "Porcentaje empresa contraria";
        //    public const string DocEntry = "Interno";
        //}

        //public static class SalesOrder_CustomLabelUIDs
        //{
        //    public const string Head_AssignedEntity = "itpsAeL";
        //    public const string Head_RelatedOrder = "itpsRoL";
        //    public const string Head_SplitPercentage = "itpsSpL";
        //    public const string Head_DocEntry = "itpsDocE";
        //}
       
        public static class ItemImport
        {
            public const string OitmImportProperty = "QryGroup1";
        }

        public static class BusinessPartner_Fields
        {
            public const string Head_SplitPercentage = "U_ITPS_SplitPercentage";
        }

        public static class Tables
        {
            public const string ConfigsDevelopment = "CONFIGS_DEVELOPMENT";
            public const string ConfigsDevelopmentWithAt = "@CONFIGS_DEVELOPMENT";
            public const string BusinessPartner = "OCRD";
            public const string UserFieldsMetadata = "CUFD";
            public const string AfipCondition = "VTL_CODCONDAFIP";
        }

        public static class ConfigProps
        {
            public const string TaxCodeSecondaryOrder = "pTaxCodeSecondaryOrder";
            public const string ImportIncomeAccountB = "pImportIncomeAccountB";
        }

        public static class DbViews
        {
            public const string StockSplitVta = "V_ITPS_EA_STOCK_SPLIT_VTA";
        }

        public static class FixedValues
        {
            public const string UomUnit = "Un";

            public const string EntityA = "NORMAL";
            public const string EntityB = "PRESUPUESTO";

            public const string Yes = "Y";
            public const string No = "N";

            public const string IvaExemptCode = "IVA_EXE";
        }

        public static class IvaCategory
        {
            public const string Exempt = "EX";
            public const string Monotax = "MT";
            public const string RegisteredTaxpayer = "RI";
            public const string FinalConsumer = "CF";
            public const string NotRegistered = "RNI";
            public const string NotTaxed = "NG";
            public const string NotReached = "NA";
            public const string Uncategorized = "NC";
        }

        public static class DocumentLetters
        {
            public const string A = "A";
            public const string B = "B";
            public const string E = "E";
        }

        public static class Messages
        {
            public const string ConnectedOk = "Addon_AutoDivSalesOrd conectado correctamente.";
            public const string FatalErrorPrefix = "Error fatal en Addon_AutoDivSalesOrd: ";

            public const string FinalizingAddon = "Finalizando Addon_AutoDivSalesOrd...";
            public const string RestartingAddon = "Reiniciando Addon_AutoDivSalesOrd...";

            public const string UiApiNotDefined = "UIAPI no definido";
            public const string DiApiNotDefined = "DIAPI no definido";

            public const string MissingSapConnectionString = "No se recibió el string de conexión de SAP Business One. Ejecute el AddOn desde SAP (Add-On Administration).";
            public const string UiApiConnectErrorPrefix = "No fue posible conectarse a SAP Business One UI API (SboGuiApi.Connect). ";
            public const string UiApiGetApplicationErrorPrefix = "No fue posible obtener la instancia de SAP Business One UI API (SboGuiApi.GetApplication). ";
            public const string UiApiApplicationNull = "SboGuiApi.GetApplication devolvió null. Verifique que SAP Business One esté abierto y que el AddOn se ejecute desde SAP.";

            public const string EventRouterInitErrorPrefix = "Error inicializando eventos: ";
            public const string ItemEventErrorPrefix = "ItemEvent: ";
            public const string RightClickEventErrorPrefix = "RightClickEvent: ";
            public const string FormDataEventErrorPrefix = "FormDataEvent: ";
            public const string MenuEventErrorPrefix = "MenuEvent: ";
            public const string AppEventErrorPrefix = "AppEvent: ";

            public const string ValidationSplitPercentageRange = "Debes ingresar un porcentaje de split válido.";
            public const string ValidationSplitPercentageChange = "El porcentaje de split no puede ser modificado.";
            public const string ValidationNonDivisibleSplit = "No se puede dividir la línea porque la cantidades no quedan enteras.";
            public const string ValidationNonStock = "No hay stock suficiente para el artículo.";
            public const string ValidationCannotUpdateSplitOrder = "No se puede actualizar una orden que ha sido dividida. Debe actualizar la orden principal.";
            public const string ValidationEntityCompanyUpdate = "No se puede actualizar una orden asignada a PRESUPUESTO que tenga Split aplicado.";
            public const string ValidationEntityToUpdate = "No se puede cambiar el campo TIPO de la orden.";
            public const string ValidationDiscountNotAllowedForNormal = "No se puede aplicar un porcentaje de split del 75% en órdenes de tipo NORMAL (Empresa A).";
            public const string ValidationImportedOrderMixedItems = "Una Orden de Artículos Importados solo puede contener artículos con la Propiedad 'Importado' activa. Verifique las líneas del pedido.";
            public const string ValidationImportedOrderNoSplit = "Orden de Artículos Importados detectada. El split se fijó automáticamente al 100%% para NORMAL.";

            public const string ConfigInfraErrorPrefix = "Error creando tabla de configuraciones de desarrollos: ";
            public const string OrderInfraErrorPrefix = "Error creando infraestructura de UDF en ORDR: ";

            public const string TxtFilePathErrorPrefix = "GetTxtFilePath Error  -> ";
        }

    }
}
