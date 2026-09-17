using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Services;
using SAPbobsCOM;
using System;
using System.Collections.Generic;

namespace Addon_AutoDivSalesOrd.Configuration
{
    /// <summary>
    /// Infraestructura de inicialización: crea tablas UDT, campos UDF
    /// y siembra datos de configuración en @CONFIGS_DEVELOPMENT.
    /// </summary>
    public class ConfigsDevelopmentInfra
    {
        private static Company _company;

        public ConfigsDevelopmentInfra()
        {
            _company = ConnectionSDK.DIAPI;
        }

        /// <summary>
        /// Ejecuta la creación de infraestructura completa (tabla + campos + datos semilla).
        /// </summary>
        public void ProcessInfrastructure()
        {
            CreateTableConfigsDevelopment();
            SeedConfigsDevelopment();
        }

        private void CreateTableConfigsDevelopment()
        {
            string table = "CONFIGS_DEVELOPMENT";
            try
            {
                // Crear tabla de configuraciones
                InfraDataService.CreateUserTable(table, "Configuraciones de desarrollos", BoUTBTableType.bott_NoObject);

                // Crear campos de la tabla
                InfraDataService.CreateUserField($"@{table}", "ITPS_Prop", "Propiedad", BoFieldTypes.db_Alpha, 30);
                InfraDataService.CreateUserField($"@{table}", "ITPS_Value", "Valor", BoFieldTypes.db_Memo);
                InfraDataService.CreateUserField($"@{table}", "ITPS_Integration", "Nombre Integración", BoFieldTypes.db_Alpha, 40);
                InfraDataService.CreateUserField($"@{table}", "ITPS_Author", "Desarrollador", BoFieldTypes.db_Alpha, 40);
                InfraDataService.CreateUserField($"@{table}", "ITPS_UpdateDate", "Fecha modificación", BoFieldTypes.db_Date);
                InfraDataService.CreateUserField($"@{table}", "ITPS_DescriptionProp", "Descripción de propiedad", BoFieldTypes.db_Memo);
            }
            catch (Exception ex)
            {
                throw new Exception("Error creando tabla de configuraciones de desarrollos: " + ex.Message, ex);
            }
        }

        private static void SeedConfigsDevelopment()
        {
            var oRec = (SAPbobsCOM.Recordset)_company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                InsertDevelopmentConfig(
                    ref oRec,
                    code: "Addon_AutoDivSalesOrd_01",
                    name: "Addon_AutoDivSalesOrd_01",
                    prop: "pTaxCodeB",
                    value: "IVA_EXE",
                    integration: "Auto. División Pedidos por Entidad",
                    author: "EA",
                    updateDate: DateTime.Today,
                    descriptionProp: "Código de Impuesto de todas las líneas."
                );

                
                InsertDevelopmentConfig(
                    ref oRec,
                    code: "Addon_AutoDivSalesOrd_02",
                    name: "Addon_AutoDivSalesOrd_02",
                    prop: "pImportIncomeAccountB",
                    value: "",
                    integration: "Auto. División Pedidos por Entidad",
                    author: "EA",
                    updateDate: DateTime.Today,
                    descriptionProp: "Cuenta de ingresos empresa B para asiento de importados."
                );

            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            }
        }

        private static void InsertDevelopmentConfig(
            ref SAPbobsCOM.Recordset oRec,
            string code,
            string name,
            string prop,
            string value,
            string integration,
            string author,
            DateTime updateDate,
            string descriptionProp)
        {
            // Verificar si el Code ya existe
            string sqlExists = $@"
                SELECT 1
                FROM ""@CONFIGS_DEVELOPMENT""
                WHERE ""Code"" = '{EscapeSql(code)}'";

            oRec.DoQuery(sqlExists);

            if (!oRec.EoF)
                return; // Ya existe, no insertar

            // Insertar solo si NO existe
            string sqlInsert = $@"
                INSERT INTO ""@CONFIGS_DEVELOPMENT""
                (
                    ""Code"", ""Name"",
                    ""U_ITPS_Prop"", ""U_ITPS_Value"",
                    ""U_ITPS_Integration"", ""U_ITPS_Author"",
                    ""U_ITPS_UpdateDate"", ""U_ITPS_DescriptionProp""
                )
                VALUES
                (
                    '{EscapeSql(code)}', '{EscapeSql(name)}',
                    '{EscapeSql(prop)}', '{EscapeSql(value)}',
                    '{EscapeSql(integration)}', '{EscapeSql(author)}',
                    '{updateDate:yyyyMMdd}', '{EscapeSql(descriptionProp)}'
                )";

            oRec.DoQuery(sqlInsert);
        }

        private static string EscapeSql(string s) => (s ?? string.Empty).Replace("'", "''");
    }

    /// <summary>
    /// Infraestructura para campos UDF en ORDR (Pedidos de venta).
    /// </summary>
    public class SalesOrderSplitInfra
    {
        public void ProcessInfrastructure()
        {
            try
            {
                // Sales Order UDFs
                InfraDataService.CreateUserField( 
                    tableName: "ORDR",
                    fieldName: "ITPS_SplitPercentage",
                    desc: "Porcentaje Split Aplicado",
                    type: BoFieldTypes.db_Float,
                    subType: BoFldSubTypes.st_Percentage,
                    validValues: new List<InfraDataService.ValidValueOption>
                    {
                        new InfraDataService.ValidValueOption { Value = "100", Description = "100%" },
                        new InfraDataService.ValidValueOption { Value = "50", Description = "50%" },
                        new InfraDataService.ValidValueOption { Value = "25", Description = "25%" },
                        new InfraDataService.ValidValueOption { Value = "0", Description = "0%" },
                        new InfraDataService.ValidValueOption { Value = "-1", Description = "-" },
                    });

                InfraDataService.CreateUserField(
                     tableName: "ORDR",
                     fieldName: "ITPS_ImportedPercentage",
                     desc: "Porcentaje de Importados",
                     type: BoFieldTypes.db_Float,
                    subType: BoFldSubTypes.st_Percentage,
                    validValues: new List<InfraDataService.ValidValueOption>
                    {
                        new InfraDataService.ValidValueOption { Value = "100", Description = "1 - 100%" },
                        new InfraDataService.ValidValueOption { Value = "50", Description = "2 - 50%" }
                    });

                InfraDataService.CreateUserField(
                    tableName: "ORDR",
                    fieldName: "ITPS_RelatedOrder",
                    desc: "Pedido Split",
                    type: BoFieldTypes.db_Numeric,
                    linkedSystemObject: UDFLinkedSystemObjectTypesEnum.ulOrders);

                InfraDataService.CreateUserField(
                    tableName: "ORDR",
                    fieldName: "CATEGORIA_CLIENTE",
                    desc: "Categoria Cliente",
                    type: BoFieldTypes.db_Alpha,
                    size: 10,
                    validValues: new List<InfraDataService.ValidValueOption>
                        {
                            new InfraDataService.ValidValueOption { Value = "AA", Description = "AA" },
                            new InfraDataService.ValidValueOption { Value = "A", Description = "A" },
                            new InfraDataService.ValidValueOption { Value = "B", Description = "B" },
                            new InfraDataService.ValidValueOption { Value = "C", Description = "C" },
                        });

                InfraDataService.CreateUserField(
                   tableName: "ORDR",
                   fieldName: "ITPS_DESCUENTO",
                   desc: "Descuento",
                   type: BoFieldTypes.db_Float,
                   subType: BoFldSubTypes.st_Percentage
                   );

                InfraDataService.CreateUserField(
                   tableName: "ORDR",
                   fieldName: "Tipo",
                   desc: "Tipo",
                   type: BoFieldTypes.db_Alpha,
                   size: 15,
                   defaultValue: "A DEFINIR",
                   validValues: new List<InfraDataService.ValidValueOption>
                        {
                            new InfraDataService.ValidValueOption { Value = "NORMAL", Description = "STS" },
                            new InfraDataService.ValidValueOption { Value = "PRESUPUESTO", Description = "RIAG" },
                            new InfraDataService.ValidValueOption { Value = "A DEFINIR", Description = "A DEFINIR" },
                        }
                   );

                InfraDataService.CreateUserTable(name: "ITPS_ZONAENTREGA", desc: "Zonas de Entrega", type: BoUTBTableType.bott_NoObject);

                InfraDataService.CreateUserField(
                    tableName: "ORDR",
                    fieldName: "ITPS_DeliveryZone",
                    desc: "Zona de Entrega",
                    type: BoFieldTypes.db_Alpha,
                    size: 50,
                    linkedTable: "ITPS_ZONAENTREGA");

                InfraDataService.CreateUserField(
                    tableName: "ORDR",
                    fieldName: "ITPS_NRO_AC",
                    desc: "Nro Acuerdo Comercial",
                    type: BoFieldTypes.db_Numeric,
                    size: 10);

                InfraDataService.CreateUserField(
                    tableName: "ORDR",
                    fieldName: "ITPS_AgreementPriceList",
                    desc: "Lista de precio del Acuerdo Global",
                    type: BoFieldTypes.db_Alpha,
                    size: 150);

                // Sales Order UDFs End

                // Direcciones de socios de negocios
                InfraDataService.CreateUserField(
                     tableName: "CRD1",
                     fieldName: "ITPS_DeliveryZone",
                     desc: "Zona de Entrega",
                     type: BoFieldTypes.db_Alpha,
                     size: 50,
                     linkedTable: "ITPS_ZONAENTREGA");

                InfraDataService.CreateUserField(
                     tableName: "OCRD",
                     fieldName: "ITPS_PRIORIDAD",
                     desc: "Prioridad",
                     type: BoFieldTypes.db_Float,
                        subType: BoFldSubTypes.st_Percentage,
                        validValues: new List<InfraDataService.ValidValueOption>
                        {
                            new InfraDataService.ValidValueOption { Value = "100", Description = "1 - 100%" },
                            new InfraDataService.ValidValueOption { Value = "50", Description = "2 - 50%" },
                            new InfraDataService.ValidValueOption { Value = "25", Description = "4 - 25%" },
                            new InfraDataService.ValidValueOption { Value = "0", Description = "3 - 0%" },
                        });

                InfraDataService.CreateUserField(
                     tableName: "OCRD",
                     fieldName: "ITPS_PRIORIDAD_IMPORTADOS",
                     desc: "Prioridad Importados",
                     type: BoFieldTypes.db_Float,
                        subType: BoFldSubTypes.st_Percentage,
                        validValues: new List<InfraDataService.ValidValueOption>
                        {
                            new InfraDataService.ValidValueOption { Value = "100", Description = "1 - 100%" },
                            new InfraDataService.ValidValueOption { Value = "50", Description = "2 - 50%" },
                        });

                InfraDataService.CreateUserField(
                    tableName: "OJDT",
                    fieldName: "ITPS_RelatedInvoice",
                    desc: "Factura Importados",
                    type: BoFieldTypes.db_Numeric,
                    linkedSystemObject: UDFLinkedSystemObjectTypesEnum.ulInvoices);

                InfraDataService.CreateViewIfNotExists(Constants.DbViews.StockSplitVta, @"WITH ""ENPICKLIST"" AS (SELECT
                                                                                            R1.""ItemCode"",
                                                                                            R1.""WhsCode"",
                                                                                            SUM(P1.""RelQtty"") AS ""CantEnPicking""
                                                                                        FROM PKL1 P1
                                                                                        INNER JOIN RDR1 R1
                                                                                            ON R1.""DocEntry"" = P1.""OrderEntry""
                                                                                           AND R1.""LineNum""  = P1.""OrderLine""
                                                                                        WHERE P1.""BaseObject"" = '17'         -- Orden de Venta; sumar más ramas con UNION ALL si hay otros BaseObject
                                                                                          AND P1.""PickStatus"" <> 'C'       -- ⚠️ no confirmé los valores reales de PickStatus (¿'C'=cerrado? ¿'Y'/'N'?) — revisar
                                                                                        GROUP BY R1.""ItemCode"", R1.""WhsCode"") SELECT
                                                                                        CASE
                                                                                            WHEN T1.""UgpEntry"" = -1
                                                                                                OR T2.""AltQty"" IS NULL
                                                                                                OR T2.""AltQty"" = 0
                                                                                                THEN T0.""OnHand"" - IFNULL(E.""CantEnPicking"", 0) -- + T0.""IsCommited"" - T0.""OnOrder"" - 
                                                                                            ELSE (T0.""OnHand""  - IFNULL(E.""CantEnPicking"", 0)) * (T2.""BaseQty"" / T2.""AltQty"")
                                                                                        END                                             AS ""AvailableStock_Unidades_Base"",
                                                                                        T0.""ItemCode"",
                                                                                        T0.""WhsCode"",
                                                                                        IFNULL(E.""CantEnPicking"", 0)                    AS ""YaEnListaPicking"",
                                                                                        T0.""ItemCode"" AS ""kEY_ItemCode"",
                                                                                        T0.""WhsCode""  AS ""kEY_WhsCode""
                                                                                    FROM
                                                                                        OITW T0
                                                                                    INNER JOIN OITM T1 ON T1.""ItemCode"" = T0.""ItemCode""
                                                                                    LEFT  JOIN UGP1 T2 ON T2.""UgpEntry"" = T1.""UgpEntry"" AND T2.""UomEntry"" = T1.""IUoMEntry""
                                                                                    LEFT  JOIN OUOM T3 ON T3.""UomEntry"" = T1.""IUoMEntry""
                                                                                    LEFT  JOIN EnPickList E ON E.""ItemCode"" = T0.""ItemCode"" AND E.""WhsCode"" = T0.""WhsCode""");
            }
            catch (Exception ex)
            {
                throw new Exception("Error creando infraestructura de UDF en ORDR: " + ex.Message, ex);
            }
        }
    }
}
