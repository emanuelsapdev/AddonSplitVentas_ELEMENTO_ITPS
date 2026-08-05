using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Services
{

    public class ValidationService
    {

        public static void ValidateImportedOrderItems(Dictionary<string, bool> importedFlags)
        {
            bool anyImported = importedFlags.Values.Any(v => v);
            if (!anyImported) return;

            bool allImported = importedFlags.Values.All(v => v);
            if (!allImported)
                throw new Exception(Constants.Messages.ValidationImportedOrderMixedItems);
        }

        public static void ValidateEntityCompanyAtUpdate(SalesOrderFormModel data)
        {
            if (data.AssignedEntity == Constants.FixedValues.EntityB && data.SplitPercentage != 100m)
            {
                throw new Exception(Constants.Messages.ValidationEntityCompanyUpdate);
            }
        }

        public static void ValidatePorcentageAtCreate(SalesOrderFormModel data)
        {
            if (data.SplitPercentage < 0 || data.SplitPercentage > 100)
            {
                throw new Exception(Constants.Messages.ValidationSplitPercentageRange);
            }

            if (data.AssignedEntity == Constants.FixedValues.EntityA)
            {
                if (data.SplitPercentage == 75m)
                {
                    throw new Exception(Constants.Messages.ValidationDiscountNotAllowedForNormal);
                }
            }
        }

        //public static void ValidateEntityTypeAtUpdate(SalesOrderFormModel data, string prevEntity)
        //{
        //    if (data.AssignedEntity != prevEntity) throw new Exception(Constants.Messages.ValidationEntityToUpdate);
        //}


        public static void ValidatePorcentageAtUpdate(SalesOrderFormModel data)
        {
            Recordset rs = null;
            try
            {
                rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

                string q = $@"SELECT ""U_ITPS_SplitPercentage"" FROM ORDR WHERE ""DocEntry"" = '{data.DocEntry}'";

                rs.DoQuery(q);

                double perc = rs.Fields.Item(0).Value;

                if(perc != (double)data.SplitPercentage)
                {
                    throw new Exception(Constants.Messages.ValidationSplitPercentageChange);
                }
            } 
            finally
            {
                if (rs != null)
                    Marshal.ReleaseComObject(rs);
            }
}

        public static void ValidateQuantitiesAtUpdate(SalesOrderFormModel data, Dictionary<int, double> prevQtiesData)
        {
            decimal perc = data.SplitPercentage;

            foreach (var line in data.Lines)
            {
                string itemCode = line.ItemCode;
                decimal quantity = line.Quantity;
                //string uomCode = line.UomCode;

                bool isGetter = prevQtiesData.TryGetValue(line.LineId, out double qtyPrev);

                if (qtyPrev == (double)quantity && isGetter) continue;

                //if (uomCode == Constants.FixedValues.UomUnit)
                //{
                    bool hasDecimal = (quantity * perc / 100) % 1 != 0;

                    if (hasDecimal)
                    {
                        throw new Exception(Constants.Messages.ValidationNonDivisibleSplit);
                    }
                //}
            }
        }



        public static void ValidateQuantitiesAtCreate(SalesOrderFormModel data)
        {
            decimal perc = data.SplitPercentage;

            foreach (var line in data.Lines)
            {
                string itemCode = line.ItemCode;
                decimal quantity = line.Quantity;
                //string uomCode = line.UomCode;

                //if (uomCode == Constants.FixedValues.UomUnit)
                //{
                    bool hasDecimal = (quantity * perc / 100) % 1 != 0;

                    if (hasDecimal)
                    {
                        throw new Exception(Constants.Messages.ValidationNonDivisibleSplit);
                    }
                //}
            }

        }


        


        //public static void ValidateStockAtCreate(SalesOrderFormModel data)
        //{
        //    Recordset rs = null;
        //    try
        //    {
        //        rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

        //        foreach (var line in data.Lines)
        //        {
        //            if (string.IsNullOrWhiteSpace(line.ItemCode) || line.Quantity <= 0)
        //                continue;

        //            string safeItemCode = line.ItemCode.Replace("'", "''");
        //            string safeWhsCode = line.WhsCode.Replace("'", "''");
        //            rs.DoQuery($@"SELECT (COALESCE(""OnHand"", 0) - COALESCE(""IsCommited"", 0)) AS ""OnAvailable"" FROM OITW 
        //                            WHERE ""ItemCode"" = '{safeItemCode}' AND ""WhsCode"" = '{safeWhsCode}'");

        //            decimal onAvailable = 0m;
        //            if (!rs.EoF)
        //            {
        //                var value = rs.Fields.Item(0).Value;
        //                if (value != null)
        //                    onAvailable = Convert.ToDecimal(value);
        //            }

        //            decimal qty = line.Quantity;

        //            if (onAvailable < qty)
        //            {
        //                throw new Exception(Constants.Messages.ValidationNonStock + $" Art.:{line.ItemCode}. Disponible: {onAvailable}, Solicitado: {qty}.");
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        if (rs != null)
        //            Marshal.ReleaseComObject(rs);
        //    }
        //}


        //public static void ValidateStockAtUpdate(SalesOrderFormModel data)
        //{
        //    Recordset rs = null;
        //    try
        //    {
        //        rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);

        //        foreach (var line in data.Lines)
        //        {
        //            if (string.IsNullOrWhiteSpace(line.ItemCode))
        //                continue;

        //            string safeItemCode = line.ItemCode.Replace("'", "''");
        //            string safeWhsCode = line.WhsCode.Replace("'", "''");
        //            rs.DoQuery($@"SELECT (COALESCE(T0.""OnHand"", 0) - COALESCE(T0.""IsCommited"", 0)) + T1.""Quantity""
        //                        FROM OITW T0
        //                        INNER JOIN RDR1 T1
        //                        ON T1.""ItemCode"" = T0.""ItemCode"" AND T1.""WhsCode"" = T0.""WhsCode"" 
        //                        WHERE T0.""ItemCode"" = '{safeItemCode}' AND T0.""WhsCode"" = '{safeWhsCode}'
        //                        AND T1.""DocEntry"" = {data.DocEntry}");

        //            decimal onAvailable = 0m;
        //            if (!rs.EoF)
        //            {
        //                var value = rs.Fields.Item(0).Value;
        //                if (value != null)
        //                {
        //                    onAvailable = Convert.ToDecimal(value);
        //                    break;
        //                }
        //            }

        //            decimal qty = line.Quantity;

        //            if (onAvailable < qty)
        //            {
        //                throw new Exception(Constants.Messages.ValidationNonStock + $" Art.:{line.ItemCode}. Disponible: {onAvailable}, Solicitado: {qty}.");
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        if (rs != null)
        //            Marshal.ReleaseComObject(rs);
        //    }
        //}
    }
}
