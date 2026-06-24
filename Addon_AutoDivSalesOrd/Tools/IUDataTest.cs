using Addon_AutoDivSalesOrd.Common;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Tools
{
    public class IUDataTest
    {
        public static void ExportMatrixColumnsToTxt(SAPbouiCOM.Matrix oMtx, string filePath)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"Matrix Name: {Constants.SalesOrder_FieldsUIDs.Head_Matrix}");
            sb.AppendLine($"Total Columns: {oMtx.Columns.Count}");
            sb.AppendLine(new string('-', 50));

            for (int i = 0; i < oMtx.Columns.Count; i++)
            {
                var column = oMtx.Columns.Item(i);

                sb.AppendLine($"Column {column.TitleObject.Caption}: UID = {column.UniqueID}");
            }

            System.IO.File.WriteAllText(filePath, sb.ToString());
        }

        public static void ExportMatrixColumnsDetailToTxt(SAPbouiCOM.Matrix oMtx, string filePath)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"Matrix UID: {Constants.SalesOrder_FieldsUIDs.Head_Matrix}");
            sb.AppendLine($"Total Columns: {oMtx.Columns.Count}");
            sb.AppendLine($"Total Rows: {oMtx.RowCount}");
            sb.AppendLine(new string('=', 80));

            for (int c = 0; c < oMtx.Columns.Count; c++)
            {
                var column = oMtx.Columns.Item(c);
                string colType = column.Type.ToString();
                string colUid = column.UniqueID;
                string colCaption = column.TitleObject.Caption;

                sb.AppendLine($"[Col {c}] Caption: \"{colCaption}\" | UID: {colUid} | Type: {colType}");

                for (int r = 1; r <= oMtx.RowCount; r++)
                {
                    string cellValue = string.Empty;
                    try
                    {
                        var cell = oMtx.GetCellSpecific(colUid, r);

                        if (cell is EditText et)
                            cellValue = et.Value;
                        else if (cell is SAPbouiCOM.ComboBox cb)
                            cellValue = cb.Value;
                        else if (cell is CheckBox chk)
                            cellValue = chk.Checked ? "Y" : "N";
                        else if (cell is ButtonCombo bc)
                            cellValue = bc.Selected?.Value ?? string.Empty;
                        else
                            cellValue = "(tipo no soportado)";
                    }
                    catch
                    {
                        cellValue = "(error al leer)";
                    }

                    sb.AppendLine($"    Row {r}: {cellValue}");
                }

                sb.AppendLine(new string('-', 80));
            }

            System.IO.File.WriteAllText(filePath, sb.ToString());
        }
    }
}
