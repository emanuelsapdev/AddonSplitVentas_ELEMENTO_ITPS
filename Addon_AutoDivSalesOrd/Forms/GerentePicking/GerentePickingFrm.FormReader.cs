using Addon_AutoDivSalesOrd.Common;
using Addon_AutoDivSalesOrd.Models;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Addon_AutoDivSalesOrd.Forms.GerentePicking
{
    public partial class GerentePickingFrm
    {

        private GerentePickingForm ReadRows(
            Matrix oMtx,
            Func<string, string, decimal> getAvailableStock,
            Func<string, string, decimal> getItemsPerUnit)
        {
            var gerentePickingForm = new GerentePickingForm();

            for (int i = 1; i <= oMtx.RowCount; i++)
            {
                //var selected = ((CheckBox)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_Select, i));

                //if (!selected.Checked) continue;

                var relatedOrder = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_RelatedOrd, i)).Value;
                var docEntry = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_DocEntry, i)).Value;
                var lineNum = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_LineNum, i)).Value;
                var lineIdRdr1 = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_LineIdRdr1, i)).Value;
                var itemCode = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_ItemCode, i)).Value;
                var whsCode = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_WhsCode, i)).Value;
                var splitPercentage = ((ComboBox)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_SplitPercentage, i)).Value;
                var tipo = ((ComboBox)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_Tipo, i)).Value;
                var qtyOpen = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_QtyOpen, i)).Value;
                var uomCode = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_UomCode, i)).Value;
                var availableForRelease = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_AvailableForRelease, i)).Value;

                var row = new GerentePickingFormRow
                {
                    DocEntry = Convert.ToInt32(docEntry),
                    LineNum = Convert.ToInt32(lineNum),
                    LineIdRdr1 = Convert.ToInt32(lineIdRdr1),
                    ItemCode = itemCode,
                    WhsCode = whsCode,
                    RelatedOrd = !string.IsNullOrEmpty(relatedOrder) ? Convert.ToInt32(relatedOrder) : -1,
                    SplitPercentage = Convert.ToDecimal(splitPercentage),
                    Tipo = tipo,
                    QtyOpen = !string.IsNullOrEmpty(qtyOpen) ? Convert.ToDecimal(qtyOpen) : 0m,
                    ItemsPerUnit = getItemsPerUnit(itemCode, uomCode),
                    AvailableForRelease = !string.IsNullOrEmpty(availableForRelease) ? decimal.Parse(availableForRelease, CultureInfo.GetCultureInfo("es-AR")) : 0m,
                };

                decimal stock = getAvailableStock(row.ItemCode, row.WhsCode);
                row.AvailableStock = stock;

                gerentePickingForm.Rows.Add(row);
            }

            return gerentePickingForm;
        }

        private IEnumerable<int> GetMatrixDocEntries(Matrix oMtx)
        {
            var docEntries = new HashSet<int>();
            for (int i = 1; i <= oMtx.RowCount; i++)
            {
                var value = ((EditText)oMtx.GetCellSpecific(Constants.GerentePicking_FieldsUIDs.Det_DocEntry, i)).Value;
                if (!string.IsNullOrEmpty(value))
                    docEntries.Add(Convert.ToInt32(value));
            }
            return docEntries;
        }

        /// <summary>
        /// <summary>
        /// Agrupa las filas de <paramref name="dataMtx"/> primero por artículo (<see cref="GerentePickingFormRow.ItemCode"/>)
        /// y luego, dentro de cada artículo, agrupa en sublistas las filas que comparten split
        /// (donde el <see cref="GerentePickingFormRow.DocEntry"/> de una coincide con el
        /// <see cref="GerentePickingFormRow.RelatedOrd"/> de la otra).
        /// Las órdenes sin split quedan como sublistas de un solo elemento.
        /// </summary>
        /// <returns>
        /// Diccionario cuya clave es el <c>ItemCode</c> y cuyo valor es la lista de grupos de split
        /// de ese artículo. Cada grupo de split es una lista de 1 o 2 filas.
        /// </returns>
        private Dictionary<string, List<List<GerentePickingFormRow>>> GroupBySplit(GerentePickingForm dataMtx)
        {
            var result = new Dictionary<string, List<List<GerentePickingFormRow>>>();

            var rowsByItem = dataMtx.Rows
                .GroupBy(r => r.ItemCode)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var kvp in rowsByItem)
            {
                var itemRows = kvp.Value;
                var splitGroups = new List<List<GerentePickingFormRow>>();
                var assignedIndexes = new HashSet<int>();

                for (int i = 0; i < itemRows.Count; i++)
                {
                    if (assignedIndexes.Contains(i)) continue;

                    var row = itemRows[i];
                    var splitGroup = new List<GerentePickingFormRow> { row };
                    assignedIndexes.Add(i);

                    if (row.RelatedOrd != -1)
                    {
                        for (int j = i + 1; j < itemRows.Count; j++)
                        {
                            if (assignedIndexes.Contains(j)) continue;

                            var other = itemRows[j];
                            if (other.DocEntry == row.RelatedOrd || other.RelatedOrd == row.DocEntry)
                            {
                                splitGroup.Add(other);
                                assignedIndexes.Add(j);
                            }
                        }
                    }

                    splitGroups.Add(splitGroup);
                }

                result[kvp.Key] = splitGroups;
            }

            return result;
        }
    }
}
