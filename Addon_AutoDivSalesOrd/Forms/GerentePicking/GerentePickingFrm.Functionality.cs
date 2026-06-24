using Addon_AutoDivSalesOrd.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Addon_AutoDivSalesOrd.Forms.GerentePicking
{
    public partial class GerentePickingFrm
    {
        /// <summary>
        /// Calcula y asigna en memoria la cantidad a liberar (<see cref="GerentePickingFormRow.ToRelease"/>)
        /// para cada fila del diccionario agrupado por artículo y split.
        /// <para>
        /// El cálculo considera la unidad de medida de cada línea mediante <c>ItemsPerUnit</c>:
        /// <list type="bullet">
        ///   <item><c>QtyOpen</c> se convierte a unidades base: <c>QtyOpen * ItemsPerUnit</c>.</item>
        ///   <item>El prorrateo se realiza en unidades base.</item>
        ///   <item><c>ToRelease</c> se devuelve en unidades de la orden: <c>ToRelease = ToReleaseEnBase / ItemsPerUnit</c>.</item>
        /// </list>
        /// </para>
        /// <para>
        /// Reglas de asignación por grupo de split dentro de un artículo:
        /// <list type="bullet">
        ///   <item>Líneas con <c>SplitPercentage == 100</c>: se asigna la cantidad completa pedida (<c>QtyOpen</c>),
        ///         limitada al stock restante del artículo.</item>
        ///   <item>Pares de split (<c>SplitPercentage != 100</c>): si el stock restante alcanza para ambas líneas
        ///         se asigna <c>QtyOpen</c> completo a cada una; en caso contrario el stock disponible
        ///         se prorratea según el <c>SplitPercentage</c> de cada fila.</item>
        /// </list>
        /// </para>
        /// El stock disponible del artículo se ajusta previamente a un múltiplo entero de bultos completos
        /// (docenas por bulto multiplicadas por 12), de modo que nunca se libera una cantidad que no cierre un bulto.
        /// El stock disponible del artículo se consume en el orden en que aparecen los grupos.
        /// </summary>
        /// <param name="groupedRows">Resultado de <see cref="GroupBySplit"/>.</param>
        private void CalculateToRelease(Dictionary<string, List<List<GerentePickingFormRow>>> groupedRows)
        {
            foreach (var itemEntry in groupedRows)
            {
                var firstRow = itemEntry.Value.SelectMany(g => g).First();

                // Ajustar el stock disponible a bultos completos
                decimal availableStock = firstRow.AvailableStock;
                decimal qtyDozenPerPack = GetQtyDozenPerPackage(firstRow.ItemCode);
                // Convertir docenas a unidades: docenas * 12
                decimal unitsPerPack = qtyDozenPerPack > 0 ? qtyDozenPerPack * 12m : 0m;

                decimal remainingStock = unitsPerPack > 0
                    ? Math.Floor(availableStock / unitsPerPack) * unitsPerPack
                    : availableStock;

                foreach (var splitGroup in itemEntry.Value)
                {
                    if (remainingStock <= 0m)
                    {
                        foreach (var row in splitGroup)
                            row.ToRelease = 0m;
                        continue;
                    }

                    bool isSplit = splitGroup.Any(r => r.SplitPercentage != 100m);

                    if (!isSplit)
                    {
                        // Línea individual o 100 %: asignar cantidad completa hasta el stock disponible
                        foreach (var row in splitGroup)
                        {
                            // Convertir QtyOpen a unidades base: QtyOpen * ItemsPerUnit
                            decimal qtyOpenInBase = row.QtyOpen * row.ItemsPerUnit;
                            decimal toReleaseInBase = Math.Floor(Math.Min(qtyOpenInBase, remainingStock));

                            // Convertir el resultado nuevamente a unidades de la orden
                            row.ToRelease = row.ItemsPerUnit > 0
                                ? Math.Floor(toReleaseInBase / row.ItemsPerUnit)
                                : 0m;

                            remainingStock -= toReleaseInBase;
                            if (remainingStock < 0m) remainingStock = 0m;
                        }
                    }
                    else
                    {
                        // Par de split con porcentajes distintos de 100 %
                        decimal totalNeededInBase = splitGroup.Sum(r => r.QtyOpen * r.ItemsPerUnit);

                        if (totalNeededInBase <= remainingStock)
                        {
                            // Stock suficiente: cada línea recibe su cantidad completa
                            foreach (var row in splitGroup)
                                row.ToRelease = Math.Floor(row.QtyOpen);

                            remainingStock -= totalNeededInBase;
                        }
                        else
                        {
                            // Stock insuficiente: prorratear según el porcentaje de split
                            foreach (var row in splitGroup)
                            {
                                decimal toReleaseInBase = Math.Floor(remainingStock * row.SplitPercentage / 100m);
                                row.ToRelease = row.ItemsPerUnit > 0
                                    ? Math.Floor(toReleaseInBase / row.ItemsPerUnit)
                                    : 0m;
                            }

                            remainingStock = 0m;
                        }
                    }
                }
            }
        }
    }
}