using Addon_AutoDivSalesOrd.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Addon_AutoDivSalesOrd.Forms.GerentePicking
{
    public partial class GerentePickingFrm
    {
        
        private void CalculateToRelease(Dictionary<string, List<List<GerentePickingFormRow>>> filasAgrupadas)
        {
            foreach (var entradaArticulo in filasAgrupadas)
            {
                var primeraFila = entradaArticulo.Value.SelectMany(g => g).First();

                //if (primeraFila.ItemCode != "TN10000258") continue; // quitar

                // Ajustar el stock disponible a bultos completos
                decimal stockDisponible = primeraFila.AvailableStock; // 46 * 12 = 552
                decimal cantidadDocenasPorBulto = GetQtyDozenPerPackage(primeraFila.ItemCode); // 20
                // AvailableStock ya está en la misma escala que QtyOpen * ItemsPerUnit,
                // por lo que el tamaño de bulto se expresa directamente en esa escala (sin *12).
                decimal unidadesPorBulto = cantidadDocenasPorBulto > 0 ? cantidadDocenasPorBulto * 12m : 0m;

                decimal stockRestante = unidadesPorBulto > 0
                    ? Math.Floor(stockDisponible / unidadesPorBulto) * unidadesPorBulto
                    : stockDisponible;

                foreach (var grupoSplit in entradaArticulo.Value)
                {
                    if (stockRestante <= 0m)
                    {
                        foreach (var fila in grupoSplit)
                            fila.ToRelease = 0m;
                        continue;
                    }

                    decimal totalNecesarioEnBase = grupoSplit.Sum(r => r.QtyOpen * r.ItemsPerUnit);

                    // Monto candidato a liberar para el grupo, ajustado a un múltiplo entero de bultos
                    decimal montoALiberarEnBase = Math.Min(totalNecesarioEnBase, stockRestante);
                    montoALiberarEnBase = unidadesPorBulto > 0
                        ? Math.Floor(montoALiberarEnBase / unidadesPorBulto) * unidadesPorBulto
                        : montoALiberarEnBase;

                    if (montoALiberarEnBase <= 0m)
                    {
                        // No alcanza para completar un bulto: no se libera nada y el stock
                        // queda disponible para que lo complete un grupo posterior.
                        foreach (var fila in grupoSplit)
                            fila.ToRelease = 0m;
                        continue;
                    }

                    if (montoALiberarEnBase >= totalNecesarioEnBase)
                    {
                        // Alcanza para cubrir toda la demanda del grupo
                        foreach (var fila in grupoSplit)
                            fila.ToRelease = Math.Min(Math.Floor(fila.QtyOpen), fila.AvailableForRelease);
                    }
                    else
                    {
                        // Stock insuficiente para cubrir toda la demanda, pero alcanza para liberar una parte
                        decimal proporcionLiberacion = totalNecesarioEnBase > 0m
                            ? montoALiberarEnBase / totalNecesarioEnBase
                            : 0m;

                        foreach (var fila in grupoSplit)
                        {
                            decimal necesidadFilaEnBase = fila.QtyOpen * fila.ItemsPerUnit;
                            decimal aLiberarEnBase = Math.Floor(necesidadFilaEnBase * proporcionLiberacion);
                            decimal aLiberarCalculado = fila.ItemsPerUnit > 0
                                ? Math.Floor(aLiberarEnBase / fila.ItemsPerUnit)
                                : 0m;

                            fila.ToRelease = Math.Min(aLiberarCalculado, fila.AvailableForRelease);
                        }
                    }

                    stockRestante -= montoALiberarEnBase;
                    if (stockRestante < 0m) stockRestante = 0m;
                }
            }
        }
    }
}