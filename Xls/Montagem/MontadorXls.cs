using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Xls.Core;

namespace Xls.Montagem
{
    public class MontadorXls
    {
        public ResultadoMontagemXls Montar<T>(ConfigMontagemXls<T> config)
        {
            using var package = new ExcelPackage();

            var ws = package.Workbook.Worksheets.Add(config.Nome);

            var resolvedorPropExpr = new ResolvedorExpr();
            var configColunas = ConfigColunaMontagemXls<T>.Resolver(resolvedorPropExpr, config.Colunas);

            var linha = 1;

            foreach (var coluna in configColunas)
            {
                var celula = ws.Cells[linha, coluna.Indice];

                celula.Value = coluna.Titulo;
                AplicarFormato(coluna.Formato, celula);

                if (coluna.Obrigatorio)
                    celula.Style.Font.Bold = true;
            }

            foreach (var dados in config.Dados ?? Array.Empty<T>())
            {
                linha += 1;

                foreach (var coluna in configColunas)
                {
                    var celula = ws.Cells[linha, coluna.Indice];
                    
                    celula.Value = coluna.LerValor(dados);
                    AplicarFormato(coluna.Formato, celula);
                }
            }

            ws.Cells.AutoFitColumns();

            return new ResultadoMontagemXls
            {
                Nome = config.Nome,
                Bytes = package.GetAsByteArray()
            };
        }

        private void AplicarFormato(EnumFormatoXls formato, ExcelRange celula)
        {
            switch (formato)
            {
                case EnumFormatoXls.Inteiro:
                    celula.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                case EnumFormatoXls.Decimal:
                    celula.Style.Numberformat.Format = "#,##0.00########";
                    celula.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    break;
                case EnumFormatoXls.Data:
                    celula.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    celula.Style.Numberformat.Format = "dd/mm/yyyy";
                    break;
                case EnumFormatoXls.DataHora:
                    celula.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    celula.Style.Numberformat.Format = "dd/mm/yyyy hh:mm:ss";
                    break;
                case EnumFormatoXls.Texto:
                default:
                    celula.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    break;

            }
        }
    }
}
