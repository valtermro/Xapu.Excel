using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Xls.Core;

namespace Xls.Leitura
{
    public class LeitorXls
    {
        public ResultadoLeituraXls<T> Ler<T>(StreamReader reader, ConfigLeituraXls<T> config)
        {
            using var package = new ExcelPackage(reader.BaseStream);

            var ws = package.Workbook.Worksheets[0];
            var linhas = new List<LinhaLeituraXls<T>>();

            var validadorCelula = new ValidadorCelula();
            var resolvedorCelula = new ResolvedorCelula();
            var resolvedorPropExpr = new ResolvedorExpr();
            var configColunas = ConfigColunaLeituraXls<T>.Resolver(resolvedorPropExpr, config.Colunas);
            
            for (var idxLinha = ws.Dimension.Start.Row + 1; idxLinha <= ws.Dimension.End.Row; idxLinha++)
            {
                var dadosLinha = config.CriarDadosLinha();
                var mensagens = new List<MensagemLinhaLeituraXls>();

                foreach (var coluna in configColunas)
                {
                    var celula = ws.Cells[idxLinha, coluna.Indice];
                    
                    var valido = validadorCelula.ValidarValor(mensagens, coluna.Titulo, coluna.TipoDado, coluna.Obrigatorio, celula.Value);
                    if (!valido) continue;

                    var valor = resolvedorCelula.ResolverValor(coluna.TipoDado, coluna.Obrigatorio, celula.Value);
                    if (valor == null) continue;
                    
                    coluna.AplicarValor(dadosLinha, valor);
                }

                linhas.Add(new LinhaLeituraXls<T>
                {
                    Numero = idxLinha,
                    Dados = dadosLinha,
                    Mensagens = mensagens
                });
            }

            return new ResultadoLeituraXls<T>
            {
                Linhas = linhas
            };
        }
    }
}
