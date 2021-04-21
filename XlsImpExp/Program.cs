using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Xls.Core;
using Xls.Leitura;
using Xls.Montagem;

namespace XlsImpExp
{
    class LinhaLeituraXls
    {
        public int Numero { get; set; }
        public DateTime Data { get; set; }
        public decimal? Valor { get; set; }
        public string Obs { get; set; }
        public bool? Ativo { get; set; }
        //public Teste Teste { get; set; } = new Teste();
    }

    public class Teste
    {
        public DateTime Data { get; set; }
        public Teste2 Teste2 { get; set; } = new Teste2();
    }

    public class Teste2
    {
        public int Numero { get; set; }
        public Teste3 Teste3 { get; set; } = new Teste3();
    }

    public class Teste3
    {
        public bool Ativo { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Ler();
        }

        private static void Ler()
        {
            using var reader = new StreamReader("D:\\Labs\\Excel\\Importacao.xlsx");

            var config = new ConfigLeituraXls<LinhaLeituraXls>
            {
                CriarDadosLinha = () => new LinhaLeituraXls(),
                Colunas = ConfigColunas()
            };

            var leitor = new LeitorXls();
            var resultado = leitor.Ler(reader, config);
        }

        private static IEnumerable<ConfigColunaXls<LinhaLeituraXls>> ConfigColunas()
        {
            return new[]
            {
                new ConfigColunaXls<LinhaLeituraXls>
                {
                    Indice = 1,
                    Titulo = "Número",
                    PropExpr = p => p.Numero
                },
                new ConfigColunaXls<LinhaLeituraXls>
                {
                    Indice = 2,
                    Titulo = "Data",
                    PropExpr = p => p.Data
                },
                new ConfigColunaXls<LinhaLeituraXls>
                {
                    Indice = 3,
                    Titulo = "Valor",
                    PropExpr = p => p.Valor
                },
                new ConfigColunaXls<LinhaLeituraXls>
                {
                    Indice = 4,
                    Titulo = "Obs",
                    PropExpr = p => p.Obs
                },
                new ConfigColunaXls<LinhaLeituraXls>
                {
                    Indice = 5,
                    Titulo = "Ativo",
                    PropExpr = p => p.Ativo
                }
            };
        }
    }
}
