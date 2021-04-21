using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Xls.Core;
using Xls.Leitura;
using Xls.Montagem;

namespace XlsImpExp
{
    class DadoXls
    {
        public int? Numero { get; set; }
        public DateTime? Data { get; set; }
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

            //Criar();
            Ler();
        }

        private static void Criar()
        {
            var config = new ConfigMontagemXls<DadoXls>
            {
                Nome = "Teste",
                Colunas = ConfigColunas(),
                Dados = new[]
                {
                    new DadoXls
                    {
                        Numero = 42,
                        Data = DateTime.Today,
                        Ativo = true,
                        Valor = 99.5687m,
                        Obs = "Teste"
                    },
                    new DadoXls
                    {
                        Numero = null,
                        Data = null,
                        Ativo = null,
                        Valor = null,
                        Obs = null
                    },
                    new DadoXls
                    {
                        Numero = null,
                        Data = DateTime.Now,
                        Ativo = null,
                        Valor = null,
                        Obs = null
                    }
                }
            };

            var montador = new MontadorXls();
            var resultado = montador.Montar(config);

            File.WriteAllBytes($"D:\\Downloads\\{resultado.FileName}", resultado.Bytes);

            //var stream = new MemoryStream(resultado.Bytes);
            //File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "test.xlsx");
        }

        private static void Ler()
        {
            using var reader = new StreamReader("D:\\Downloads\\Teste.xlsx");

            var config = new ConfigLeituraXls<DadoXls>
            {
                CriarDadosLinha = () => new DadoXls(),
                Colunas = ConfigColunas()
            };

            var leitor = new LeitorXls();
            var resultado = leitor.Ler(reader, config);
        }

        private static IEnumerable<ConfigColunaXls<DadoXls>> ConfigColunas()
        {
            return new[]
            {
                new ConfigColunaXls<DadoXls>
                {
                    Indice = 1,
                    Titulo = "Número",
                    PropExpr = p => p.Numero
                },
                new ConfigColunaXls<DadoXls>
                {
                    Indice = 2,
                    Titulo = "Data",
                    PropExpr = p => p.Data,
                    TipoDado = EnumTipoDadoXls.DataHora,
                    Formato = EnumFormatoXls.DataHora
                },
                new ConfigColunaXls<DadoXls>
                {
                    Indice = 3,
                    Titulo = "Valor",
                    PropExpr = p => p.Valor,
                    LerValor = p => p.Valor.HasValue ? decimal.Truncate(p.Valor.Value) : (decimal?)null
                },
                new ConfigColunaXls<DadoXls>
                {
                    Indice = 4,
                    Titulo = "Obs",
                    Obrigatorio = false,
                    PropExpr = p => p.Obs
                },
                new ConfigColunaXls<DadoXls>
                {
                    Indice = 5,
                    Titulo = "Ativo",
                    PropExpr = p => p.Ativo
                }
            };
        }
    }
}
