using System;
using System.Collections.Generic;
using System.Diagnostics;
using Xls.Core;

namespace Xls.Montagem
{
    internal class ConfigColunaMontagemXls<T>
    {
        public int Indice { get; internal set; }
        public string Titulo { get; internal set; }
        public EnumFormatoXls Formato { get; internal set; }
        public EnumTipoDadoXls TipoDado { get; internal set; }
        public bool Obrigatorio { get; internal set; }
        public Func<T, object> LerValor { get; internal set; }

        public static IReadOnlyList<ConfigColunaMontagemXls<T>> Resolver(IResolvedorExprMontagemXls resolvedor, IEnumerable<ConfigColunaXls<T>> configColunas)
        {
            var configs = new List<ConfigColunaMontagemXls<T>>();

            foreach (var coluna in configColunas)
            {
                var config = new ConfigColunaMontagemXls<T>
                {
                    Indice = coluna.Indice,
                    Titulo = coluna.Titulo,
                    TipoDado = coluna.TipoDado ?? EnumTipoDadoXls.Unknown,
                    Formato = coluna.Formato ?? EnumFormatoXls.Texto,
                    Obrigatorio = coluna.Obrigatorio ?? false,
                    LerValor = coluna.LerValor
                };

                if (coluna.PropExpr != null)
                {
                    if (!coluna.TipoDado.HasValue)
                        config.TipoDado = resolvedor.ResolverTipoDado(coluna.PropExpr);

                    if (!coluna.Obrigatorio.HasValue)
                        config.Obrigatorio = resolvedor.ResolverObrigatorio(coluna.PropExpr);

                    if (!coluna.Formato.HasValue)
                        config.Formato = resolvedor.ResolverFormato(coluna.PropExpr);

                    if (config.LerValor == null)
                        config.LerValor = resolvedor.CriarLeitorValor(coluna.PropExpr);
                }

                ValidarMapeador(config);
                configs.Add(config);
            }

            return configs;
        }

        [Conditional("DEBUG")]
        private static void ValidarMapeador(ConfigColunaMontagemXls<T> config)
        {
            if (config.LerValor == null)
                throw new Exception($"Não foi possível resolver '{nameof(config.LerValor)}' para coluna {config.Titulo}");
        }
    }
}