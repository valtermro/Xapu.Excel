using System;
using System.Collections.Generic;
using System.Diagnostics;
using Xls.Core;

namespace Xls.Leitura
{
    internal class ConfigColunaLeituraXls<T>
    {
        public int Indice { get; internal set; }
        public string Titulo { get; internal set; }
        public EnumTipoDadoXls TipoDado { get; internal set; }
        public bool Obrigatorio { get; internal set; }
        public Action<T, object> AplicarValor { get; internal set; }

        internal static IReadOnlyList<ConfigColunaLeituraXls<T>> Resolver(ResolvedorExpr resolvedor, IEnumerable<ConfigColunaXls<T>> configs)
        {
            var infos = new List<ConfigColunaLeituraXls<T>>();

            foreach (var config in configs)
            {
                var info = new ConfigColunaLeituraXls<T>
                {
                    Indice = config.Indice,
                    Titulo = config.Titulo,
                    TipoDado = config.TipoDado ?? EnumTipoDadoXls.Unknown,
                    Obrigatorio = config.Obrigatorio ?? false,
                    AplicarValor = config.AplicarValor
                };

                if (config.PropExpr != null)
                {
                    if (!config.TipoDado.HasValue)
                        info.TipoDado = resolvedor.ResolverTipoDado(config.PropExpr);

                    if (!config.Obrigatorio.HasValue)
                        info.Obrigatorio = resolvedor.ResolverObrigatorio(config.PropExpr);

                    if (info.AplicarValor == null)
                        info.AplicarValor = resolvedor.CriarAplicadorValor(config.PropExpr);
                }

                Validar(info);
                infos.Add(info);
            }

            return infos;
        }

        [Conditional("DEBUG")]
        private static void Validar(ConfigColunaLeituraXls<T> info)
        {
            if (info.AplicarValor == null)
                throw new Exception($"Não foi possível resolver '{nameof(info.AplicarValor)}' para coluna {info.Titulo}");
        }
    }
}
