﻿using System;
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

        public static IReadOnlyList<ConfigColunaLeituraXls<T>> Resolver(IResolvedorExprLeituraXls resolvedor, IEnumerable<ConfigColunaXls<T>> configColunas)
        {
            var configs = new List<ConfigColunaLeituraXls<T>>();

            foreach (var coluna in configColunas)
            {
                var config = new ConfigColunaLeituraXls<T>
                {
                    Indice = coluna.Indice,
                    Titulo = coluna.Titulo,
                    TipoDado = coluna.TipoDado ?? EnumTipoDadoXls.Unknown,
                    Obrigatorio = coluna.Obrigatorio ?? false,
                    AplicarValor = coluna.AplicarValor
                };

                if (coluna.PropExpr != null)
                {
                    if (!coluna.TipoDado.HasValue)
                        config.TipoDado = resolvedor.ResolverTipoDado(coluna.PropExpr);

                    if (!coluna.Obrigatorio.HasValue)
                        config.Obrigatorio = resolvedor.ResolverObrigatorio(coluna.PropExpr);

                    if (config.AplicarValor == null)
                        config.AplicarValor = resolvedor.CriarAplicadorValor(coluna.PropExpr);
                }

                Validar(config);
                configs.Add(config);
            }

            return configs;
        }

        [Conditional("DEBUG")]
        private static void Validar(ConfigColunaLeituraXls<T> config)
        {
            if (config.AplicarValor == null)
                throw new Exception($"Não foi possível resolver '{nameof(config.AplicarValor)}' para coluna {config.Titulo}");
        }
    }
}
