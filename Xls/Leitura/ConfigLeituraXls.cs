using System;
using System.Collections.Generic;
using Xls.Core;

namespace Xls.Leitura
{
    public class ConfigLeituraXls<T>
    {
        public Func<T> CriarDadosLinha { get; set; }
        public IEnumerable<ConfigColunaXls<T>> Colunas { get; set; }
    }
}
