using System.Collections.Generic;
using Xls.Core;

namespace Xls.Montagem
{
    public class ConfigMontagemXls<T>
    {
        public string Nome { get; set; }
        public IEnumerable<ConfigColunaXls<T>> Colunas { get; set; }
        public IEnumerable<T> Dados { get; set; }
    }
}