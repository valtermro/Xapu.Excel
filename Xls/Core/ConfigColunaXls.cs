using System;
using System.Linq.Expressions;

namespace Xls.Core
{
    public class ConfigColunaXls<T>
    {
        public int Indice { get; set; }
        public string Titulo { get; set; }
        public Expression<Func<T, object>> PropExpr { get; set; }
        // 
        public EnumTipoDadoXls? TipoDado { get; set; }
        public bool? Obrigatorio { get; set; }
        public Action<T, object> AplicarValor { get; set; }

    }
}
