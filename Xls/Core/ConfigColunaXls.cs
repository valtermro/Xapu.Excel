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
        public EnumFormatoXls? Formato { get; set; }
        public bool? Obrigatorio { get; set; }
        public Func<T, object> LerValor { get; set; }
        public Action<T, object> AplicarValor { get; set; }
    }
}
