using System;
using System.Linq.Expressions;

namespace Xls.Core
{
    internal interface IResolvedorExprMontagemXls
    {
        EnumTipoDadoXls ResolverTipoDado<T>(Expression<Func<T, object>> propExpr);
        EnumFormatoXls ResolverFormato<T>(Expression<Func<T, object>> propExpr);
        bool ResolverObrigatorio<T>(Expression<Func<T, object>> propExpr);
        Func<T, object> CriarLeitorValor<T>(Expression<Func<T, object>> propExpr);
    }
}
