using System;
using System.Linq.Expressions;

namespace Xls.Core
{
    internal interface IResolvedorExprLeituraXls
    {
        EnumTipoDadoXls ResolverTipoDado<T>(Expression<Func<T, object>> propExpr);
        bool ResolverObrigatorio<T>(Expression<Func<T, object>> propExpr);
        Action<T, object> CriarAplicadorValor<T>(Expression<Func<T, object>> propExpr);
    }
}
