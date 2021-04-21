using System;
using System.Linq.Expressions;
using System.Reflection;

namespace Xls.Core
{
    internal class ResolvedorExpr
    {
        internal EnumTipoDadoXls ResolverTipoDado<T>(Expression<Func<T, object>> propExpr)
        {
            var prop = ResolverProp(propExpr.Body);
            var type = prop.PropertyType;

            if (type == typeof(string))
                return EnumTipoDadoXls.Texto;

            if (type == typeof(int) || type == typeof(int?))
                return EnumTipoDadoXls.Inteiro;

            if (type == typeof(decimal) || type == typeof(decimal?))
                return EnumTipoDadoXls.Decimal;

            if (type == typeof(bool) || type == typeof(bool?))
                return EnumTipoDadoXls.Bool;

            if (type == typeof(DateTime) || type == typeof(DateTime?))
                return EnumTipoDadoXls.Data;

            return EnumTipoDadoXls.Unknown;
        }

        internal bool ResolverObrigatorio<T>(Expression<Func<T, object>> propExpr)
        {
            var prop = ResolverProp(propExpr.Body);
            return !IsNullable(prop.PropertyType);
        }

        internal Action<T, object> CriarAplicadorValor<T>(Expression<Func<T, object>> propExpr)
        {
            var prop = ResolverProp(propExpr.Body);
            var expression = ResolverMemberExpression(propExpr.Body);
             
            var objectParam = Expression.Parameter(typeof(T));
            var valueParam = Expression.Parameter(typeof(object));
            var left = AplicarLambdaParam(objectParam, expression.Expression, expression.Member);
            var right = Expression.Convert(valueParam, prop.PropertyType);
            var body = Expression.Assign(left, right);
            var lambda = Expression.Lambda<Action<T, object>>(body, new[] { objectParam, valueParam });

            return lambda.Compile();
        }

        private Expression AplicarLambdaParam(ParameterExpression param, Expression expression, MemberInfo member)
        {
            if (expression is MemberExpression me)
            {
                var expr = AplicarLambdaParam(param, me.Expression, me.Member);
                return Expression.MakeMemberAccess(expr, member);
            }
            return Expression.MakeMemberAccess(param, member);
        }

        private MemberExpression ResolverMemberExpression(Expression body)
        {
            switch (body)
            {
                case UnaryExpression ue:
                    return ResolverMemberExpression(ue.Operand);
                case MemberExpression me:
                    return me;
                default:
                    throw new Exception("Tipo de expressão não reconhecido.");
            }
        }

        private static PropertyInfo ResolverProp(Expression expr)
        {
            switch (expr)
            {
                case UnaryExpression ue:
                    return ResolverProp(ue.Operand);
                case MemberExpression me:
                    return ((PropertyInfo)me.Member);
                default:
                    throw new Exception("Tipo de expressão não reconhecido.");
            }
        }

        private static bool IsNullable(Type type)
        {
            return type.IsGenericType &&
                   type.GetGenericTypeDefinition() == typeof(Nullable<>) &&
                   !type.IsGenericTypeDefinition;
        }
    }
}
