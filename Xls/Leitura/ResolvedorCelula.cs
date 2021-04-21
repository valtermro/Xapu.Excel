using System;
using Xls.Core;

namespace Xls.Leitura
{
    internal class ResolvedorCelula
    {
        internal object ResolverValor(EnumTipoDadoXls tipoDado, bool obrigatorio, object valor)
        {
            if (valor == null && !obrigatorio)
                return null;

            switch (tipoDado)
            {
                case EnumTipoDadoXls.Bool:
                    return ResolverBool(valor);
                case EnumTipoDadoXls.Texto:
                    return ResolverTexto(valor);
                case EnumTipoDadoXls.Inteiro:
                    return ResolverInteiro(valor);
                case EnumTipoDadoXls.Decimal:
                    return ResolverDecimal(valor);
                case EnumTipoDadoXls.Data:
                    return ResolverData(valor);
                case EnumTipoDadoXls.DataHora:
                    return ResolverDataHora(valor);
                default:
                    return valor;
            }
        }

        private object ResolverBool(object valor)
        {
            var str = (valor as string)?.ToLower();
            return str == "true";
        }

        private object ResolverTexto(object valor)
        {
            return Convert.ToString(valor);
        }

        private object ResolverInteiro(object valor)
        {
            return Convert.ToInt32(valor);
        }

        private object ResolverDecimal(object valor)
        {
            return Convert.ToDecimal(valor);
        }

        private object ResolverData(object valor)
        {
            return Convert.ToDateTime(valor).Date;
        }

        private object ResolverDataHora(object valor)
        {
            return Convert.ToDateTime(valor);
        }
    }
}