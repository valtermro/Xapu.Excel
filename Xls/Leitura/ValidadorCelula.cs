using System.Collections.Generic;
using Xls.Core;

namespace Xls.Leitura
{
    internal class ValidadorCelula
    {
        public bool ValidarValor(ICollection<MensagemLinhaLeituraXls> mensagens,
                                   string rotulo,
                                   EnumTipoDadoXls tipoDado,
                                   bool obrigatorio,
                                   object valor)
        {
            if (valor == null)
            {
                if (!obrigatorio) return true;
                mensagens.Add(new MensagemLinhaLeituraXls($"Valor para o campo '{rotulo}' não informado."));
                return false;
            }

            switch (tipoDado)
            {
                case EnumTipoDadoXls.Bool:
                    return ValidarBool(mensagens, rotulo, valor);
                case EnumTipoDadoXls.Texto:
                    return ValidarTexto(mensagens, rotulo, valor);
                case EnumTipoDadoXls.Inteiro:
                    return ValidarInteiro(mensagens, rotulo, valor);
                case EnumTipoDadoXls.Decimal:
                    return ValidarDecimal(mensagens, rotulo, valor);
                case EnumTipoDadoXls.Data:
                    return ValidarData(mensagens, rotulo, valor);
                case EnumTipoDadoXls.DataHora:
                    return ValidarDataHora(mensagens, rotulo, valor);
                default:
                    return true;
            }
        }

        private bool ValidarBool(ICollection<MensagemLinhaLeituraXls> mensagens, string rotulo, object valor)
        {
            var str = ((valor as string) ?? string.Empty).ToUpper();

            if (str != "FALSE" && str != "TRUE")
            {
                mensagens.Add(new MensagemLinhaLeituraXls($"Valor para '{rotulo}' precisa ser 'TRUE' ou 'FALSE'."));
                return false;
            }
            return true;
        }

        private bool ValidarTexto(ICollection<MensagemLinhaLeituraXls> mensagens, string rotulo, object valor)
        {
            // qualquer coisa é válida
            return true;
        }

        private bool ValidarInteiro(ICollection<MensagemLinhaLeituraXls> mensagens, string rotulo, object valor)
        {
            // TODO
            return true;
        }

        private bool ValidarDecimal(ICollection<MensagemLinhaLeituraXls> mensagens, string rotulo, object valor)
        {
            // TODO
            return true;
        }

        private bool ValidarData(ICollection<MensagemLinhaLeituraXls> mensagens, string rotulo, object valor)
        {
            // TODO
            return true;
        }

        private bool ValidarDataHora(ICollection<MensagemLinhaLeituraXls> mensagens, string rotulo, object valor)
        {
            // TODO
            return true;
        }
    }
}