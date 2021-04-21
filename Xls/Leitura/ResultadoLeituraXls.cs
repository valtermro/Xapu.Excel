using System.Collections.Generic;

namespace Xls.Leitura
{
    public class ResultadoLeituraXls<T>
    {
        public IReadOnlyList<LinhaLeituraXls<T>> Linhas { get; internal set; }
    }

    public class LinhaLeituraXls<T>
    {
        public int Numero { get; internal set; }
        public T Dados { get; internal set; }
        public IReadOnlyList<MensagemLinhaLeituraXls> Mensagens { get; internal set; }
    }

    public class MensagemLinhaLeituraXls
    {
        public MensagemLinhaLeituraXls(string mensagem) => Mensagem = mensagem;
        public string Mensagem { get; internal set; }
    }
}
