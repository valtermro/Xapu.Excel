namespace Xls.Montagem
{
    public class ResultadoMontagemXls
    {
        public string Nome { get; internal set; }
        public string Extensao => ".xlsx";
        public string FileName => Nome + Extensao;
        public byte[] Bytes { get; set; }
    }
}
