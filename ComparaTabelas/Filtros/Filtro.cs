using ClosedXML.Excel;

namespace ComparaTabelas
{
    class Filtro
    {
        private int linha;
        private string colunaFiltro;
        private bool achou = false;
        private string caminho1;
        //public string caminho2;
        private string plan;

        private void filtro()
        {
            this.linha = 1;
            this.colunaFiltro = "";
            this.achou = false;
        }
        public void setCaminho(string caminho)
        {
            caminho1 = caminho;
        }

        public void setPlanilha(string planilha)
        {
            plan = planilha;
        }

        public void comparacomFitro(string colunaA)
        {

            XLWorkbook wb = new XLWorkbook(caminho1);
            IXLWorksheet planilha = wb.Worksheet(plan);
            filtro();

            while (true)
            {
                colunaFiltro = planilha.Cell(linha, 1).Value.ToString();
                if (colunaA == colunaFiltro)
                {
                    this.achou = true;
                    break;
                }
                if (string.IsNullOrEmpty(colunaFiltro)) { this.achou = false; break; }
                linha++;
            }
            // return achou;
        }

        public bool getAchou()
        {
            return achou;
        }
    }
}
