using System;
using System.Collections.Generic;
using System.Text;
using ClosedXML.Excel;
namespace ComparaTabelas
{
    class VerificarDuplicidadeProject 
    {

        private string cellAnterior,cellProxima;
        private bool achouDuplicidade;
        public void chamarComparacao()
        {
            ComparaAnteriorEProxima();
        }

        private void iniciaVariavel()
        {
            cellAnterior = "";
            cellProxima = "";
            achouDuplicidade = false;
        }

        public void setCellAnterior(string cellAnteriorProject)
        {
            iniciaVariavel();
            cellAnterior = cellAnteriorProject;
        }

        public void setCellproxima(string cellProximaComparar)
        {
            cellProxima = cellProximaComparar;
        }
        public bool getAchouduplicidade()
        {
            return this.achouDuplicidade; 
        }

        private void ComparaAnteriorEProxima()
        {
            if(cellAnterior==cellProxima)
            {
                //Console.ForegroundColor = ConsoleColor.Red;
                this.achouDuplicidade = true;
            }
            else if(cellAnterior != cellProxima)
            {
                this.achouDuplicidade = false;
                //Console.ResetColor();
            }
        }
    }
}
