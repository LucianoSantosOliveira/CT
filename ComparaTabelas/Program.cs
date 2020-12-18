using ClosedXML.Excel;
using ComparaTabelas.Filtros;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using System;


namespace ComparaTabelas
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.SetWindowSize(28, 40);

          //

            VerificarDuplicidadeProject Verificar = new VerificarDuplicidadeProject();

            int colunaBComp = 0;
            int colunaAComp = 0;
            bool achouIgual = false;
            string colunaA;
            string colunaB;

            do
            {
                Filtro Ft = new Filtro();
                FiltroCliente Fc = new FiltroCliente();
                FiltroEncerrados Fe = new FiltroEncerrados();
                
                
                    
                    Ft.setCaminho(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminhoFiltro.txt").ToString());
                    Ft.setPlanilha("Plani1");

                    
                    Fc.setCaminho(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminhoFiltroCliente.txt").ToString());
                    Fc.setPlanilha("Sheet1");

                    
                    Fe.setCaminho(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminhoFiltroEncerrado.txt").ToString());
                    Fe.setPlanilha("Sheet1");
                
               
                
                    
                
                XLWorkbook wb = new XLWorkbook(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminho.txt").ToString());
                var planilha = wb.Worksheet("Planilha1");
                // Ft.setCaminho(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminhoFiltro.txt").ToString());
                // Ft.setCaminho(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminhoFiltro.txt").ToString());
                // Fc.setCaminho(System.IO.File.ReadAllText(@"C:\ComparaTabelas\caminhoFiltroCliente.txt").ToString());

                Console.Clear();
                Console.WriteLine("-------------------------");
                colunaBComp = 1;
                colunaAComp = 1;
                achouIgual = false;
                //primeiro while é para seguir proxima celula da coluna 1
                while (true)
                {
                    colunaBComp = 1;
                    colunaA = planilha.Cell(colunaAComp, 1).Value.ToString();
                    //while para comparar celula da coluna 1 com a 2
                    while (true)
                    {
                        colunaB = planilha.Cell(colunaBComp, 2).Value.ToString();

                        if (colunaA == colunaB)
                        {
                            achouIgual = true;
                            break;
                        }
                        if (string.IsNullOrEmpty(colunaB)) { achouIgual = false; break; }
                        colunaBComp++;
                    }

                   
                    if (achouIgual == false)
                    {
                        Ft.comparacomFitro(colunaA);
                        if (Ft.getAchou() == true)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine(colunaA + " Chamado vinculado");
                        }
                        else if (Ft.getAchou() == false)
                        {                       
                            Fc.comparacomFitro(colunaA);
                            if (Fc.getAchou() == true) { Console.ForegroundColor = ConsoleColor.DarkRed; Console.WriteLine(colunaA + " Executor MOVIDESK"); } else { Console.ResetColor(); Console.WriteLine(colunaA); }
                            Console.ResetColor();
                            
                        }
                    }
                    if (string.IsNullOrEmpty(colunaA)) { break; }
                    colunaAComp++;
                }                
                Console.ResetColor();
                colunaBComp = 1;
                Console.WriteLine("-------------------------");

                while (true)
                {
                    colunaAComp = 1;
                    colunaB = planilha.Cell(colunaBComp, 2).Value.ToString();
                    if (colunaBComp > 1) { Verificar.setCellAnterior(planilha.Cell(colunaBComp - 1, 2).Value.ToString()); }
                    Verificar.setCellproxima(planilha.Cell(colunaBComp, 2).Value.ToString());
                    Verificar.chamarComparacao();
                    while (true)
                    {
                        colunaA = planilha.Cell(colunaAComp, 1).Value.ToString();
                        if (colunaB == colunaA)
                        {
                            achouIgual = true;
                           //Fc.comparacomFitro(colunaB);
                            break;
                        }
                        if (string.IsNullOrEmpty(colunaA)) { achouIgual = false; break; }
                        colunaAComp++;
                    }

                    if (achouIgual == false || Verificar.getAchouduplicidade() == true)  { 

                        //Verificar.chamarComparacao();
                        if(Verificar.getAchouduplicidade() == true) { Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine(colunaB); } else { Console.ResetColor(); }
                       // Console.WriteLine(colunaB);
                        Fe.comparacomFitro(colunaB);
                        //Fc.comparacomFitro(colunaB);
                        if (Fe.getAchou() == true)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkYellow;
                            Console.WriteLine(colunaB + " Encerrado");
                        }
                        else if(achouIgual == false && Verificar.getAchouduplicidade() == false )
                        {
                            Console.ResetColor();
                            Console.WriteLine(colunaB);
                        }
                       // Console.WriteLine(colunaB);
                    }
                    Console.ResetColor();
                    if (string.IsNullOrEmpty(colunaB)) { break; }
                    colunaBComp++;
                }
                Console.WriteLine("-------------------------");
            } while (Console.ReadKey().Key == ConsoleKey.Enter);
        }
    }
}
