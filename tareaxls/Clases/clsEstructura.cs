using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace tareaxls.Clases
{
    class clsEstructura
    {
        public String Nombre { get; set; }
        public String direccion { get; set; }

        private String[,] matrizR;
        private String[,] matrizR2;
        private int f1, c1;

        private void TamanoMatriz(int fila, int Columna) {
            matrizR = new string[fila, Columna];
            matrizR2 = new string[fila, Columna];
        }
        private void GuardaDatosMatriz(int f, int c, string dato) {
            matrizR[f, c] = dato;
            
        }
        private void CambiaDatos() {
            int ftemp=-1, ctemp=-1;
            Console.WriteLine("Matriz Obtenida");
            Console.WriteLine();
            for (int i = 0; i < f1; i++) {
                for (int j = 0; j < c1; j++) {
                    Console.Write(matrizR[i, j] + " ");
                    if (matrizR[i, j] == "0") {
                        ftemp = i;
                        ctemp = j;
                    }
                }
                Console.WriteLine();
            }

            for (int a = 0; a < f1; a++) {
                for (int b = 0; b < c1; b++) {
                    if (a == ftemp || b == ctemp) {
                        matrizR2[a, b] = "0";
                    }
                    else
                    {
                        matrizR2[a, b] = "1";
                    }
                }
            }
            Console.WriteLine("");
        }
        public void MuestraNuevaMatriz() {
            Console.WriteLine("Nueva Matriz");
            for (int i = 0; i < f1; i++)
            {
                for (int j = 0; j < c1; j++)
                {
                    Console.Write(matrizR2[i, j] + " ");
                }
                Console.WriteLine();
            }
        }
        
        public void MuestraDatos() {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;
            int rCnt;
            int rCnt2;
            int rw = 0;
            int cl = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\villa\Desktop\5to Semestre\Programación III\Tareas\Tarea 1\tarea1.xlsx");
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            TamanoMatriz(rw, cl);
            f1 = rw; c1 = cl;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (rCnt2 = 1; rCnt2 <= cl; rCnt2++) {
                    int datoTemp =((int)(range.Cells[rCnt, rCnt2] as Excel.Range).Value2);
                    string datoTe = datoTemp + "";
                    GuardaDatosMatriz(rCnt - 1, rCnt2 - 1, datoTe);
                }
               // Console.WriteLine();

            }

            CambiaDatos();
            MuestraNuevaMatriz();
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }







        /* public List<clsEstructura> cargaDatosXLS()
         {
             Excel.Application xlApp;
             Excel.Workbook xlWorkBook;
             Excel.Worksheet xlWorkSheet;
             Excel.Range range;
             int rCnt;
             int rw = 0;
             int cl = 0;

             xlApp = new Excel.Application();
             xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\villa\Desktop\5to Semestre\Programación III\Tareas\Tarea 1\tarea.xlsx");
             xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

             range = xlWorkSheet.UsedRange;
             rw = range.Rows.Count;
             cl = range.Columns.Count;

             List<clsEstructura> todos = new List<clsEstructura>();
             clsEstructura individual = new clsEstructura();

             for (rCnt = 1; rCnt <= rw; rCnt++)
             {
                 individual.Nombre = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                 individual.direccion = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;

                 todos.Add(individual);
                 individual = new clsEstructura();

             }
             xlWorkBook.Close(true, null, null);
             xlApp.Quit();

             Marshal.ReleaseComObject(xlWorkSheet);
             Marshal.ReleaseComObject(xlWorkBook);
             Marshal.ReleaseComObject(xlApp);

             return todos;
         }*/
    }


}
