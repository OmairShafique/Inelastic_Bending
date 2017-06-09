using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;




namespace Inelastic_Bending
{
    class Program
    {

        
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);





            Console.WriteLine("Hello and Welcome," + "\n" + "Made by Omair Shafiq");
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            Console.WriteLine("Please enter the initial Value for Strain In outer most fibre:"); 
            e = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the initial Value for Distance to outer most fibre:");
            y = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the initial Value for Distance to yield strain:");
            c = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the Elastic Modulus:");
            E = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the applied Moment:");
            M = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the total length for the beam:");
            h = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the web thickness of the T-Section:");
            t_w = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the web height of the T-Section:");
            h_w = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the flange width for the T-Section:");
            w_f = Convert.ToDouble(Console.ReadLine());
            Console.WriteLine("Please enter the flange thickness for the T-Section:");
            t_f = Convert.ToDouble(Console.ReadLine());



            Processing(e, y, c, E, h, M, out T, out C, out s_y, out tol, out TOL1);
            
            while (abs(tol) > 0.01 | abs(TOL1) > 0.01)
            {
                Processing(e, y, c, E, h, M, out T, out C, out s_y, out tol, out TOL1);
                 
                if (tol < 0)
                {
                    My_Case = Case.LessThanZero;                  
                }
                else if (tol < 0 & tol > -1)
                {
                    My_Case = Case.LessThanZeroAndMoreThan1;
                }
                else if (tol > 0)
                {
                    My_Case = Case.MoreThanZero;               

                }
                else if (tol > 0 & tol < 1)
                {
                    My_Case = Case.MoreThanZeroAndLessThan1;
                }

                switch (My_Case)
                {
                    case Case.MoreThanZero:
                        y = y - y / 100;
                        c = c - c / 100;
                        break;
                    case Case.MoreThanZeroAndLessThan1:
                        y = y - y / 100000;
                        c = c - c / 100000;
                        break;
                    case Case.LessThanZero:
                        y = y + y / 100;
                        c = c + c / 100;
                        break;
                    case Case.LessThanZeroAndMoreThan1:
                        y = y + y / 100000;
                        c = c + c / 100000;
                        break;                   
                    default:
                        break;
                }


                if (TOL1 < 0)
                {
                    e = e - e / 100;

                }

                else if (TOL1 > 0)
                {
                    e = e + e / 100;
                }

                Iteration++;


                xlWorkSheet.Cells[1, 1] = "Iteration";
                xlWorkSheet.Cells[1, 2] = "e";
                xlWorkSheet.Cells[1, 3] = "y";
                xlWorkSheet.Cells[1, 4] = "c";
                xlWorkSheet.Cells[1, 5] = "Error in T+C ";
                xlWorkSheet.Cells[1, 6] = "Error in M = Ty+Cy'";


                xlWorkSheet.Cells[Iteration+1, 1] = Iteration ;
                xlWorkSheet.Cells[Iteration+1, 2] = e;
                xlWorkSheet.Cells[Iteration+1, 3] = y;
                xlWorkSheet.Cells[Iteration+1, 4] = c;
                xlWorkSheet.Cells[Iteration+1, 5] = tol;
                xlWorkSheet.Cells[Iteration + 1, 6] = TOL1;
            


                Result(T, C, s_y, tol, TOL1, e, y);

                if (Iteration == 100)
                {
                    Console.WriteLine("Do you want to continue:[Y/N]");
                    string a = Console.ReadLine();
                    if (a=="N")
                    {
                        break;
                    }
                    else
                    {
                        continue;
                    }
                    
                }
                if (Iteration == 500)
                {
                    Console.WriteLine("Do you want to continue:[Y/N]");
                    string a = Console.ReadLine();
                    if (a == "N")
                    {
                        break;
                    }
                    else
                    {
                        continue;
                    }

                }
                if (Iteration == 1000)
                {
                    Console.WriteLine("Do you want to continue:[Y/N]");
                    string a = Console.ReadLine();
                    if (a == "N")
                    {
                        break;
                    }
                    else
                    {
                        continue;
                    }

                }

            }

            Excel.Range formatRange;

            xlWorkSheet.Activate();
            xlWorkSheet.Application.ActiveWindow.SplitRow = 1;
            xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
           
            xlWorkSheet.Activate();
            xlWorkSheet.Application.ActiveWindow.SplitColumn = 1;
            xlWorkSheet.Application.ActiveWindow.FreezePanes = true;
           
            formatRange = xlWorkSheet.get_Range("a1");
            formatRange.EntireRow.Font.Bold = true;

            formatRange = xlWorkSheet.get_Range("a1", "f1");
            formatRange.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            formatRange.Font.Size = 14;
            formatRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);


            formatRange = xlWorkSheet.get_Range("a2");
            formatRange.EntireColumn.NumberFormat = "#,###,###";
            formatRange.EntireColumn.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            formatRange.EntireColumn.ColumnWidth = 13.33;
            formatRange.Font.Bold = true;

            formatRange = xlWorkSheet.get_Range("b2");
            formatRange.EntireColumn.NumberFormat = "0.00E+00";
            formatRange.EntireColumn.ColumnWidth = 15.33;

            formatRange = xlWorkSheet.get_Range("c2");
            formatRange.EntireColumn.NumberFormat = "#,##0.00000";
            formatRange.EntireColumn.ColumnWidth = 15.33;

            formatRange = xlWorkSheet.get_Range("d2");
            formatRange.EntireColumn.NumberFormat = "#,##0.00000";
            formatRange.EntireColumn.ColumnWidth = 15.33;

            formatRange = xlWorkSheet.get_Range("e2");
            formatRange.EntireColumn.NumberFormat = "0.00%";
            formatRange.EntireColumn.ColumnWidth = 25.33;

            formatRange = xlWorkSheet.get_Range("f2");
            formatRange.EntireColumn.NumberFormat = "0.00%";
            formatRange.EntireColumn.ColumnWidth = 25.33;

            
            string fileName = String.Format(@"{0}\Inelastic_Bending.xls", System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase));

            Console.WriteLine("\nResult:" + "\n" + "Strain = " + Convert.ToString(e) + "\nLocation of NA from bottom of beam = " + Convert.ToString(y) + "\nLocation of Yield Strian from NA = " + Convert.ToString(c));
            xlWorkBook.SaveAs(fileName, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file in " + fileName);

            Console.ReadLine();

        }

        



        private static Case My_Case = default(Case);
        private static Double T, ey, C, s_y, tol, TOL1, e, M, h, y, c, E, Iteration, t_w, t_f, h_w, w_f;
        

        private static Double abs(Double a)
        {
            if (a < 0)
            {
                a = a * -1;
            }
            else if (a > 0)
            {
                a = a * 1;
            }
            return a;
        }

        private static void Processing(Double e, Double y, Double c, Double E, Double h, Double M, out Double T, out Double C, out Double s_y, out Double tol, out Double TOL1)
        {
            
            Calc(e, y, c, E, h, out T, out C, out s_y, out My_Case);

            tol = ((T - C) / T);
            Double a1, b1;
            TOL1 = 0;
            if ((h-y)<c)
            {
                My_Case = Case.Less;
               
            }
            
            if (My_Case == Case.More)
            {
                a1 = C*(2*(h-y)/3);
                b1 = T * (((2 * c / 3) + ((y + c) * (y - c)) / 2)) / ((2 * y + c) / 2);
                TOL1 = (M - (a1 + b1)) / M;
            }
            else if (My_Case == Case.Less)
            {
                a1 = C * (((h-y+c)*(h-y-c)/2)+(c*c)/3)/((2*(h-y-c)+c)/2);
                b1 = T * (((2 * c / 3) + ((y + c) * (y - c)) / 2)) / ((2 * y + c) / 2);
                TOL1 = (M - (a1 + b1)) / M;
            }
            
        }

        private static void Result(Double T, Double C, Double s_y, Double tol, Double TOL, Double e, Double y)
        {
            Console.WriteLine("\n\nIteration = "+ Iteration +"\n\nT = " + Convert.ToString(T) + "\nC = " + Convert.ToString(C) + "\nYield Stress = " + Convert.ToString(s_y) + "\nError in T+C = " + Convert.ToString(tol) + "\nError in M=Ty+Cy' = " + Convert.ToString(TOL) + "\nResult:" + "\n" + "Strain = " + Convert.ToString(e) + "\nLocation of NA from bottom of beam = " + Convert.ToString(y) + "\nLocation of Yield Strian from NA = " + Convert.ToString(c));
        }

        private static void Calc(Double e, Double y, Double c, Double E, Double h, out Double T, out Double C, out Double s_y, out Case My_Case)
        {
            ey = (e * c) / y;
            s_y = E * ey;
            T =((s_y / 2) * c * t_w + s_y * (y-c) * t_w);
            C = 0;
            My_Case = Case.Fail;
            if ((h - y) > c)
            {
                if (y > h_w | y == h_w)
                {
                    C = s_y * w_f * (h - y - c) + (s_y / 2) * (t_f - h - y - c) * w_f;
                }
                else if (y < h_w)
                {
                        if ((y + c) > h_w)
                        {
                        Double sl = (s_y * (h_w - y)) / c;
                        C = s_y * w_f * (h - y - c) + ((s_y - sl) / 2) * (t_f - h - y - c) * w_f + (sl / 2) * t_w * (h_w - y) + (sl) * (t_f - h - y - c) * w_f;
                        }
                        else if ((y + c) < h_w)
                        {
                            C = s_y * w_f * t_f + s_y * (h_w - y - c) * t_w + (s_y / 2) * c * t_w;
                        }
                }
            }
            else if (( h - y ) < c | ( h - y ) == c)
            {
                if (y > h_w | y == h_w)
                {
                    C = (s_y / 2) * (t_f - h - y) * w_f;
                }
                else if (y < h_w)
                {
                    Double sl = (s_y * (h - y - t_w)) / t_w;
                    C = ((s_y - sl) / 2) * w_f * t_f + (sl) * w_f * t_f + (sl / 2) * t_w * (h_w - y);
                }

            }
        }
        enum Case
        {
            MoreThanZero,
            MoreThanZeroAndLessThan1,
            LessThanZero,
            LessThanZeroAndMoreThan1,
            Pass,
            Fail,
            Less,
            More
        }
    }
    
}
