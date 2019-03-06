using Microsoft.Vbe.Interop;
using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    class Program
    {
        //static string fileLocation = "C:\\Users\\chace\\Desktop\\CompoundInterst.xls";
        static string fileLocation = System.IO.Directory.GetCurrentDirectory() + "\\CompoundInterest.xls";
        static double start_amount = 0;
        static double intial = 0;
        static double interest = 0;
        static double years = 0;
        static double monthlyDeposit = 0;

        public static double YearlyInterst()
        {
            double monthlyRate = (interest / 100) / 12;
            double totalInterst = 0;
            double amount = intial;

            for (int i = 1; i <= 12; i++)
            {
                totalInterst = totalInterst + (amount * monthlyRate);
                amount = amount + (amount * monthlyRate);
                amount = amount + monthlyDeposit;
                
            }

            return Math.Round(totalInterst, 2);
        }

        public static void GenerateFile()
        {
          Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
          Excel.Workbook xlWorkBook;
          Excel.Worksheet xlWorkSheet;
          object misValue = System.Reflection.Missing.Value;
          xlWorkBook = xlApp.Workbooks.Add(misValue);
          xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

          // calculates heading
          //[row, col]
          xlWorkSheet.Cells[1, 1] = "intial amount";    xlWorkSheet.Cells[1, 2] = String.Format("{0:0,0.00}", start_amount);
          xlWorkSheet.Cells[2, 1] = "interest";  xlWorkSheet.Cells[2, 2] = String.Format("{0:0,0.00}", interest);
          xlWorkSheet.Cells[3, 1] = "years invested";  xlWorkSheet.Cells[3, 2] = years.ToString();
          xlWorkSheet.Cells[4, 1] = "yearly deposit"; xlWorkSheet.Cells[4, 2] = String.Format("{0:0,0.00}", monthlyDeposit);
          xlWorkSheet.Cells[7, 1] = "year";
          xlWorkSheet.Cells[7, 2] = "start amount";
          xlWorkSheet.Cells[7, 3] = "yearly deposit";
          xlWorkSheet.Cells[7, 4] = "interest earned";
          xlWorkSheet.Cells[7, 5] = "total";

          // calculates main section of compound interest 
          int line = 8;
          for (int i = 0; i < years; i++)
          {
            double total = intial + YearlyInterst() + (monthlyDeposit * 12);
            xlWorkSheet.Cells[line, 1] = String.Format("{0:0,0.00}", (i + 1));
            xlWorkSheet.Cells[line, 2] = String.Format("{0:0,0.00}", intial);
            xlWorkSheet.Cells[line, 3] = String.Format("{0:0,0.00}", (monthlyDeposit * 12));
            xlWorkSheet.Cells[line, 4] = String.Format("{0:0,0.00}", YearlyInterst());
            xlWorkSheet.Cells[line, 5] = String.Format("{0:0,0.00}", total);
            line = line + 1;
            intial = total;
          }

          double total_Invested = ((monthlyDeposit * 12) * years) + start_amount;
          
          // calculates the footer
          line = line + 2;
          xlWorkSheet.Cells[line, 1] = "total amount";
          xlWorkSheet.Cells[line, 2] = String.Format("{0:0,0.00}", intial);
          line = line + 1;
          xlWorkSheet.Cells[line, 1] = "amount invested";
          xlWorkSheet.Cells[line, 2] = String.Format("{0:0,0.00}", total_Invested);
          line = line + 1;
          xlWorkSheet.Cells[line, 1] = "accumalated amount";
          xlWorkSheet.Cells[line, 2] = String.Format("{0:0,0.00}", (intial - total_Invested));

          xlWorkSheet.Columns.AutoFit();
          xlWorkBook.SaveAs(fileLocation, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
          xlWorkBook.Close(true, misValue, misValue);
          xlApp.Quit();

          Marshal.ReleaseComObject(xlWorkSheet);
          Marshal.ReleaseComObject(xlWorkBook);
          Marshal.ReleaseComObject(xlApp);
        }

        public static void WriteFile()
        {
            using (StreamWriter writetext = new StreamWriter(@"C:\Users\chace\Desktop\write.csv"))
            {
                writetext.WriteLine("intial amount," + start_amount.ToString());
                writetext.WriteLine("interest," + interest.ToString());
                writetext.WriteLine("years invested," + years.ToString());
                writetext.WriteLine("yearly deposit," + monthlyDeposit.ToString());
                writetext.WriteLine();
                writetext.WriteLine();
                writetext.WriteLine("year, start amount, interest earned, total");

               
                
                for (int i = 0; i < years; i++)
                {
                    double total = intial + YearlyInterst() + (monthlyDeposit * 12);
                    writetext.WriteLine((i+1).ToString() + "," + intial.ToString() + "," +  (monthlyDeposit*12).ToString()  + "," + YearlyInterst().ToString() + "," + total);
                    intial = total;
                }

                double total_Invested = ((monthlyDeposit * 12) * years) + start_amount;

                writetext.WriteLine();
                writetext.WriteLine("total amount," + intial);
                writetext.WriteLine("amount invested," + total_Invested);
                writetext.WriteLine("accumalated amount," + (intial - total_Invested));

                writetext.Close();
                Console.WriteLine("file created");
            }
        }

        public static void GetInput()
        {
            Console.WriteLine("type in intial amount");
            start_amount = double.Parse(Console.ReadLine());
            intial = start_amount;
            Console.WriteLine("type in interest amount");
            interest = double.Parse(Console.ReadLine());
            Console.WriteLine("type in number of years");
            years = double.Parse(Console.ReadLine());
            Console.WriteLine("type in monthly amount invested");
            monthlyDeposit = double.Parse(Console.ReadLine());
        }

        public static void Main(string[] args)
        {
          GetInput();
          GenerateFile();
        }
    }
}
