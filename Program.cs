using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1

{
    class Program
    {
        static void Main(string[] args)
        {

            string[,] testdata;

            //This is your first test calling

            string path = @"C:\Users\Mahesh Kolukula\Desktop\Files\testdata1.xlsx";
            string sheetname = "sheet1";

            string[,] exceldata = readDataFromExcel(path, sheetname);

            Console.WriteLine("============================================");
            Console.WriteLine("             Sheet 1 data             ");
            Console.WriteLine("============================================");

            int arrayrowCount1 = exceldata.GetLength(0);
            int arraycolCount1 = exceldata.GetLength(1);
            for (int i = 0; i < arrayrowCount1; i++)
            {
                for (int j = 0; j < arraycolCount1; j++)
                {
                    Console.Write("{0} \t", exceldata[i, j]);
                }
                Console.WriteLine();
            }


            //this is your second test
            string[,] exceldata1 = readDataFromExcel(@"C:\Users\Mahesh Kolukula\Desktop\Files\testdata1.xlsx", "sheet2");

            Console.WriteLine("============================================");
            Console.WriteLine("                sheet 2 Data        ");
            Console.WriteLine("============================================");

            int arrayrowCount2 = exceldata1.GetLength(0);
            int arraycolCount2 = exceldata1.GetLength(1);

            for (int i = 0; i < arrayrowCount2; i++)
            {
                for (int j = 0; j < arraycolCount2; j++)
                {
                    Console.Write("{0} \t", exceldata1[i, j]);
                }
                Console.WriteLine();
            }



            ////this is your Third test


            string[,] ExcelData3 = readDataFromExcel(@"C:\Users\Mahesh Kolukula\Desktop\Files\testdata1.xlsx", "sheet3");

            // Sending table data to other method-- this is another approch.
            SendingArrayData(ExcelData3);

            void SendingArrayData(string[,] data)

            {

                Console.WriteLine("============================================");
                Console.WriteLine("                sheet 3 Data        ");
                Console.WriteLine("============================================");

                int arrayrowCount3 = data.GetLength(0);
                int arraycolCount3 = data.GetLength(1);

                for (int i = 0; i < arrayrowCount3; i++)

                {
                    for (int j = 0; j < arraycolCount3; j++)
                    {
                        Console.Write("{0} \t", data[i, j]);
                    }
                    Console.WriteLine();
                }

            }

            Console.ReadKey();

            //this is excel method to Read the data

            string[,] readDataFromExcel(string excelPath, string sheet)
            {

                Application xlApp = new Application();
                Workbook xlWorkbook = xlApp.Workbooks.Open(excelPath);
                Worksheet xlWorksheet = xlWorkbook.Sheets[sheet];
                Range xlRange = xlWorksheet.UsedRange;


                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                testdata = new string[rowCount, colCount];

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        testdata[i - 1, j - 1] = xlRange.Cells[i, j].Value2.ToString();
                    }

                }



                //release com objects to fully kill excel process from running in the background

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();

                Marshal.ReleaseComObject(xlApp);

                return testdata;

            }

        }

    }



}

