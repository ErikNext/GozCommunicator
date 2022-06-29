using GozCommunicator.Managers;
using System;
using System.Data;

namespace GozCommunicator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var wordManager = new WordManager(@"C:\Shikar\Договоры.docx");
            var excelManager = new ExcelManager(@"C:\Shikar\ГОЗ.xlsx");

            var tableFromWord = wordManager.GetTableFromWord();
            var contracts = wordManager.TableParserToContracts(tableFromWord);

            var table = excelManager.ReadExcelFile("Лист1");
            excelManager.CheckTable(table, contracts);

            Console.ReadLine();
            //excelManager.ContractToExcel(contracts);
        }

        public static void AbortApp(string message)
        {
            Console.WriteLine(message);
            Console.ReadLine();
            Environment.Exit(0);
        }
    }
}
