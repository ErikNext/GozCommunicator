using GozCommunicator.Core;
using GozCommunicator.Managers;
using System;

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

            excelManager.ContractToExcel(contracts);

            Console.WriteLine("\nИзменения успешно применены!\n");

            foreach(var line in Statistic.CreatedLines)
            {
                Console.WriteLine($"На строке {line.Row} была создана новая запись!");
            }

            Console.ReadLine();
        }

        public static void AbortApp(string message)
        {
            Console.WriteLine(message);
            Console.ReadLine();
            Environment.Exit(0);
        }
    }
}
