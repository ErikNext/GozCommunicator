using GozCommunicator.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace GozCommunicator.Managers
{
    internal class ExcelManager
    {
        public static Dictionary<int, string> ColumnsExcel = new Dictionary<int, string>();

        private string PathFile { get; }

        public ExcelManager()
        {
            ColumnsExcel.Add(1, "A");
            ColumnsExcel.Add(2, "B");
            ColumnsExcel.Add(3, "C");
            ColumnsExcel.Add(4, "D");
            ColumnsExcel.Add(5, "E");
            ColumnsExcel.Add(6, "F");
            ColumnsExcel.Add(7, "G");
            ColumnsExcel.Add(8, "H");
            ColumnsExcel.Add(9, "I");
            ColumnsExcel.Add(10, "J");
        }

        public ExcelManager(string pathFile)
        {
            if (File.Exists(pathFile))
            {
                Console.WriteLine("Файл Excel найден");
                PathFile = pathFile;
            }
            else
            {
                Console.WriteLine("Файл Excel не был найден");
            }
        }

        private static CellExcel GetCellSystemOrNull(Contract contract, Worksheet ObjWorkSheet)
        {

            for (int j = 6; j < int.MaxValue; j++)
            {
                Range range = ObjWorkSheet.get_Range($"{ColumnsExcel[2]}{j}");
                if (range.Text == string.Empty)
                    break;

                string igkFromContract = contract.Igk.Substring(contract.Igk.Length - 1);
                string igkFromExcel = Convert.ToString(range.Text);
                igkFromExcel = igkFromExcel.Substring(igkFromExcel.Length - 1);


                if (igkFromExcel == igkFromContract)
                {
                    return new CellExcel(ColumnsExcel[1], j);
                }
            }

            return null;
        }

        private void CreateSystem(Contract contract, Worksheet ObjWorkSheet)
        {
            for (int j = 6; j < int.MaxValue; j++)
            {
                Range range = ObjWorkSheet.get_Range($"{ColumnsExcel[1]}{j}");

                if (range.Text == string.Empty)
                {
                    Range lastNonEmptyCell = ObjWorkSheet.get_Range($"{ColumnsExcel[1]}{j - 1}");
                    var numberSystemString = lastNonEmptyCell.Text.Substring(lastNonEmptyCell.Text.Length - 1);
                    int.TryParse(numberSystemString, out int numberSystemInt);

                    ObjWorkSheet.Cells[j, ColumnsExcel[1]] = $"Система {numberSystemInt + 1}";
                    ObjWorkSheet.Cells[j, ColumnsExcel[2]] = contract.Igk;
                    ObjWorkSheet.Cells[j, ColumnsExcel[3]] = contract.Igk;
                    ObjWorkSheet.Cells[j, ColumnsExcel[4]] = contract.AccountNumberAvionika;
                    ObjWorkSheet.Cells[j, ColumnsExcel[4]].HorizontalAlignment = Constants.xlLeft;
                    ObjWorkSheet.Cells[j, ColumnsExcel[9]] = contract.Remark;

                    break;
                }
            }
        }

        private void UpdateSystem(Contract contract, Worksheet ObjWorkSheet, CellExcel cell)
        {
            ObjWorkSheet.Cells[cell.Row, ColumnsExcel[2]] = contract.Igk;
            ObjWorkSheet.Cells[cell.Row, ColumnsExcel[3]] = contract.Igk;
            ObjWorkSheet.Cells[cell.Row, ColumnsExcel[4]] = contract.AccountNumberAvionika;
            ObjWorkSheet.Cells[cell.Row, ColumnsExcel[4]].HorizontalAlignment = Constants.xlLeft;
            ObjWorkSheet.Cells[cell.Row, ColumnsExcel[9]] = contract.Remark;
        }

        public void ContractToExcel(List<Contract> contracts)
        {
            if (PathFile != null)
            {
                Application ObjExcel = new Application();
                Workbook ObjWorkBook = ObjExcel.Workbooks.Open(PathFile, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Worksheet ObjWorkSheet;
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                foreach (var contract in contracts)
                {
                    var cell = GetCellSystemOrNull(contract, ObjWorkSheet);
                    if (cell != null)
                    {
                        UpdateSystem(contract, ObjWorkSheet, cell);
                    }
                    else
                    {
                        CreateSystem(contract, ObjWorkSheet);
                    }
                }
                ObjWorkBook.Save();
                ObjExcel.Quit();
            }
        }

        public void CheckTable(System.Data.DataTable table, List<Contract> contracts)
        {
            foreach (DataRow dataRow in table.Rows)
            {
                foreach (var item in dataRow.ItemArray)
                {
                    Console.WriteLine(item);
                }
                break;
            }
        }

        public System.Data.DataTable ReadExcelFile(string sheetName)
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                string fileExtension = Path.GetExtension(PathFile);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + PathFile + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + PathFile + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
    }
}
