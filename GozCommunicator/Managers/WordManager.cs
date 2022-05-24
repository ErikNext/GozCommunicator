using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using GozCommunicator.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace GozCommunicator.Managers
{
    internal class WordManager
    {
        private string PathFile { get; }

        public WordManager(string pathFile)
        {
            if (File.Exists(pathFile))
            {
                Console.WriteLine("Файл Word найден");
                PathFile = pathFile;
            }
            else
            {
                Console.WriteLine("Файл Word не был найден");
            }
        }

        public Table GetTableFromWord()
        {
            if (PathFile != null)
            {
                try
                {
                    using (WordprocessingDocument wDoc = WordprocessingDocument.Open(PathFile, false))
                    {
                        var parts = wDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
                        if (parts != null)
                        {
                            foreach (var node in parts.ChildElements)
                            {
                                if (node is Table)
                                {
                                    return (Table)node;
                                }
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    Program.AbortApp(ex.Message);
                }
            }
            return new Table();
        }

        public Contract GetContractAnIgkOrNull(List<Contract> contracts, Contract searchedContract)
        {
            foreach(var contract in contracts)
            {
                if(contract.Igk == searchedContract.Igk)
                {
                    return contract;
                }
            }
            return null;
        }

        public List<Contract> TableParserToContracts(Table node)
        {
            Contract contract;
            var contracts = new List<Contract>();

            foreach (var row in node.Descendants<TableRow>())
            {
                if (Regex.IsMatch(row.Descendants<TableCell>().ElementAt(0).InnerText, @"\d"))
                {
                    contract = new Contract()
                    {
                        Id = row.Descendants<TableCell>().ElementAt(0).InnerText,
                        Customer = row.Descendants<TableCell>().ElementAt(1).InnerText,
                        Theme = row.Descendants<TableCell>().ElementAt(2).InnerText,
                        NumberGosContract = row.Descendants<TableCell>().ElementAt(3).InnerText,
                        Igk = row.Descendants<TableCell>().ElementAt(4).InnerText,
                        CustomersСurrentAccountNumber = row.Descendants<TableCell>().ElementAt(5).InnerText,
                    };
                    foreach (var lines in row.Descendants<TableCell>().ElementAt(6).Descendants<Paragraph>())
                    {
                        contract.AccountNumberAvionika += lines.InnerText + "\n";
                    }
                    foreach (var lines in row.Descendants<TableCell>().ElementAt(7).Descendants<Paragraph>())
                    {
                        contract.Remark += lines.InnerText + "\n";
                    }

                    var foundContract = GetContractAnIgkOrNull(contracts, contract);

                    if (foundContract == null)
                        contracts.Add(contract);
                    else
                    {
                        foundContract.Theme += "\n" + contract.Theme;
                        foundContract.CustomersСurrentAccountNumber += "\n" + contract.CustomersСurrentAccountNumber;
                        foundContract.AddingAccountNumberAvionika("\n" + contract.AccountNumberAvionika);
                        foundContract.Remark += "\n" + contract.Remark;
                    }
                }
                else if (row.Descendants<TableCell>().ElementAt(0).InnerText == "")
                {
                    contracts.Last().NumberGosContract += "\n" + row.Descendants<TableCell>().ElementAt(3).InnerText;
                    contracts.Last().Remark += "\n" + row.Descendants<TableCell>().ElementAt(7).InnerText;
                }
            }
            return contracts;
        }
    }
}
