using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Data.OleDb;
using System.Data;
using System.Windows.Controls;
using CheZaProgaV2.Controls;
using ClosedXML.Excel;

namespace CheZaProgaV2
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var addresses = GetSourceAddresses();
            tbRecognizedAddressesInSource.Text = addresses.Count.ToString();
            FindMatches(addresses);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            UpdateFile();
        }

        private List<SourceAddress> GetSourceAddresses()
        {
            var _addresses = new List<SourceAddress>();

            var wb = new XLWorkbook("source.xlsx");
            var ws = wb.Worksheets.FirstOrDefault();

            tbAddressesInSource.Text = (ws.RowsUsed().Count() - 3).ToString();
            // Парсим адреса из исходного файла
            foreach (var row in ws.Rows())
            {
                var name = row.Cell(1).Value.ToString();
                if (name.ToLower().Contains("гипер"))
                {
                    row.Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                    continue;
                }

                var address = row.Cell(2).Value.ToString();
                var pattern = @"(р-н .+?,)|(рп .+?,)|(г (?<city>.+?),)|(с .+?,)|((ул|ул|проезд|пр-кт|б-р|пл|пер|ш|наб|снт|мкр|тер|тракт)( им)? (?<street>.+?),)|(?<house>\d+[\dа-яА-Я -\/]*)";

                string street = "";
                string house = "";
                foreach (Match match in Regex.Matches(address, pattern, RegexOptions.IgnoreCase))
                {
                    if (!string.IsNullOrEmpty(match.Groups["street"].Value))
                        street = match.Groups["street"].Value.Trim();

                    if (!string.IsNullOrEmpty(match.Groups["house"].Value))
                        house = match.Groups["house"].Value.Trim();
                }

                if (!string.IsNullOrEmpty(street) && !string.IsNullOrEmpty(house))
                    _addresses.Add(new SourceAddress
                    {
                        Address = address,
                        Street = street,
                        House = house,
                        Summ = row.Cell(3).GetValue<double>(),
                        RowNumber = row.RowNumber()
                    });
            }

            wb.Save();
            return _addresses;
        }

        private bool ContainsInAddress(string address, string stringToFind)
        {
            return GetClearString(address).Contains(GetClearString(stringToFind));
        }

        private string GetClearString(string stringToClear)
        {
            const string charsToDelete = " -\"";

            foreach (var c in charsToDelete)
                stringToClear = stringToClear.Replace(c.ToString(), "");

            return stringToClear.ToLower();
        }

        private void FindMatches(List<SourceAddress> addresses)
        {
            var matches = addresses
                .Select(a => new SearchMatch
                {
                    SourceAdress = a,
                    ResultAddresses = new List<ResultAddress>()
                })
                .ToList();

            spMain.Children.Clear();
            var wb = new XLWorkbook("result.xlsx");
            var ws = wb.Worksheets.FirstOrDefault();

            // Узнаем номера нужных столбцов 
            var colNumberFilial = ws.Row(1).CellsUsed(c => c.GetString() == "Филиал").FirstOrDefault().WorksheetColumn().ColumnNumber();
            var colNumberAddress = ws.Row(1).CellsUsed(c => c.GetString() == "Адрес").FirstOrDefault().WorksheetColumn().ColumnNumber();
            var colNumberComment = ws.Row(1).CellsUsed(c => c.GetString() == "Адрес в системе").FirstOrDefault().WorksheetColumn().ColumnNumber();

            foreach (var row in ws.Rows())
            {
                var address = row.Cell(colNumberAddress).GetString();
                var comment = row.Cell(colNumberComment).GetString();

                if (row.Cell(colNumberFilial).Value.ToString().ToLower() != tbFilial.Text.ToLower())
                    continue;

                // Ищем соотвествие адресов
                foreach (var a in addresses)
                {
                    if ((!ContainsInAddress(address, a.Street) || !ContainsInAddress(address, a.House))
                        && (!ContainsInAddress(comment, a.Street) || !ContainsInAddress(comment, a.House)))
                        continue;

                    matches.First(m => m.SourceAdress == a)
                        .ResultAddresses
                        .Add(new ResultAddress
                        {
                            Address = address,
                            Comment = comment,
                            RowNumber = row.RowNumber()
                        });
                }
            }
            matches = matches.OrderByDescending(m => m.ResultAddresses.Count).ToList();
            tbMatches.Text = matches.Count(m => m.ResultAddresses.Count > 0).ToString();

            foreach (var m in matches.Where(m => m.ResultAddresses.Count > 1))
            {
                spMain.Children.Add(new MultiMatch
                {
                    SourceAddress = m.SourceAdress,
                    ResultAddresses = m.ResultAddresses
                });
            }

            foreach (var m in matches.Where(m => m.ResultAddresses.Count == 1))
            {
                spMain.Children.Add(new SingleMatch
                {
                    SourceAddress = m.SourceAdress,
                    ResultAddress = m.ResultAddresses.FirstOrDefault()
                });
            }

            foreach (var m in matches.Where(m => m.ResultAddresses.Count == 0))
            {
                spMain.Children.Add(new WithoutMatch
                {
                    SourceAddress = m.SourceAdress.Address
                });
            }
        }

        private void UpdateFile()
        {
            if (string.IsNullOrEmpty(tbMonth.Text))
            {
                MessageBox.Show("Введите месяц");
                return;
            }

            var wb = new XLWorkbook("result.xlsx");
            var ws = wb.Worksheets.FirstOrDefault();

            var wbSource = new XLWorkbook("source.xlsx");
            var wsSource = wbSource.Worksheets.FirstOrDefault();

            int columnNumber = 0;
            foreach (var column in ws.Columns())
            {
                if (column.Cell(1).Value.ToString().ToLower() == tbMonth.Text.ToLower())
                    columnNumber = column.ColumnNumber();
            }

            foreach (var control in spMain.Children)
            {
                if (control is MultiMatch)
                {
                    var mm = control as MultiMatch;
                    var selectedAddress = mm.ResultAddresses.FirstOrDefault(ra => ra.IsChecked);
                    if (selectedAddress == null)
                        continue;

                    var currentCell = ws.Cell(selectedAddress.RowNumber, columnNumber);
                    double currentValue = 0;
                    currentCell.TryGetValue<double>(out currentValue);
                    currentCell.SetValue<double>(currentValue + mm.SourceAddress.Summ);
                    wsSource.Cell(mm.SourceAddress.RowNumber, 3).Style.Fill.BackgroundColor = XLColor.Green;
                }

                if (control is SingleMatch)
                {
                    var sm = control as SingleMatch;
                    if (!sm.IsChecked)
                        continue;
                    var currentCell = ws.Cell(sm.ResultAddress.RowNumber, columnNumber);
                    double currentValue = 0;
                    currentCell.TryGetValue<double>(out currentValue);
                    currentCell.SetValue<double>(currentValue + sm.SourceAddress.Summ);
                    wsSource.Cell(sm.SourceAddress.RowNumber, 3).Style.Fill.BackgroundColor = XLColor.Green;
                }
            }

            wb.Save();
            wbSource.Save();
            MessageBox.Show("Обновление данных завершено.");
        }
    }

    public class SourceAddress
    {
        public string Address { get; set; }
        public string Street { get; set; }
        public string House { get; set; }
        public double Summ { get; set; }
        public int RowNumber { get; set; }
    }

    public class ResultAddress
    {
        public string Address { get; set; }
        public string Comment { get; set; }
        public int RowNumber { get; set; }
        public bool IsChecked { get; set; }
    }

    public class SearchMatch
    {
        public SourceAddress SourceAdress { get; set; }
        public List<ResultAddress> ResultAddresses { get; set; }
    }
}