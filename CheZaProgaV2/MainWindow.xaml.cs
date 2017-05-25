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
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
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
            SetSummToResult(addresses);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            UpdateFile();
        }

        private List<SourceAddress> GetSourceAddresses()
        {
            var _addresses = new List<SourceAddress>();

            //string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=source.xls;Extended Properties=Excel 8.0";

            //// Create the connection object 
            //OleDbConnection oledbConn = new OleDbConnection(connString);
            //try
            //{
            //    // Open connection
            //    oledbConn.Open();

            //    // Create OleDbCommand object and select data from worksheet Sheet1
            //    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [sheet$]", oledbConn);

            //    // Create new OleDbDataAdapter 
            //    OleDbDataAdapter oleda = new OleDbDataAdapter();

            //    oleda.SelectCommand = cmd;

            //    // Create a DataSet which will hold the data extracted from the worksheet.
            //    DataSet ds = new DataSet();
            //    oleda.Fill(ds);

            //    var dt = ds.Tables[0];

            //    int firstAddressIndex = 7;
            //    // Парсим адреса из исходного файла
            //    for (int i = firstAddressIndex; i < dt.Rows.Count; i++)
            //    {
            //        tbAddressesInSource.Text = (dt.Rows.Count - firstAddressIndex).ToString();

            //        var address = dt.Rows[i][1].ToString();
            //        var pattern = @"(р-н .+?,)|(рп .+?,)|(г (?<city>.+?),)|(ул (?<street>.+?),)|(с .+?,)|(проезд .+?,)|(пр-кт .+?,)|(?<house> \d+[\dа-яА-Я -]*)";

            //        string street = "";
            //        string house = "";
            //        foreach (Match match in Regex.Matches(address, pattern))
            //        {
            //            if (!string.IsNullOrEmpty(match.Groups["street"].Value))
            //                street = match.Groups["street"].Value;

            //            if (!string.IsNullOrEmpty(match.Groups["house"].Value))
            //                house = match.Groups["house"].Value;
            //        }

            //        if (!string.IsNullOrEmpty(street) && !string.IsNullOrEmpty(house))
            //            _addresses.Add(new Address
            //            {
            //                FullAddress = address,
            //                Street = street,
            //                House = house,
            //                Summ = dt.Rows[i][2].ToString()
            //            });
            //    }
            //}
            //catch { }

            //// Close connection
            //oledbConn.Close();

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
                var pattern = @"(р-н .+?,)|(рп .+?,)|(г (?<city>.+?),)|(с .+?,)|((ул|ул|проезд|пр-кт|б-р|пл|пер|ш)( им)? (?<street>.+?),)|(?<house> \d+[\dа-яА-Я -]*)";

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

        private void SetSummToResult(List<SourceAddress> addresses)
        {
            var matches = addresses
                .Select(a => new SearchMatch
                {
                    SourceAdress = a,
                    ResultAddresses = new List<ResultAddress>()
                })
                .ToList();

            //string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=result.xls;Extended Properties=Excel 8.0";

            //// Create the connection object 
            //OleDbConnection oledbConn = new OleDbConnection(connString);
            //try
            //{
            //    // Open connection
            //    oledbConn.Open();

            //    // Create OleDbCommand object and select data from worksheet Sheet1
            //    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Лист1$]", oledbConn);

            //    // Create new OleDbDataAdapter 
            //    OleDbDataAdapter oleda = new OleDbDataAdapter();

            //    oleda.SelectCommand = cmd;

            //    // Create a DataSet which will hold the data extracted from the worksheet.
            //    DataSet ds = new DataSet();
            //    oleda.Fill(ds);

            //    var dt = ds.Tables[0];

            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        var address = dt.Rows[i][3].ToString();

            //        if (dt.Rows[i][0].ToString() != filial)
            //            continue;

            //        bool isMatched = false;
            //        // Ищем соотвествие адресов
            //        foreach (var a in addresses)
            //        {
            //            if (!address.Contains(a.Street) || !address.Contains(a.House))
            //                continue;

            //            isMatched = true;
            //            matches.First(m => m.SourceAdress == a).ResultAddresses.Add(address);
            //        }
            //    }
            //    matches = matches.OrderByDescending(m => m.ResultAddresses.Count).ToList();

            //    foreach (var m in matches.Where(m => m.ResultAddresses.Count > 1))
            //    {
            //        spMain.Children.Add(new MultiMatch
            //        {
            //            SourceAddress = m.SourceAdress.FullAddress,
            //            ResultAddresses = m.ResultAddresses
            //        });
            //    }

            //    foreach (var m in matches.Where(m => m.ResultAddresses.Count == 1))
            //    {
            //        spMain.Children.Add(new SingleMatch
            //        {
            //            SourceAddress = m.SourceAdress.FullAddress,
            //            ResultAddress = m.ResultAddresses.FirstOrDefault()
            //        });
            //    }

            //    foreach (var m in matches.Where(m => m.ResultAddresses.Count == 0))
            //    {
            //        spMain.Children.Add(new WithoutMatch
            //        {
            //            SourceAddress = m.SourceAdress.FullAddress
            //        });
            //    }
            //}
            //catch { }
            //oledbConn.Close();

            spMain.Children.Clear();
            var wb = new XLWorkbook("result.xlsx");
            var ws = wb.Worksheets.FirstOrDefault();

            foreach (var row in ws.Rows())
            {
                var address = row.Cell(4).Value.ToString();

                if (row.Cell(1).Value.ToString().ToLower() != tbFilial.Text.ToLower())
                    continue;

                // Ищем соотвествие адресов

                foreach (var a in addresses)
                {
                    if (!address.Replace(" ", "").ToLower().Contains(a.Street.Replace(" ", "").ToLower())
                        || !address.Replace(" ", "").ToLower().Contains(a.House.Replace(" ", "").ToLower()))
                        continue;

                    matches.First(m => m.SourceAdress == a)
                        .ResultAddresses
                        .Add(new ResultAddress
                        {
                            Address = address,
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
            //string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=result.xls;Extended Properties=Excel 8.0";

            //// Create the connection object 
            //OleDbConnection oledbConn = new OleDbConnection(connString);
            //try
            //{
            //    // Open connection
            //    oledbConn.Open();


            //    foreach (var control in spMain.Children)
            //    {
            //        if (control is SingleMatch)
            //        {
            //            OleDbCommand cmd = new OleDbCommand($"UPDATE [Лист1$] SET (апрель = '123') WHERE Филиал = '{tbFilial.Text}' and Адрес = '{(control as SingleMatch).ResultAddress}';", oledbConn);
            //            //cmd.Parameters.Add("@var1", OleDbType.Double).Value = 123;
            //            var f = cmd.ExecuteNonQuery();
            //        }
            //    }
            //}
            //catch { }
            //oledbConn.Close();

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
        public int RowNumber { get; set; }
        public bool IsChecked { get; set; }
    }

    public class SearchMatch
    {
        public SourceAddress SourceAdress { get; set; }
        public List<ResultAddress> ResultAddresses { get; set; }
    }
}