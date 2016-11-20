using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Windows;
using CsQuery;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Parser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        List<Ad> ads = new List<Ad>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void Parse(object sender, RoutedEventArgs e)
        {
            await Task.Run(() =>
            {
                var baseUrl = System.Configuration.ConfigurationManager.AppSettings["baseUrl"];
                var numberPages = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["numberOfPages"]);
                for (var pNum = 1; pNum <= numberPages; pNum++)
                {
                    var page = new Uri(baseUrl + pNum);
                    var html = new WebClient() { Encoding = UTF8Encoding.UTF8 }
                                .DownloadString(page);
                    CQ cq = CQ.Create(html);
                    var headerElems = cq.Find(".all_advert .sa_content.realty_content .sa_header");

                    foreach (var elem in headerElems)
                    {
                        var name = elem.Cq().Find(".title_realty").Text().Trim(' ', '\n');
                        var price = elem.Cq().Find(".price_realty").Text().Trim(' ', '\n');
                        ads.Add(new Ad(name, price));
                    }
                }

                this.CreateExcelAsync();
            });      
        }

        public async void CreateExcelAsync()
        {
            await Task.Run(() =>
            {
                var app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = true;
                Workbook book = app.Workbooks.Add(Type.Missing);
                Worksheet sheet = (Worksheet)book.Worksheets.get_Item(1);
                sheet.Cells[1, 1] = "Name";
                sheet.Cells[1, 2] = "Price";

                for (var i = 0; i < ads.Count; i++)
                {
                    sheet.Cells[i + 2, 1] = ads[i].Name;
                    sheet.Cells[i + 2, 2] = ads[i].Price;
                }

                book.SaveAs(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("../../file{0}.xls", DateTime.Now.ToString("hh_mm_ss_tt"))), XlFileFormat.xlWorkbookNormal);
                book.Close(true);
                app.Quit();

                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(book);
                Marshal.ReleaseComObject(app);
            });           
        }
    }
}
