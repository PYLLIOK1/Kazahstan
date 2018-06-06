using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using HtmlAgilityPack;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using System.Net;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Diagnostics;

namespace Kazahstan
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        private static string DelProb( string str)
        {
            string newstr = "";
            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] != ' ')
                {
                    newstr += str[i];
                }
            }
            return newstr;
        }

    List<Parse> Liste = new List<Parse>();
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string kz = ".kz";
            string ru = ".ru";
            WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
            string s = client.DownloadString("http://www.e-talgar.com/node/7588");
            client.Dispose();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(s);
            HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//tr[@class='c12']");
            if (all != null)
            {
                foreach (HtmlNode n in all)
                {
                    HtmlDocument rows = new HtmlDocument();
                    rows.LoadHtml(n.InnerHtml);
                    HtmlNodeCollection cell = rows.DocumentNode.SelectNodes("//td");
                    if ((cell[1] != null) && (cell[3] != null))
                    {
                        string name = cell[1].InnerText;
                        name = name.Replace('\n',' ');
                        name = name.Replace("&nbsp;", "");
                        string email = cell[3].InnerText;
                        email = email.Replace('\n', ' ');
                        email = DelProb(email);
                        email = email.Replace("&nbsp;", "");
                        email = email.Replace("Е-mail:", "");
                        email = email.Replace("E-mail:", "");
                        email = email.Replace("..", ".");
                        if (name != "  ")
                        {
                            if (email != "")
                            {
                                if(!email.Contains("Ректор"))
                                {
                                    int m = email.IndexOf(ru);
                                        while ((m != -1) && (email.Length > (m+ru.Length)))
                                        {
                                            email = email.Remove(m + ru.Length);
                                            m = email.IndexOf(ru, m + ru.Length - 1);
                                        }
                                    int l = email.IndexOf(kz);
                                    while ((l != -1) && (email.Length > (l + kz.Length)))
                                    {
                                        email = email.Remove(l + kz.Length);
                                        l = email.IndexOf(kz, l + kz.Length - 1);
                                    }
                                    Liste.Add(new Parse { Name = name, Email = email });
                                }
                            }                               
                        }  
                    }
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = (XSSFSheet)wb.CreateSheet("Лист 1");
            int countColumn = 2;
            for (int i = 0; i < Liste.Count; i++)
            {
                var currentRow = sh.CreateRow(i);
                for (int j = 0; j < countColumn; j++)
                {
                    var currentCell = currentRow.CreateCell(j);
                    if(j==0)
                    {
                        currentCell.SetCellValue(Liste[i].Name);
                    }
                    if(j==1)
                    {
                        currentCell.SetCellValue(Liste[i].Email);
                    }
                    sh.AutoSizeColumn(j);
                }

            }
            if (!File.Exists("d:\\etalgarcomnode7588.xlsx"))
            {
                File.Delete("d:\\etalgarcomnode7588.xlsx");
            }
            using (var fs = new FileStream("d:\\etalgarcomnode7588.xlsx", FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }
            Process.Start("d:\\etalgarcomnode7588.xlsx");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
            string s = client.DownloadString("http://egov.kz/cms/ru/articles/2Fvusi_rk");
            client.Dispose();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(s);
            HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//div[@class='slidedown-content']");
            if (all != null)
            {
                foreach (HtmlNode n in all)
                {

                }
            }
        }
    }
    public class Parse
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }
}
