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
using System.IO.Packaging;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace Blycolin_Print_and_Mail
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string docName = @"testSheet.xlsx";
        private const string sheetName = "Blad1";
        private List<Article> articles = new List<Article>();
        private List<string> articleDescriptions = new List<string>();

        private Dictionary<string, string> columnNames = new Dictionary<string, string>
        {
            { "description", "Beschrijving" },
            { "reject", "Aantal afkeur" },
            { "rewash", "Aantal overwas" },
            { "linen", "Vreemd linnen" },
            { "returns", "Aantal retouren" },
            { "comments", "Opmerking" }
        };

        public MainWindow()
        {
            InitializeComponent();

            datumTextBox.Text = DateTime.Now.ToShortDateString();

            try
            {
                for (int i = 14; i <= 20; i++)
                {
                    string addressName = "B" + i;

                    articleDescriptions.Add(Model.GetCellValue(docName, sheetName, addressName));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Oops!");
            }

            foreach (string description in articleDescriptions)
            {
                articles.Add(new Article(description));
            }

            dataGrid.ItemsSource = articles;
        }

        private void submitButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Model.InsertText(docName, sheetName, "B", 6, "Tekst vanuit programma");
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message, "Oops!");
            }
        }

        //Access and update columns during autogeneration
        private void AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            string headername = e.Column.Header.ToString();

            if (columnNames[headername] != null)
            {
                //update column details when generating
                if (headername == "comments")
                {
                    e.Column.Header = columnNames[headername];
                }
                else
                {
                    e.Column.Header = columnNames[headername];
                    e.Column.Width = new DataGridLength(1.0, DataGridLengthUnitType.Auto);
                }
            }
        }
    }

    public class Article
    {
        public Article(string description)
        {
            this.description = description;
        }
        
        public string description { get; }
        public int reject { get; set; }
        public int rewash { get; set; }
        public int linen { get; set; }
        public int returns { get; set; }
        public string comments { get; set; }
    }
}
