using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

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
        private List<Tuple<string, int>> articleDescriptions = new List<Tuple<string, int>>();

        private Dictionary<string, string> columnNames = new Dictionary<string, string>
        {
            { "description", "Beschrijving" },
            { "reject", "Aantal afkeur" },
            { "rewash", "Aantal overwas" },
            { "linen", "Vreemd linnen" },
            { "returns", "Aantal retouren" },
            { "comments", "Opmerking" }
        };

        private Dictionary<string, string> columns = new Dictionary<string, string>
        {
            { "B", "description" },
            { "D", "reject" },
            { "E", "rewash" },
            { "F", "linen" },
            { "G", "returns" },
            { "H", "comments" }
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

                    articleDescriptions.Add(new Tuple<string, int>(Model.GetCellValue(docName, sheetName, addressName), i));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Oops!");
            }

            foreach (Tuple<string, int> description in articleDescriptions)
            {
                articles.Add(new Article(description.Item1, description.Item2));
            }

            dataGrid.ItemsSource = articles;
        }

        private void submitButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Update from the static textboxes
                Model.InsertText(docName, sheetName, "B", 6, datumTextBox.Text);
                Model.InsertText(docName, sheetName, "H", 6, containersTextBox.Text);
                Model.InsertText(docName, sheetName, "H", 8, zakkenTextBox.Text);

                // Update from the DataGrid
                foreach (var article in articles)
                {
                    foreach (var column in columns)
                    {
                        string columnName = column.Key;
                        uint row = Convert.ToUInt32(article.GetRow());
                        string value = article.GetProperty(column.Value);
                        Model.InsertText(docName, sheetName, columnName, row, value);
                    }
                }

                MessageBox.Show("Bestand is opgeslagen.");

                Model.SendMail(docName);
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
        public Article(string description, int row)
        {
            this.description = description;
            _row = row;
        }
        
        public string description { get; }
        public int reject { get; set; }
        public int rewash { get; set; }
        public int linen { get; set; }
        public int returns { get; set; }
        public string comments { get; set; }

        private int _row;

        public int GetRow()
        {
            return _row;
        }

        public string GetProperty(string name)
        {
            switch (name)
            {
                case "description":
                    return description;
                case "reject":
                    return reject.ToString();
                case "rewash":
                    return rewash.ToString();
                case "linen":
                    return linen.ToString();
                case "returns":
                    return returns.ToString();
                case "comments":
                    return comments;
                default:
                    return null;
            }
        }
    }
}
