using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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
using ExcelDataReader;
using Microsoft.Win32;

namespace Журнал
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IExcelDataReader edr;
        public MainWindow()
        {
            InitializeComponent();

        }

        private void Zagruz_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL files (*.xlsx)|.xslx|EXCEL Files (*.xls)| *.xls|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() != true)
            {
                return;
            }
            Baza.ItemsSource = readFile(openFileDialog.FileName);
        }

        private DataView readFile(string fileNames)
        {
            var extencion = fileNames.Substring(fileNames.LastIndexOf('.'));
            FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
            if (extencion == ".xlsx")
            {
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            if (extencion == ".xls")
            {
                edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            var conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            DataSet dataSet = edr.AsDataSet(conf);
            DataView dataView = dataSet.Tables[0].AsDataView();
            edr.Close();
            return dataView;
        }
        private void saveFile()
        {
            //var ext = fileNames.Substring(fileNames.LastIndexOf('.'));
            //FileStream stream = File.Sa(fileNames, FileMode.Open, FileAccess.Read);
            //if (extencion == ".xlsx")
            //{
            //    edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //}
            //if (extencion == ".xls")
            //{
            //    edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //}
            //var conf = new ExcelDataSetConfiguration
            //{
            //    ConfigureDataTable = _ => new ExcelDataTableConfiguration
            //    {
            //        UseHeaderRow = true
            //    }
            //};
            //DataSet dataSet = edr.AsDataSet(conf);
            //DataView dataView = dataSet.Tables[0].AsDataView();
            //edr.Close();
            //return dataView;
            var dialog = new SaveFileDialog();
            dialog.Filter = "EXCEL files (*.xlsx)|.xlsx|EXCEL Files (*.xls)| *.xls|All files (*.*)|*.*";

            var result = dialog.ShowDialog();

            if (result is null) return;
            if (result == false) return;

            using (OleDbConnection connection = new OleDbConnection())
            {
                connection.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dialog.FileName};" +
                                          "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                connection.Open();
                using (var command = new OleDbCommand())
                {
                    command.Connection = connection;
                    var columnNames = Baza.Columns.Select(x => x.Header).ToList();
                    var tableName = Guid.NewGuid().ToString();
                    command.CommandText = 
                        $"CREATE TABLE [{tableName}] " +
                        $"({string.Join(",", columnNames.Select(c => $"[{c}] VARCHAR"))})";
                    command.ExecuteNonQuery();
                    foreach (DataRow row in (Baza.ItemsSource as DataView).Table.Rows)
                    {
                        var rowValues = (from DataColumn column in (Baza.ItemsSource as DataView).Table.Columns select (row[column] != null && row[column] != DBNull.Value) ? row[column].ToString() : string.Empty).ToList();
                        command.CommandText = $"INSERT INTO [{tableName}]({string.Join(",", columnNames.Select(c => $"[{c}]"))}) VALUES ({string.Join(",", rowValues.Select(r => $"'{r}'").ToArray())});";
                        command.ExecuteNonQuery();
                    }
                }
                connection.Close();
            }

        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Baza.SelectAllCells();
            Baza.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
            ApplicationCommands.Copy.Execute(null, Baza);
            Baza.UnselectAllCells();
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "EXCEL files (*.xlsx)|.xsls|EXCEL Files (*.xls)| *.xls|All files (*.*)|*.*";
            if (saveFileDialog1.ShowDialog() != true)
            {
                return;
            }
            Baza.ItemsSource = readFile(saveFileDialog1.FileName);

        }

       private void Button1_Click(object sender, RoutedEventArgs e)
        {
            saveFile();
        }
       
    }
    
}
