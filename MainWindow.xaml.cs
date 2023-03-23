using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
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
            openFileDialog.Filter = "EXCEL files (*.xlsx)|.xsls|EXCEL Files (*.xls)| *.xls|All files (*.*)|*.*";
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

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "EXCEL files (*.xlsx)|.xsls|EXCEL Files (*.xls)| *.xls|All files (*.*)|*.*";

            if (saveFileDialog1.ShowDialog() == true)
            {
 
                
            }
        }
    }
}
