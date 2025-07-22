using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ExcelDataReader;
using MessageBox = System.Windows.MessageBox;

namespace Dnbim_Tool.Placecassette
{
    /// <summary>
    /// Interaction logic for PlaceCassetteView.xaml
    /// </summary>
    public partial class PlaceCassetteView : Window
    {
        private ExternalEvent SheetEvent;
        private Document doc;

        public string filepath;

        public string sheetNames;
        public string sheetNamesisSelected;

      
        public PlaceCassetteView()
        {
            InitializeComponent();
            this.doc = doc;
            //cbb_TitleBlocks.ItemsSource = GetListTitleBlocks();
            //cbb_TitleBlocks.SelectedIndex = 0;
            var option = new List<string>() { "Kinh Tế", "Vận hành" };
            cbbOption.ItemsSource = option;
            cbbOption.SelectedIndex = 0;
        }
       

        private void bt_Browse_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select Excel File";
            dialog.Filter = "Excel Files| *xls; *xlsx; *xlsm";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tb_FilePath.Text = dialog.FileName;
                cbbSheetName.ItemsSource = GetExcelSheetNames(tb_FilePath.Text);
                cbbSheetName.SelectedIndex = 0;
            }
            else tb_FilePath.Text = "";
        }
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                bt_Ok_Click(sender, e); // Gọi logic của Button "OK"
            }

            else
            {
                if (e.Key == Key.Escape)
                {

                    bt_Cancel_Click(sender, e);
                }
            }
        }
        public List<string> GetExcelSheetNames(string filePath)
        {
            List<string> sheetNames = new List<string>();

            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Chọn file excel hoặc copy/paste đường dẫn!", "Message");
            }
            else
            {
                if (filePath.Contains("\"")) filePath = filePath.Replace("\"", "");

                System.Data.DataTable excelData = new System.Data.DataTable();
                using (FileStream fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(fileStream))
                    {
                        var data = reader.AsDataSet();
                        // Lấy tên các sheet
                        foreach (System.Data.DataTable table in data.Tables)
                        {
                            sheetNames.Add(table.TableName); // Tên sheet tương ứng với TableName

                        }

                    }
                }

            }
            return sheetNames;
        }
        private List<string> GetListTitleBlocks()
        {
            var collector = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .WhereElementIsElementType()
                        .Cast<FamilySymbol>()
                        .ToList();

            var listNames = (collector.Select(x => $"{x.FamilyName}: {x.Name}")).ToList();
            listNames.Sort();
            return listNames;

        }

        private void bt_Ok_Click(object sender, RoutedEventArgs e)
        {
            //string blockName = cbb_TitleBlocks.SelectedValue.ToString();
            //ElementId blockId = GetBlockId(blockName);
            string filepath = tb_FilePath.Text.ToString();
            string sheetNamesisSelected = cbbSheetName.SelectedValue.ToString();



            if (string.IsNullOrEmpty(filepath))
            {
                MessageBox.Show("Chọn file excel hoặc copy/paste đường dẫn!", "Message");
            }
            else
            {
                DialogResult = true;

            }
        }

        private void bt_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
