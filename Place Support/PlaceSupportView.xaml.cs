using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Drawing;
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
//using DnBim_Tool.Properties;


namespace DnBim_Tool.Place_Support
{
    /// <summary>
    /// Interaction logic for PlaceSupportView.xaml
    /// </summary>
    public partial class PlaceSupportView : Window
    {
        private ExternalEvent DuctCutDownEvent;
        private ExternalEvent PipeEvent;
        private ExternalEvent CableTrayCutDownEvent;
        // Biến để lưu giá trị cũ
        private static string LastSelectedTab = "Pipe";  // Giá trị mặc định ban đầu
        private static string LastSelectedFamily = null;
        private static string LastSelectedType = null;
        private static string LastSelectedOption = "Auto";
        private static double LastDistance = 500;
        private static double LastOffset = 1200;
        private Document doc;
        public string tentabcontrol { get; set; }
        public double Distance { get; set; }
        public double Offset { get; set; }
       

        public bool check = true;
       
        public PlaceSupportView(Document doc)
        {
            InitializeComponent();
            this.doc = doc;

            var listOption = new List<string>() { "Auto", "Manual" };
            cbbOption.ItemsSource = listOption;
            cbbOption.SelectedItem = LastSelectedOption; // Set giá trị mặc định từ biến

            tbdtt.Text = LastDistance.ToString(); // Set giá trị mặc định
            tbOffset.Text = LastOffset.ToString(); // Set giá trị mặc định

            cbbFamily.ItemsSource = GetFamilySupport(doc);
            if (LastSelectedFamily != null && cbbFamily.Items.Contains(LastSelectedFamily))
            {
                cbbFamily.SelectedItem = LastSelectedFamily; // Set giá trị mặc định nếu có
            }
            else
            {
                cbbFamily.SelectedIndex = 0; // Nếu không có, dùng giá trị đầu tiên
            }

            if (LastSelectedType != null)
            {
                cbbType.ItemsSource = GetlistTypeFamilySupport(doc, LastSelectedFamily);
                cbbType.SelectedItem = LastSelectedType; // Set giá trị mặc định nếu có
            }

            SelectTabByName(LastSelectedTab);

        }
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btOk_Click(sender, e); // Gọi logic của Button "OK"
            }

            else
            {
                if (e.Key == Key.Escape)
                {

                    this.Close();
                }
            }
        }




        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        private void btOk_Click(object sender, RoutedEventArgs e)
        {

            string value = tbdtt.Text;
            string number = tbOffset.Text;

            bool check1 = double.TryParse(value, out double giatri1);
            bool check2 = double.TryParse(number, out double giatri2);
            if (check1 == true && check2 == true)
            {
                Distance = giatri1;
                Offset = giatri2;
                DialogResult = true;

                TabItem tabItem = tabControl.SelectedItem as TabItem;
                tentabcontrol = tabItem.Header.ToString();
                LastSelectedTab=tentabcontrol;
                // Cập nhật giá trị lưu trữ
                LastDistance = Distance;
                LastOffset = Offset;

                LastSelectedOption = cbbOption.SelectedItem.ToString();
                LastSelectedFamily = cbbFamily.SelectedItem?.ToString();
                LastSelectedType = cbbType.SelectedItem?.ToString();

            }
            else
            {
                MessageBox.Show("Enter a number!", "Message");
            }



        }
        private void cbbFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {




            var listtypename = GetlistTypeFamilySupport(doc, cbbFamily.SelectedValue.ToString());


            cbbType.ItemsSource = listtypename;
            cbbType.SelectedIndex = 0;


        }

        private void SelectTabByName(string tabName)
        {
            foreach (TabItem tab in tabControl.Items)
            {
                if (tab.Header.ToString() == tabName)
                {
                    tabControl.SelectedItem = tab; // Chọn TabItem
                    break;
                }
            }
        }









        public static Dictionary<string, HashSet<string>> GetFamilyTypeNamesWithFilter(Document doc)
        {
            // Dictionary lưu trữ Family Name và HashSet các Type Name tương ứng
            Dictionary<string, HashSet<string>> familyTypeNames = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            // Lọc tất cả các FamilyInstance trong tài liệu
            var allFamilyInstances = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType() // Chỉ lấy các FamilyInstance
                .OfClass(typeof(FamilyInstance)) // Chỉ lấy FamilyInstance
                .Cast<FamilyInstance>()
                .ToList();

            // Duyệt qua tất cả các FamilyInstance và lấy tên Family từ Symbol
            foreach (var familyInstance in allFamilyInstances)
            {
                string familyName = familyInstance.Symbol.Family.Name;  // Lấy tên Family từ FamilyInstance
                string typeName = familyInstance.Symbol.Name;  // Lấy Type Name của FamilySymbol

                // Kiểm tra nếu Family Name chứa "hanger" hoặc "support" (không phân biệt hoa thường)
                if (familyName.ToLower().Contains("hanger") || familyName.ToLower().Contains("support"))
                {
                    // Nếu Family chưa có trong Dictionary, tạo mới HashSet
                    if (!familyTypeNames.ContainsKey(familyName))
                    {
                        familyTypeNames[familyName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // Tạo HashSet mới cho FamilyName
                    }

                    // Thêm Type Name vào HashSet của Family
                    familyTypeNames[familyName].Add(typeName);
                }
            }

            // Lọc tất cả các Family trong tài liệu
            var allFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family)) // Lấy tất cả các Family trong tài liệu
                .Cast<Family>()
                .ToList();

            // Duyệt qua các Family chưa được thêm vào Dictionary và thêm các Type Name của chúng
            foreach (var family in allFamilies)
            {
                string familyName = family.Name;

                // Kiểm tra nếu Family Name chứa "hanger" hoặc "support" (không phân biệt hoa thường)
                if (familyName.ToLower().Contains("hanger") || familyName.ToLower().Contains("support"))
                {
                    // Nếu Family chưa có trong Dictionary, tạo mới HashSet
                    if (!familyTypeNames.ContainsKey(familyName))
                    {
                        familyTypeNames[familyName] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    }

                    // Lấy các FamilySymbol từ Family này
                    var familySymbols = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol)) // Chọn FamilySymbol
                        .Cast<FamilySymbol>()
                        .Where(symbol => symbol.Family.Id == family.Id) // Kiểm tra FamilySymbol có cùng Family ID
                        .ToList();

                    // Duyệt qua tất cả các FamilySymbol của Family này và thêm Type Name vào HashSet
                    foreach (FamilySymbol symbol in familySymbols)
                    {
                        string typeName = symbol.Name;
                        familyTypeNames[familyName].Add(typeName); // Thêm Type Name vào HashSet của Family
                    }
                }
            }

            // Trả về Dictionary chứa các Family Name và HashSet các Type Name tương ứng
            return familyTypeNames;
        }

        private List<string> GetlistTypeFamilySupport(Document doc, string a)
        {
            List<string> ListTypeName = new List<string>();

            Dictionary<string, HashSet<string>> familyTypeNames = GetFamilyTypeNamesWithFilter(doc);
            // Duyệt qua tất cả các entry trong Dictionary
            foreach (var entry in familyTypeNames)
            {
                string familyName = entry.Key;  // Tên Family (key)
                HashSet<string> typeNames = entry.Value;  // Các Type Name của Family

                // Kiểm tra nếu key trong Dictionary khớp với giá trị a
                if (familyName == a)
                {
                    List<string> list = typeNames.ToList();
                    foreach (string type in list)
                    {
                        ListTypeName.Add(type);
                    }

                }
            }
            return ListTypeName;
        }
        private List<string> GetFamilySupport(Document doc)
        {
            // Lấy tất cả các FamilyInstance từ mọi category
            var listFamily = new FilteredElementCollector(doc)
       .WhereElementIsNotElementType() // Lọc bỏ ElementType
       .OfType<FamilyInstance>()       // Chỉ lấy các phần tử là FamilyInstance
       .Where(x => x.Symbol.Family.Name.ToLower().Contains("hanger") || x.Symbol.Family.Name.ToLower().Contains("support"))
       .Select(x => x.Symbol.Family)
       .ToList();

            // Lấy tất cả các Family trong tài liệu
            var allFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family)) // Lấy tất cả Family
                .Cast<Family>()
                .Where(family => family.Name.ToLower().Contains("hanger") || family.Name.ToLower().Contains("support"))
                .ToList();

            // Tạo một HashSet để theo dõi các tên Family đã gặp
            HashSet<string> familyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Kết hợp danh sách Family từ FamilyInstance và tất cả Family
            var combinedFamilyList = new List<Family>();

            foreach (var family in listFamily.Concat(allFamilies))
            {
                if (familyNames.Add(family.Name)) // Nếu tên chưa có trong HashSet
                {
                    combinedFamilyList.Add(family);
                }
            }

            // Trả về danh sách tên Family duy nhất
            return combinedFamilyList.Select(f => f.Name).ToList();
        }


    }
}
