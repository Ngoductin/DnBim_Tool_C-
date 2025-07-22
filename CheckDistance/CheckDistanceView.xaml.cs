using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using DnBim_Tool;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Color = Autodesk.Revit.DB.Color;
using MessageBox = System.Windows.MessageBox;
using Window = System.Windows.Window;

namespace Dnbim_Tool.CheckDistance
{
    /// <summary>
    /// Interaction logic for CheckDistanceView.xaml
    /// </summary>
    /// 
    

    public partial class CheckDistanceView : Window
    {

        // Biến instance để lưu trữ fi1Id và fi2Id
        private ElementId storedFi1Id;
        private ElementId storedFi2Id;
        // Biến để lưu trữ các ElementId đã được highlight
        public ElementId highlightedFi1Id;
        public ElementId highlightedFi2Id;
        private ExternalEvent TreeNodeEvent; // Đổi tên biến để rõ ràng hơn
        private TreeNodeEvent TreeNodeEventHandler; // Lưu trữ handler
        private ExternalEvent CloseEvent; // Đổi tên biến để rõ ràng hơn
        private ExternalEvent RefreshEvent; // Sự kiện để gọi RefreshEvent
        private RefreshEvent RefreshEventHandler; // Handler cho RefreshEvent


        private ExternalEvent ShowEvent; // Đổi tên biến để rõ ràng hơn
        private ShowEvent ShowEventHandler; // Lưu trữ handler
        public List<Hashtable> data; // Dữ liệu đầu vào
        private UIDocument _uidoc;
        private Document _doc;
        private Dictionary<string, bool> _expandedStates; // Lưu trạng thái mở rộng của các node
        public CheckDistanceView(UIDocument uidoc, Document doc, List<Hashtable> inputData)
        //public CheckDistanceView()
        {

            InitializeComponent();
            _uidoc = uidoc;
            _doc = doc;
            _expandedStates = new Dictionary<string, bool>(); // Khởi tạo dictionary để lưu trạng thái
            DateTime currentTime = DateTime.Now;
           var timecreated= currentTime.ToString("dddd, MMMM d, yyyy h:mm:ss tt");
            cbb_Daycreated.Text = timecreated;
            cbb_Dayupdate.Text = string.Empty;
           
            this.data = inputData; // Lưu dữ liệu đầu vào
            PopulateTreeView(); // Đổ dữ liệu vào TreeView
            Topmost = true;

            // Gắn cửa sổ vào Revit để cải thiện quản lý focus
            Loaded += (s, e) =>
            {
                IntPtr revitHandle = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
                new WindowInteropHelper(this).Owner = revitHandle;
                Focus();
            };

          
            //event
            TreeNodeEventHandler = new TreeNodeEvent { window = this };
            TreeNodeEvent = ExternalEvent.Create(TreeNodeEventHandler);


            ShowEventHandler = new ShowEvent { window = this };
            ShowEvent = ExternalEvent.Create(ShowEventHandler);

            CloseEvent closeEvent = new CloseEvent { window = this };
            CloseEvent = ExternalEvent.Create(closeEvent);

            // Khởi tạo RefreshEvent
            RefreshEventHandler = new RefreshEvent { window = this };
            RefreshEvent = ExternalEvent.Create(RefreshEventHandler);
        }
        public void UpdateData(List<Hashtable> newData)
        {
            SaveExpandedStates(); // Lưu trước khi làm mới
            this.data = newData;
            PopulateTreeView(); // TreeView sẽ được làm mới, sau đó RestoreExpandedStates sẽ được gọi


        }
        private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                CloseButton_Click(sender, e);
                e.Handled = true;
                Console.WriteLine("Escape Pressed - Window Closed");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            CloseEvent.Raise();
            Close();
            //a
        }
        //private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        //{
        //    if (e.Key == Key.Enter)
        //    {
        //        //btOk_Click_1(sender, e); // Gọi logic của Button "OK"
        //    }

        //    else
        //    {
        //        if (e.Key == Key.Escape)
        //        {

        //            Close();
        //        }
        //    }
        //}
        public void myTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {


           
                TreeNode selectedNode = e.NewValue as TreeNode;
                if (selectedNode != null)
                {
                    var (fi1Id, fi2Id) = selectedNode.ExtractIdsIfContainsDistance();

                    if (fi1Id != null && fi1Id is ElementId)
                    {
                        // Truyền dữ liệu ElementId vào handler
                        TreeNodeEventHandler.Fi1Id = fi1Id;
                        TreeNodeEventHandler.Fi2Id = fi2Id;
                    // Lưu giá trị fi1Id và fi2Id vào biến instance
                    storedFi1Id = fi1Id;
                    storedFi2Id = fi2Id;

                    TreeNodeEvent.Raise(); // Kích hoạt sự kiện
                    }
                    else
                    {
                      
                    }
                }

        }
       
        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Lấy thời gian hiện tại
            DateTime currentTime = DateTime.Now;

            // Định dạng thời gian theo kiểu "Friday, April 11, 2025 8:37:02 PM"
            string formattedTime = currentTime.ToString("dddd, MMMM d, yyyy h:mm:ss tt");

            // Cập nhật TextBox cbb_Dayupdate
            cbb_Dayupdate.Text = formattedTime;
            // Kích hoạt RefreshEvent để lấy dữ liệu mới
            RefreshEvent.Raise();

        }
        private void PopulateTreeView()
        {
            List<TreeNode> rootNodes = new List<TreeNode>();

            foreach (Hashtable hashtable in data)
            {
                foreach (DictionaryEntry entry in hashtable)
                {
                    string key = entry.Key.ToString();
                    var values = entry.Value as IEnumerable<object>;

                    TreeNode rootNode = new TreeNode(key);
                    rootNodes.Add(rootNode);

                    if (values != null)
                    {
                        foreach (var value in values)
                        {
                            TreeNode childNode = new TreeNode(value.ToString());
                            rootNode.Children.Add(childNode);
                        }
                    }
                }
            }

            // Gán dữ liệu mới vào TreeView
            myTreeView.ItemsSource = rootNodes;

            // Khôi phục trạng thái mở rộng
            foreach (TreeNode node in rootNodes)
            {
                RestoreExpandedStates(node);
            }
        }
        public class TreeNode : INotifyPropertyChanged
        {
            //public string Name { get; set; }
            //public object Key { get; set; }
            //public List<TreeNode> Children { get; set; }
            //public bool IsExpanded { get; set; } // Thêm thuộc tính IsExpanded

            //public TreeNode(string name, object key = null)
            //{
            //    Name = name;
            //    Key = key;
            //    Children = new List<TreeNode>();
            //    IsExpanded = false; // Mặc định không mở rộng
            //}
            public string Name { get; set; }
            public object Key { get; set; }

            private bool _isExpanded;
            public bool IsExpanded
            {
                get => _isExpanded;
                set
                {
                    _isExpanded = value;
                    OnPropertyChanged(nameof(IsExpanded));
                }
            }

            public List<TreeNode> Children { get; set; }

            public TreeNode(string name, object key = null)
            {
                Name = name;
                Key = key;
                Children = new List<TreeNode>();
                IsExpanded = false;
            }

            public event PropertyChangedEventHandler PropertyChanged;

            protected void OnPropertyChanged(string name)
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
            }

            public (ElementId Fi1Id, ElementId Fi2Id) ExtractIdsIfContainsDistance()
            {
                ElementId fi1Id = null;
                ElementId fi2Id = null;

                if (string.IsNullOrEmpty(Name))
                {
                    return (null, null);
                }

                var idMatches = Regex.Matches(Name, @"Id (\d+)");
                if (!Name.Contains("Distance"))
                {
                    if (idMatches.Count > 0)
                    {
                        fi1Id = new ElementId(int.Parse(idMatches[0].Groups[1].Value));
                        if (idMatches.Count > 1)
                        {
                            fi2Id = new ElementId(int.Parse(idMatches[1].Groups[1].Value));
                        }
                    }
                }
                else
                {
                    if (idMatches.Count > 0)
                    {
                        fi1Id = new ElementId(int.Parse(idMatches[0].Groups[1].Value));
                        if (idMatches.Count > 1)
                        {
                            fi2Id = new ElementId(int.Parse(idMatches[1].Groups[1].Value));
                        }
                    }
                }

                return (fi1Id, fi2Id);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            
           ShowEventHandler.Fi1Id=storedFi1Id;  
           ShowEventHandler.Fi2Id=storedFi2Id;
            


            ShowEvent.Raise();
        }
        // Lưu trạng thái mở rộng của các node
        private void SaveExpandedStates()
        {
            _expandedStates.Clear(); // Xóa trạng thái cũ
            if (myTreeView.ItemsSource != null)
            {
                foreach (TreeNode node in myTreeView.ItemsSource)
                {
                    SaveExpandedStateRecursive(node);
                }
            }
        }

        private void SaveExpandedStateRecursive(TreeNode node)
        {
            // Lưu trạng thái của node hiện tại
            _expandedStates[node.Name] = node.IsExpanded;

            // Đệ quy để lưu trạng thái của các node con
            foreach (var child in node.Children)
            {
                SaveExpandedStateRecursive(child);
            }
        }

        // Khôi phục trạng thái mở rộng
        private void RestoreExpandedStates(TreeNode node)
        {
            if (_expandedStates.TryGetValue(node.Name, out bool isExpanded))
            {
                node.IsExpanded = isExpanded;
            }

            foreach (var child in node.Children)
            {
                RestoreExpandedStates(child);
            }
        }
    }
}
