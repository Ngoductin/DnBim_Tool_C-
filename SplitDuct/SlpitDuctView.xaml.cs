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

namespace DnBim_Tool
{
    /// <summary>
    /// Interaction logic for SlpitDuctView.xaml
    /// </summary>
    public partial class SlpitDuctView : Window
    {
        public string SplitDuct { get; set; }
        public double Distance { get; set; }
        public SlpitDuctView()
        {
            InitializeComponent();

            /*ComboBox*/
            //List<string> luachon = new List<string>() { "Slipt Duct From Start Point", "Slipt Duct From Mid Point", "Slipt Duct From End Point" };
            List<string> luachon = new List<string>() { "START OR END","MID"};
            cbbLuachon.ItemsSource = luachon;
            cbbLuachon.SelectedIndex = 0;

            /*TextBox*/
            tbDistance.Text = "1110";
        }
        

        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btOK_Click(object sender, RoutedEventArgs e)
        {
            SplitDuct = cbbLuachon.SelectedValue.ToString();//kiểu trả về của phương thức selectedValue là object nên phải ép kiểu  nó về string
            string value = tbDistance.Text;
            bool check = double.TryParse(value, out double Giatri);
            if (check)
            {
                Distance = Giatri;

            }
            else
            {
                MessageBox.Show("Nhập số vào", "Message");

            }
            DialogResult = true;

        }
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btOK_Click(sender, e); // Gọi logic của Button "OK"
            }
            
            else
            {
                if (e.Key == Key.Escape)
                {
                    
                    btCancel_Click(sender, e);
                }
            }
        }
    }
}
