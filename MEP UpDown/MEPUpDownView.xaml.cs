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
using Autodesk.Revit.UI;
using DnBim_Tool;
using DnBim_Tool.Properties;
using Microsoft.Vbe.Interop;
using Window = System.Windows.Window;
namespace Dnbim_Tool.MEP_UpDown
{
    /// <summary>
    /// Interaction logic for MEPUpDownView.xaml
    /// </summary>
    /// 
    //a
    public partial class MEPUpDownView : Window
    {
        
        public  string LastSelectedOption = "Cut Up";
        public  string LastAngle = "45°";
        public  double LastOffset = 1000;
        private ExternalEvent DuctEvent;
        public string tentabcontrol { get; set; }


        private static string LastSelectedTab = "Pipe";  // Giá trị mặc định ban đầu
        public MEPUpDownView()
        {
            InitializeComponent();

            var listOption = new List<string>() { "Cut Up", "Cut Down", "Move Up", "Move Down" };
            cbboption.ItemsSource = listOption;
            cbboption.SelectedItem = LastSelectedOption;

            var angle = new List<string>() { "45°", "90°" };
            cbbangle.ItemsSource = angle;
            cbbangle.SelectedItem = LastAngle;

            // CHỖ NÀY SỬA:
            tboffset.Text = LastOffset.ToString(); // dùng giá trị cuối cùng người dùng đã nhập





        }
        private void BingdingImage(string option, string angle)
        {
            string findOption = $"{option}-{angle}";
            List<BitmapImage> images = new List<BitmapImage>()
            {
                
                //CT.Convert(Resources.CutElbowDown45Img),
                //CT.Convert(Resources.CutElbowUp90Img),
                //CT.Convert(Resources.CutElbowDown45Img),
                //CT.Convert(Resources.CutElbowDown90Img),


                //CT.Convert(Resources.MoveUpElbow45Img),
                //CT.Convert(Resources.MoveUpElbow90Img),
                //CT.Convert(Resources.MoveDownElbow45Img),
                //CT.Convert(Resources.MoveDownElbow90Img),


            };
            List<PictureItem> pictureItems = new List<PictureItem>();

            int i = 0;
            foreach (string op in cbboption.Items)
            {
                foreach (string goc in cbbangle.Items)
                {
                    string name = $"{op}-{goc}";
                    PictureItem pictureItem = new PictureItem(images[i], name);
                    pictureItems.Add(pictureItem);
                    i++;
                }
            }
            PictureItem item = pictureItems.Find(x => x.Familyname == findOption);
            Image.Source = item.Image;//Image ni là của bên wpf


        }


        public string option;
        public string angle;
        public double Offset;
        private void cbboption_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                option = cbboption.SelectedValue.ToString();
                 angle = cbbangle.SelectedValue.ToString();
                BingdingImage(option, angle);
            }
            catch { }
        }
        private void cbbangle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                option = cbboption.SelectedValue.ToString();
                angle = cbbangle.SelectedValue.ToString();
                BingdingImage(option, angle);
            }
            catch { }
            //a
        }

        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
       
       

        private void btOk_Click_1(object sender, RoutedEventArgs e)
        {
            option = cbboption.SelectedValue.ToString();
            LastSelectedOption = option;
            angle = cbbangle.SelectedValue.ToString();
            LastAngle = angle;
            bool isNumber = double.TryParse(tboffset.Text, out double number);
            if (!isNumber)
            {
                MessageBox.Show("Enter a number!");
                tboffset.Text = LastOffset.ToString(); // quay về giá trị cũ nếu sai
            }
            else
            {
                Offset = number;
                LastOffset = number; // LƯU lại giá trị người dùng đã nhập

                DialogResult = true;
            }

        }
        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                btOk_Click_1(sender, e); // Gọi logic của Button "OK"
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
