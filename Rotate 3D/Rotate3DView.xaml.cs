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
using System.Xml.Linq;
using Autodesk.Revit.DB;
//using DnBim_Tool.Properties;

namespace DnBim_Tool.Rotate_3D
{
    /// <summary>
    /// Interaction logic for Rotate3DView.xaml
    /// </summary>
    public partial class Rotate3DView : Window
    {
        public double Angle { get; set; }
        private static double previousAngle = 0;
        public Rotate3DView()
        {

            InitializeComponent();
            BindingImage();
            tbAngle.Text = previousAngle.ToString();
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

                    btCancel_Click(sender, e);
                }
            }
        }

        private void btOk_Click(object sender, RoutedEventArgs e)
        {
            

            bool check2 = double.TryParse(tbAngle.Text , out double giatri2);
            if (check2 == true)
            {
                previousAngle = giatri2;
                Angle = Math.Round(giatri2, 2);
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("Góc quay chưa đúng hãy nhập lại góc quay", "Cảnh báo!");
            }
        }

        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void BindingImage()
        {
            List<BitmapImage> images = new List<BitmapImage>()
            //{ CT.Convert(DnBim_Tool.Properties.Resources.Rotate3D) };
            { CT.Convert(Dnbim_Tool.Properties.Resources.Rotate3D) };
            List<PictureItem> pictureItems = new List<PictureItem>();
            PictureItem pictureItem = new PictureItem(images[0]);
            pictureItems.Add(pictureItem);
            PictureItem item = pictureItems[0];
            //image.Source = item.Image;
        }
    }
}
