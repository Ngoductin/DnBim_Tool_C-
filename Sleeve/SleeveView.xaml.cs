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
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace Dnbim_Tool.Sleeve
{
    /// <summary>
    /// Interaction logic for SleeveView.xaml
    /// </summary>
    public partial class SleeveView : Window
    {
        public bool isCeilingsChecked { get; set; }
        public bool isFloorsChecked { get; set; }
        public bool isWallsChecked { get; set; }
        public bool isPipeChecked { get; set; }
        public bool isDuctChecked { get; set; }
        public bool isCableTrayChecked { get; set; }
        public SleeveView()
        {
            InitializeComponent();
            
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
            DialogResult = true;
            isCeilingsChecked = cellings.IsChecked ?? false;
            isFloorsChecked = floors.IsChecked ?? false;
            isWallsChecked = walls.IsChecked ?? false;
            isPipeChecked = pipe.IsChecked ?? false;
            isDuctChecked = duct.IsChecked ?? false;
            isCableTrayChecked = cabletray.IsChecked ?? false;
        }
    }
}
