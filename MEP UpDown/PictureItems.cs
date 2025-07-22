using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace Dnbim_Tool.MEP_UpDown
{
    public class PictureItem
    {
        public BitmapImage Image { get; set; }
        public string Familyname { get; set; }

        public PictureItem(BitmapImage bitmapImage, string familyname)
        {
            Image = bitmapImage;
            Familyname = familyname;

        }


    }
}
