using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Media.Imaging;
using System.Reflection;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Dnbim_Tool.Properties;

namespace DnBim_Tool
{
    [Transaction(TransactionMode.Manual)]
    class Ribbon : IExternalApplication
    {
        private readonly string nameSpace = "DnBim_Tool.";
        private readonly string tabName = "DnBIM tool";
        private readonly string path = Assembly.GetExecutingAssembly().Location;

        private BitmapImage Convert(Bitmap bimapImage)
        {
            MemoryStream memory = new MemoryStream();
            bimapImage.Save(memory, ImageFormat.Png);
            memory.Position = 0;
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = memory;
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            return bitmapImage;
        }

        //a

        private void SetPull_Image(PulldownButton pull, Bitmap imageSource)
        {
            pull.LargeImage = Convert(imageSource);
        }

        private void SetPush_Image(PushButton push, Bitmap imageSource)
        {
            push.LargeImage = Convert(imageSource);
        }

        private void MyPush(RibbonPanel panel, PushButtonData data, Bitmap bitmap, string description)
        {
            PushButton push = panel.AddItem(data) as PushButton;
            push.ToolTip = description;
            SetPush_Image(push, bitmap);
        }


        private void MyPull(RibbonPanel panel, PulldownButtonData data, Bitmap bitmap, List<PushButtonData> list, string des, List<string> listdes)
        {
            PulldownButton pulldown = panel.AddItem(data) as PulldownButton;
            for (int i = 0; i < list.Count; i++)
            {
                PushButtonData pushdata = list[i];
                PushButton bt = pulldown.AddPushButton(pushdata);
                if (listdes.Count > 0)
                {
                    bt.ToolTip = listdes[i];
                }
            }
            SetPull_Image(pulldown, bitmap);
            pulldown.ToolTip = des;
        }


        private void MEP_Tool (RibbonPanel panel)
        {

            //place_Family
            PushButtonData place_FamilySupport = new PushButtonData("place_FamilySupport", "Auto Support", path, nameSpace + "PlaceSupportCmd");
            MyPush(panel, place_FamilySupport, Resources.Support, "Place family support.");

            //Connect element
            PushButtonData connnectFWA = new PushButtonData("Connect FWA", "Connect FWA", path, nameSpace + "ConnectingCmd");
            MyPush(panel, connnectFWA, Resources.Connect_FWA, "Connect auoto fitting with pipe accessory");

         
            PushButtonData Rotate = new PushButtonData("Rotate element", "Rotate element", path, nameSpace + "Rotate3DCmd");
            MyPush(panel, Rotate, Resources.Rotate3D_icon, "Xoay các đối tượng theo trục nhất định");


        }
        private void Duct_Tool(RibbonPanel panel)
        {

            PushButtonData SlpitDuct = new PushButtonData("Slpit Duct", "Auto Slpit Duct", path, nameSpace + "SliptDuctCMD");
            MyPush(panel, SlpitDuct, Resources.Slpit_Duct, "Auto slpit duct.");

            PushButtonData CalculateLength = new PushButtonData("CALCULATE LENGTH", "Tự động tính chiều dài đường ống", path, nameSpace + "CalculateLength");
            MyPush(panel, CalculateLength, Resources.calculate, "Tự động tính chiều dài đường ống.");


        }
        private void PlaceCassette_Tool(RibbonPanel panel)
        {


            PushButtonData placeCassette = new PushButtonData("Place Cassette", "Auto Place Cassette", path, nameSpace + "PlaceCassetteCmd");
            MyPush(panel, placeCassette, Resources.logo_khoa_nhiet, "Auto place element cassette");

            //a

        }



        public Result OnStartup(UIControlledApplication application)
        {

            //tạo tab có tên là "RevitAPI_ARC_2023"
            application.CreateRibbonTab(tabName);

            //tạo panel
            RibbonPanel mepPanel = application.CreateRibbonPanel(tabName, "MEP");

            RibbonPanel ductPanel = application.CreateRibbonPanel(tabName, "DUCT");
            RibbonPanel placecassettePanel = application.CreateRibbonPanel(tabName, "Excel");
            MEP_Tool(mepPanel);
            Duct_Tool(ductPanel);
            PlaceCassette_Tool(placecassettePanel);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
