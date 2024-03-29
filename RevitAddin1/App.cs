#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;

#endregion

namespace RevitAddin1
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
           // step 1: create ribbon tab
            try
            {
                a.CreateRibbonTab("Test Tab");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

           // step 2: create ribbon panel            
            RibbonPanel curPanel = CreateRibbonPanel(a, "Test Tab", "Test Panel");

            // step 3: create button data instances
            PushButtonData pData1 = new PushButtonData("button 1", "This is /rButton 1", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData2 = new PushButtonData("button 2", "Button 2", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData3 = new PushButtonData("button 3", "Button 3", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData4 = new PushButtonData("button 4", "Button 4", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData5 = new PushButtonData("button 5", "Button 5", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData6 = new PushButtonData("button 6", "Button 6", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData7 = new PushButtonData("button 7", "Button 7", GetAssemblyName(), "RevitAddin1.Command01");

            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button 1");
            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "Pulldown" +"/r" + "Button 1");

            // step 4: add images
            pData1.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_16);
            pData1.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_32);

            pData2.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_16);
            pData2.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_32);

            pData3.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_16);
            pData3.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_32);

            pData4.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_16);
            pData4.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_32);

            pData5.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_16);
            pData5.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_32);

            pData6.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_16);
            pData6.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_32);

            pData7.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_16);
            pData7.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_32);

            pbData1.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_16);
            pbData1.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_32);

            // step 5: add tool tips
            pData1.ToolTip = "Button 1 tool tip";
            pData2.ToolTip = "Button 2 tool tip";
            pData3.ToolTip = "Button 3 tool tip";
            pData4.ToolTip = "Button 4 tool tip";
            pData5.ToolTip = "Button 5 tool tip";
            pData6.ToolTip = "Button 6 tool tip";
            pData7.ToolTip = "Button 7 tool tip";

            pbData1.ToolTip = "Group of tools";

            // step 6: create buttons
            PushButton B1 = curPanel.AddItem(pData1) as PushButton;

            curPanel.AddStackedItems(pData2, pData3);

            SplitButton splitButton1 = curPanel.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData4);
            splitButton1.AddPushButton(pData5);

            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(pData6);
            pulldownButton1.AddPushButton(pData7);
                       
            return Result.Succeeded;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication a, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in a.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }
            RibbonPanel returnPanel = a.CreateRibbonPanel(tabName, panelName);

            return returnPanel;
        }

        private BitmapImage BitmaptoImageSource(System.Drawing.Bitmap bm)
        {
           using(MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
