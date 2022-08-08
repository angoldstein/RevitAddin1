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
    internal class App06Challenge : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
           // step 1: create ribbon tab
            try
            {
                a.CreateRibbonTab("Revit Add-in Academy");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists");
            }

           // step 2: create ribbon panel            
            RibbonPanel curPanel = CreateRibbonPanel(a, "Revit Add-in Academy", "Revit Tools");

            // step 3: create button data instances
            PushButtonData pData1 = new PushButtonData("button 1", "Fizz Buzz", GetAssemblyName(), "RevitAddin1.FizzBuzz");
            PushButtonData pData2 = new PushButtonData("button 2", "Project Setup", GetAssemblyName(), "RevitAddin1.Command02Challenge");
            PushButtonData pData3 = new PushButtonData("button 3", "Project Setup v2", GetAssemblyName(), "RevitAddin1.Command03Challenge");
            PushButtonData pData4 = new PushButtonData("button 4", "Elements From Lines", GetAssemblyName(), "RevitAddin1.Command04Challenge");
            PushButtonData pData5 = new PushButtonData("button 5", "Insert Furniture", GetAssemblyName(), "RevitAddin1.Command05Challenge");
            PushButtonData pData6 = new PushButtonData("button 6", "Button 6", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData7 = new PushButtonData("button 7", "Button 7", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData8 = new PushButtonData("button 8", "Button 8", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData9 = new PushButtonData("button 9", "Button 9", GetAssemblyName(), "RevitAddin1.Command01");
            PushButtonData pData10 = new PushButtonData("button 10", "Button 10", GetAssemblyName(), "RevitAddin1.Command01");

            SplitButtonData sData1 = new SplitButtonData("splitButton1", "Split Button 1");
            PulldownButtonData pbData1 = new PulldownButtonData("pulldownButton1", "More Tools");

            // step 4: add images
            pData1.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_16);
            pData1.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_32);

            pData2.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_16);
            pData2.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_32);

            pData3.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_16);
            pData3.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_32);

            pData4.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_16);
            pData4.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_32);

            pData5.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.icons8_semi_truck_side_view_16);
            pData5.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.icons8_semi_truck_side_view_32);

            pData6.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_16);
            pData6.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_32);

            pData7.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_16);
            pData7.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Red_32);

            pData8.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_16);
            pData8.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Yellow_32);

            pData9.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_16);
            pData9.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Blue_32);

            pData10.Image = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_16);
            pData10.LargeImage = BitmaptoImageSource(RevitAddin1.Properties.Resources.Green_32);

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
            pData8.ToolTip = "Button 8 tool tip";
            pData9.ToolTip = "Button 9 tool tip";
            pData10.ToolTip = "Button 10 tool tip";

            pbData1.ToolTip = "Group of tools";

            // step 6: create buttons
            PushButton B1 = curPanel.AddItem(pData1) as PushButton;
            PushButton B2 = curPanel.AddItem(pData2) as PushButton;

            curPanel.AddStackedItems(pData3, pData4, pData5);

            SplitButton splitButton1 = curPanel.AddItem(sData1) as SplitButton;
            splitButton1.AddPushButton(pData6);
            splitButton1.AddPushButton(pData7);

            PulldownButton pulldownButton1 = curPanel.AddItem(pbData1) as PulldownButton;
            pulldownButton1.AddPushButton(pData8);
            pulldownButton1.AddPushButton(pData9);
            pulldownButton1.AddPushButton(pData10);

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
