#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel=Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace RevitAddin1
{
    [Transaction(TransactionMode.Manual)]
    public class Command03 : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.InitialDirectory = @"C:\";
            dialog.Multiselect = false;
            dialog.Filter = "Revit Files | *.rvt; *.rfa"; 

            string filePath = "";
            string[] filePaths;
            if (dialog.ShowDialog()==Forms.DialogResult.OK)  //only if user selects OK will an action go
            {
                filePath = dialog.FileName;
                //run if change multiselect = true
                //filePaths = dialog.FileNames;
            }

            Forms.FolderBrowserDialog folderDialog = new Forms.FolderBrowserDialog();

            string folderPath = "";
            if(folderDialog.ShowDialog()==Forms.DialogResult.OK)
            {
                folderPath = folderDialog.SelectedPath;
            }

            Tuple<string, int> t1 = new Tuple<string, int>("String 1",55);
            Tuple<string, int> t2 = new Tuple<string, int>("String 2", 24);

            TestStruct struct1;
            struct1.Name = "Structure 1";
            struct1.Value = 100;
            struct1.Value2 = 10.5;

            TestStruct struct2 = new TestStruct("structure 1", 10, 2.5);
        
            List<TestStruct> structlist = new List<TestStruct>();
            structlist.Add(struct1);

            

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType curVFT = null;
            ViewFamilyType curRCPVFT = null;
            foreach(ViewFamilyType curElem in collector)
            {
                if(curElem.ViewFamily == ViewFamily.FloorPlan)
                {
                    curVFT = curElem;
                }
                else if (curElem.ViewFamily == ViewFamily.CeilingPlan)
                {
                    curRCPVFT = curElem;
                }
            }

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector2.WhereElementIsElementType();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Stuff");

                Level newLevel = Level.Create(doc, 100);
                ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, newLevel.Id);
                ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, newLevel.Id);
                curRCP.Name = curRCP.Name + " RCP";

                ViewSheet newSheet = ViewSheet.Create(doc, collector2.FirstElementId());
                Viewport newVP = Viewport.Create(doc, newSheet.Id, curPlan.Id, new XYZ(0, 0, 0));

                newSheet.Name = "TEST SHEET";
                newSheet.SheetNumber = "S100";

                string paramValue = "";
                foreach(Parameter curParam in newSheet.Parameters)
                {
                    if(curParam.Definition.Name == "Drawn By")
                    {
                        curParam.Set("AMG");
                        paramValue = curParam.AsString();
                    }
                }

                t.Commit();
            }

          

            return Result.Succeeded;
        }

        internal View getviewbyname(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View));

            foreach(View curView in collector)
            {
                if(curView.Name == viewName)
                {
                    return curView;
                }
            }
            return null;
        }

        internal struct TestStruct
        {
            public string Name;
            public int Value;
            public double Value2;

            public TestStruct(string name, int value, double value2)
            {
                Name = name;
                Value = value;
                Value2 = value2;    
            }

        }
    }
}
