#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace RevitAddin1
{
    [Transaction(TransactionMode.Manual)]
    public class Command03Challenge : IExternalCommand
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
            dialog.Filter = "Excel Files | *.xlsx; *.xlsm; *.xls | All Files | *.*";

            string filePath = "";
            if (dialog.ShowDialog() == Forms.DialogResult.OK)  
            {
                filePath = dialog.FileName;                
            }

            
            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook excelWB = excelapp.Workbooks.Open(filePath);
            Excel.Worksheet excelWSlevel = excelWB.Worksheets.Item[1];
            Excel.Worksheet excelWSsheet = excelWB.Worksheets.Item[2];

            Excel.Range excelRnglevel = excelWSlevel.UsedRange;
            int rowcountlev = excelRnglevel.Rows.Count;

            Excel.Range excelRngSh = excelWSsheet.UsedRange;
            int rowcountsh = excelRngSh.Rows.Count;

                     
            //create levels & views
            
            for (int ii = 2; ii <= rowcountlev; ii++)
            {
                Excel.Range levelelev = excelWSlevel.Cells[ii, 2];
             
                Excel.Range levelname = excelWSlevel.Cells[ii, 1];
                
                levels levelstruct = new levels(levelelev.Value, levelname.Value.tostring());

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(ViewFamilyType));

                ViewFamilyType curVFT = null;
                ViewFamilyType curRCPVFT = null;
                foreach (ViewFamilyType curElem in collector)
                {
                    if (curElem.ViewFamily == ViewFamily.FloorPlan)
                    {
                        curVFT = curElem;
                    }
                    else if (curElem.ViewFamily == ViewFamily.CeilingPlan)
                    {
                        curRCPVFT = curElem;
                    }
                }

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Levels");

                    Level curLevel = Level.Create(doc, levelstruct.Elev);
                    curLevel.Name = levelstruct.Name;

                    ViewPlan curPlan = ViewPlan.Create(doc, curVFT.Id, curLevel.Id);
                    ViewPlan curRCP = ViewPlan.Create(doc, curRCPVFT.Id, curLevel.Id);
                    curRCP.Name = curRCP.Name + " RCP";

                    t.Commit();
                }
            }

            //create sheets

            for (int i = 2; i <= rowcountsh; i++)
            {
                Excel.Range sheetnum = excelWSsheet.Cells[i, 1];
                
                Excel.Range sheetname = excelWSsheet.Cells[i, 2];
                                                
                Excel.Range sheetview = excelWSsheet.Cells[i, 3];
                
                Excel.Range drawnby = excelWSsheet.Cells[i, 4];
               
                Excel.Range checkby = excelWSsheet.Cells[i, 5];
              
                sheets sheetstruct = new sheets(sheetnum.Value.ToString(), sheetname.Value.ToString(), sheetview.Value.ToString(), drawnby.Value.ToString(), checkby.Value.ToString());

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Sheets");

                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                    collector.WhereElementIsElementType();

                    ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    curSheet.SheetNumber = sheetstruct.Number;
                    curSheet.Name = sheetstruct.Name;

                    View curView = GetViewByName(doc, sheetstruct.View);
                                                         

                    Viewport newVP = Viewport.Create(doc, curSheet.Id, curView.Id, new XYZ(1, .5, 0));

                    string paramDrawnby = "";
                    foreach (Parameter curParam in curSheet.Parameters)
                    {
                        if (curParam.Definition.Name == "Drawn By")
                        {
                            curParam.Set(sheetstruct.Drawnby);
                            paramDrawnby = curParam.AsString();
                        }
                    }

                    string paramCheckby = "";
                    foreach (Parameter curParam in curSheet.Parameters)
                    {
                        if (curParam.Definition.Name == "Checked By")
                        {
                            curParam.Set(sheetstruct.Checkby);
                            paramCheckby = curParam.AsString();
                        }
                    }

                    t.Commit();
                }
            }


            excelWB.Close();
            excelapp.Quit();

            return Result.Succeeded;
        }

        internal struct levels
        {
            public string Name;
            public double Elev;

            public levels (string name, double elev)
            {
                Name = name;    
                Elev = elev;
            }
            
        }

        internal struct sheets
        {
            public string Number;
            public string Name;
            public string View;
            public string Drawnby;
            public string Checkby;

            public sheets (string number, string name, string view, string drawnby, string checkby)
            {
                Number = number;
                Name = name;
                View = view;
                Drawnby = drawnby;
                Checkby = checkby;
            }

        }

    }
}
