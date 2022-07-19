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
    public class Command03Challengecopy : IExternalCommand
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

            
            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel File");
                return Result.Failed;
            }

            string filePath = dialog.FileName;
            int levelcounter = 0;
            int sheetcounter = 0;

            try
            {
                Excel.Application excelapp = new Excel.Application();
                Excel.Workbook excelWB = excelapp.Workbooks.Open(filePath);

                Excel.Worksheet excelWSlevel = GetExcelWorksheetByName(excelWB, "Levels");
                Excel.Worksheet excelWSsheet = GetExcelWorksheetByName(excelWB, "Sheets");

                List<levels> levelData = GetLevelDataFromExcel(excelWSlevel);
                List<sheets> sheetData = GetSheetDataFromExcel(excelWSsheet);

                excelWB.Close();
                excelapp.Quit();

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Levels");

                    ViewFamilyType planVFT = GetViewFamilyType(doc, "plan");
                    ViewFamilyType RCPVFT = GetViewFamilyType(doc, "rcp");

                    foreach (levels curLevel in levelData)
                    {
                        Level newLevel = Level.Create(doc, curLevel.LevelElev);
                        newLevel.Name = curLevel.LevelName;
                        levelcounter++;

                        ViewPlan curFloorPlan = ViewPlan.Create(doc, planVFT.Id, newLevel.Id);
                        ViewPlan curRCP = ViewPlan.Create(doc, RCPVFT.Id, newLevel.Id);

                        curRCP.Name = curRCP.Name + " RCP";

                    }

                    FilteredElementCollector collector = GetTitleblock(doc);


                    foreach (sheets curSheet in sheetData)
                    {
                        ViewSheet newSheet = ViewSheet.Create(doc, collector.FirstElementId());

                        newSheet.SheetNumber = curSheet.SheetNumber;
                        newSheet.Name = curSheet.SheetName;

                        SetParamValue(newSheet, "Drawn By", curSheet.Drawnby);
                        SetParamValue(newSheet, "Checked By", curSheet.Checkby);

                        View curView = GetViewByName(doc, curSheet.SheetView);

                        if (curView != null)
                        {
                            Viewport curVP = Viewport.Create(doc, newSheet.Id, curView.Id, new XYZ(1, 1, 0));
                        }

                        sheetcounter++;

                    }

                    t.Commit();
                }
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }

            TaskDialog.Show("Complete", "Created " + levelcounter.ToString() + " levels.");
            TaskDialog.Show("Complete", "Created " + sheetcounter.ToString() + " sheets.");

            return Result.Succeeded;
        }

        private void SetParamValue(ViewSheet newSheet, string paramname, string paramvalue)
        {
            foreach (Parameter curParam in newSheet.Parameters)
            {
                if (curParam.Definition.Name == paramname)
                {
                    curParam.Set(paramvalue );
                    
                }
            }
        }
        private static FilteredElementCollector GetTitleblock(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            collector.WhereElementIsElementType();
            return collector;
        }

        private ViewFamilyType GetViewFamilyType(Document doc, string type)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewFamilyType));


            foreach (ViewFamilyType vft in collector)
            {
                if (vft.ViewFamily == ViewFamily.FloorPlan && type == "plan")
                {
                    return vft;
                }
                else if (vft.ViewFamily == ViewFamily.CeilingPlan && type == "rcp")
                {
                    return vft;
                }
            }

            return null;
        }

        private List<sheets> GetSheetDataFromExcel(Excel.Worksheet excelWSsheet)
        {
            List<sheets> returnlist = new List<sheets>();
            Excel.Range excelRngSh = excelWSsheet.UsedRange;
            
            int rowcountsh = excelRngSh.Rows.Count;

            for (int i = 2; i <= rowcountsh; i++)
            {
                Excel.Range sheetnum = excelWSsheet.Cells[i, 1];

                Excel.Range sheetname = excelWSsheet.Cells[i, 2];

                Excel.Range sheetview = excelWSsheet.Cells[i, 3];

                Excel.Range drawnby = excelWSsheet.Cells[i, 4];

                Excel.Range checkby = excelWSsheet.Cells[i, 5];

                sheets curSheet = new sheets();
                curSheet.SheetNumber = sheetname.Value.ToString();
                curSheet.SheetName = sheetname.Value.ToString();
                curSheet.SheetView = sheetview.Value.ToString();
                curSheet.Drawnby = drawnby.Value.ToString();
                curSheet.Checkby = checkby.Value.ToString();

                returnlist.Add(curSheet);
            }

            return returnlist;
        }

        private List<levels> GetLevelDataFromExcel(Excel.Worksheet excelWSlevel)
        {
            List<levels> returnlist = new List<levels>();
            Excel.Range excelRnglevel = excelWSlevel.UsedRange;
            
            int rowcountlev = excelRnglevel.Rows.Count;

            for (int ii = 2; ii <= rowcountlev; ii++)
            {
                Excel.Range levelelev = excelWSlevel.Cells[ii, 2];

                Excel.Range levelname = excelWSlevel.Cells[ii, 1];
                                
                double LevelElev = levelelev.Value;
                string LevelName = levelname.Value.ToString();

                levels curLevel = new levels(LevelName, LevelElev);

                returnlist.Add(curLevel);
            }
            return returnlist;
        }

        private View GetViewByName(Document doc, string viewName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Views);
                
            foreach(View curView in collector)
            {
                if (curView.Name == viewName)
                    return curView;            
            }

            return null;
        }

        private Excel.Worksheet GetExcelWorksheetByName(Excel.Workbook excelWB, string wsName)
        {
            foreach(Excel.Worksheet worksheet in excelWB.Worksheets)
            {
                if(worksheet.Name == wsName)
                {
                    return worksheet;
                }
            }

            return null;
        }

        private struct levels
        {
            public string LevelName;
            public double LevelElev;

            public levels (string name, double elev)
            {
                LevelName = name;
                LevelElev = elev;
            }
            
        }

        private struct sheets
        {
            public string SheetNumber;
            public string SheetName;
            public string SheetView;
            public string Drawnby;
            public string Checkby;

            public sheets (string number, string name, string view, string drawnby, string checkby)
            {
                SheetNumber = number;
                SheetName = name;
                SheetView = view;
                Drawnby = drawnby;
                Checkby = checkby;
            }

        }

    }
}
