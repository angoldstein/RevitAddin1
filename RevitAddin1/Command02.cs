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

#endregion

namespace RevitAddin1
{
    [Transaction(TransactionMode.Manual)]
    public class Command02 : IExternalCommand
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

            string excelfile = @"L:\COMMITTEES\Revit\Development\2022 ArchSmarter Add-in Academy\Session02 test.xlsx";

            Excel.Application excelapp = new Excel.Application();
            Excel.Workbook excelWB = excelapp.Workbooks.Open(excelfile);
            Excel.Worksheet excelWS = excelWB.Worksheets.Item[1];

            Excel.Range excelRng = excelWS.UsedRange;
            int rowcount = excelRng.Rows.Count;

            //do some stuff in Excel
            List<string[]> dataList = new List<string[]>();

            for(int i = 1; i <= rowcount; i++)
            {
                Excel.Range cell1 = excelWS.Cells[i, 1];
                Excel.Range cell2 = excelWS.Cells[i, 2];

                string data1 = cell1.Value.tostring();
                string data2 = cell2.Value.tostring();

                string[] dataArray = new string[2];
                dataArray[0] = data1;
                dataArray[1] = data2;

                dataList.Add(dataArray);
            }

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create some revit stuff");
                Level curLevel = Level.Create(doc, 100);

                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                collector.WhereElementIsElementType();

                ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                curSheet.SheetNumber = "S101";
                curSheet.Name = "New Sheet";

                t.Commit();
            }
          
          
                
            excelWB.Close();
            excelapp.Quit();

            return Result.Succeeded;
        }
    }
}
