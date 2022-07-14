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
            dialog.Filter = "Revit Files | *.rvt; *.rfa";

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

            //do some stuff in Excel

            for (int i = 2; i <= rowcountsh; i++)
            {
                Excel.Range sheetname = excelWSsheet.Cells[i, 2];

                string datashname = sheetname.Value.ToString();
                                
                Excel.Range sheetnum = excelWSsheet.Cells[i, 1];
                              
                string datashnum = sheetnum.Value.ToString();

         
                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Sheets");
                                       
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
                    collector.WhereElementIsElementType();

                    ViewSheet curSheet = ViewSheet.Create(doc, collector.FirstElementId());
                    curSheet.SheetNumber = datashnum;
                    curSheet.Name = datashname;

                    t.Commit();
                }
            }

            for (int ii = 2; ii <= rowcountlev; ii++)
            {
                Excel.Range levelelev = excelWSlevel.Cells[ii, 2];

                double dataelev = levelelev.Value;

                Excel.Range levelname = excelWSlevel.Cells[ii, 1];

                string dataname = levelname.Value.ToString();


                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Create Levels");

                    Level curLevel = Level.Create(doc, dataelev);
                    curLevel.Name = dataname;
                                     
                    t.Commit();
                }
            }

            excelWB.Close();
            excelapp.Quit();

            return Result.Succeeded;
        }
    }
}
