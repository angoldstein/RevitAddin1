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
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Forms = System.Windows.Forms;


#endregion

namespace RevitAddin1
{
    [Transaction(TransactionMode.Manual)]
    public class Command05Challenge : IExternalCommand
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

            int counter = 0;

            string excelFile = "";

            Forms.OpenFileDialog dialog = new Forms.OpenFileDialog();
            dialog.Title = "Select Furnuture Excel File";
            dialog.InitialDirectory = @"C:\";
            dialog.Multiselect = false;
            dialog.Filter = "Excel Files | *.xlsx; *.xlsm; *.xls | All Files | *.*";


            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                TaskDialog.Show("Error", "Please select an Excel File");
                return Result.Failed;
            }

            excelFile = dialog.FileName;

            List<string[]> excelFurnSetData = GetDataFromExcel(excelFile, "Furniture sets", 3);
            List<string[]> excelFurnData = GetDataFromExcel(excelFile, "Furniture types", 3);

            excelFurnSetData.RemoveAt(0); //remove header
            excelFurnData.RemoveAt(0);

            List<FurnSet> furnSetList = new List<FurnSet>();
            List<FurnData> furnDataList = new List<FurnData>();

            foreach (string[] curRow in excelFurnSetData)
            {
                FurnSet tmpFurnset = new FurnSet(curRow[0].Trim(), curRow[1].Trim(), curRow[2].Trim());
                furnSetList.Add(tmpFurnset);
            }

            foreach (string[] curRow in excelFurnData)
            {
                FurnData tmpFurnData = new FurnData(doc, curRow[0].Trim(), curRow[1].Trim(), curRow[2].Trim());
                furnDataList.Add(tmpFurnData);
            }

            List<SpatialElement> roomList = GetAllRooms(doc);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert Furniture");

                foreach (SpatialElement room in roomList)
                {
                    string curFurnSet = GetParamValue(room, "Furniture Set");
                    LocationPoint roomPt = room.Location as LocationPoint;
                    XYZ insPoint = roomPt.Point;

                    foreach (FurnSet tmpFurnSet in furnSetList)
                    {
                        if (tmpFurnSet.setType == curFurnSet)
                        {
                            foreach (string curFurn in tmpFurnSet.furnList)
                            {
                                string tmpFurn = curFurn.Trim();
                                FurnData fd = GetFamilyInfo(tmpFurn, furnDataList);

                                if (fd != null && fd.familySymbol != null)
                                {                                    
                                    fd.familySymbol.Activate();

                                    FamilyInstance newFamInst = doc.Create.NewFamilyInstance(insPoint, fd.familySymbol, StructuralType.NonStructural);
                                    counter++;
                                }
                            }
                        }

                        SetParamValueAsInt(room, "Furniture Count", tmpFurnSet.furnList.Count);
                    }
                }

                t.Commit();
            }

            TaskDialog.Show("Complete", "Inserted " + counter + " families");
            return Result.Succeeded;
        }

        private void SetParamValueAsInt(Element curElem, string paramName, int paramValue)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name == paramName)
                    curParam.Set(paramValue);
            }
        }

        private FurnData GetFamilyInfo(string furnName, List<FurnData> furnDataList)
        {
            foreach(FurnData furn in furnDataList)
            {
                if (furn.furnName == furnName)
                    return furn;
            }
            return null;
        }

        private string GetParamValue(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                Debug.Print(curParam.Definition.Name);
                if (curParam.Definition.Name == paramName)
                {
                    Debug.Print(curParam.AsString());
                    return curParam.AsString();
                }
            }

            return null;
        }

        private List<SpatialElement> GetAllRooms(Document doc)
        {
            List<SpatialElement> returnList = new List<SpatialElement>();

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);
            collector.WhereElementIsNotElementType();

            foreach(Element curElem in collector)
            {
                SpatialElement curRoom = curElem as SpatialElement;
                returnList.Add(curRoom);
            }

            return returnList;

        }

        private List<string[]> GetDataFromExcel(string excelFile, string wsName, int numColumns)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Open(excelFile);

            Excel.Worksheet excelWS = GetExcelWorksheetByName(excelWB, wsName);
            Excel.Range excelRng = excelWS.UsedRange as Excel.Range;

            int rowCount = excelRng.Rows.Count;

            List<string[]> data = new List<string[]>();

            for (int i = 1; i <= rowCount; i++)
            {
                string[] rowData = new string[numColumns];

                for(int j = 1; j <= numColumns; j++)
                {
                    Excel.Range cellData = excelWS.Cells[i, j];
                    rowData[j - 1] = cellData.Value.ToString(); //need j-1 to get 0 value item for C#
                }
                data.Add(rowData);
            }

            excelWB.Close();
            excelApp.Quit();

            return data;
        }

        private Excel.Worksheet GetExcelWorksheetByName(Excel.Workbook excelWB, string wsName)
        {
            foreach(Excel.Worksheet sheet in excelWB.Worksheets)
            {
                if (sheet.Name == wsName)
                    return sheet;
            }
            return null;
        }
    }   
}
