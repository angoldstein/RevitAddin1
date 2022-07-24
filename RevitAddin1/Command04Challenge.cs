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
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB.Architecture;

#endregion

namespace RevitAddin1
{
    [Transaction(TransactionMode.Manual)]
    public class Command04Challenge : IExternalCommand
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


            IList<Element> pickList = uidoc.Selection.PickElementsByRectangle("Select some elements");
            List<CurveElement> curveList = new List<CurveElement>();

            WallType curWallType = GetWallTypeByName(doc, @"Generic - 8""");
            WallType storeWallType = GetWallTypeByName(doc, "Storefront");
            Level curLevel = GetLevelByName (doc, "Level 1");
            MEPSystemType curSystemtype = GetSystemTypeByName(doc, "Domestic Hot Water");
            PipeType curPipeType = GetPipeTypeByName(doc, "Default");
            MEPSystemType ductSystemType = GetSystemTypeByName(doc, "Supply Air");
            DuctType curDuctType = GetDuctTypeByName(doc, "Default");

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create Revit Stuff");

                foreach (Element element in pickList)
                {
                    if (element is CurveElement)
                    {
                        CurveElement curve = (CurveElement)element;
                        
                        curveList.Add(curve);

                        GraphicsStyle curGS = curve.LineStyle as GraphicsStyle;
                        Curve curCurve = null;
                        XYZ startpoint=null, endpoint = null;

                        /*switch (curGS.Name)
                        {
                            case "A-GLAZ":
                            case "A-WALL":                             
                            case "M-DUCT":                              
                            case "P-PIPE":
                                curCurve = curve.GeometryCurve;
                                startpoint = curCurve.GetEndPoint(0);
                                endpoint = curCurve.GetEndPoint(1);

                                break;
                        }
                        */

                        try
                        {
                            startpoint = curCurve.GetEndPoint(0);
                            endpoint = curCurve.GetEndPoint(1);
                        }
                        catch
                        {
                            Debug.Print("no endpoints");
                        }

                        switch (curGS.Name)
                        {
                            case "A-GLAZ":
                                Wall newstoreWall = Wall.Create(doc, curCurve, storeWallType.Id, curLevel.Id, 15, 0, false, false);
                                break;

                            case "A-WALL":
                                Wall newWall = Wall.Create(doc, curCurve, curWallType.Id, curLevel.Id, 15, 0, false, false);
                                break;

                            case "M-DUCT":
                                Duct newDuct = Duct.Create(doc, ductSystemType.Id, curDuctType.Id, curLevel.Id, startpoint, endpoint);
                                break;

                            case "P-PIPE":
                                Pipe newPipe = Pipe.Create(doc, curSystemtype.Id, curPipeType.Id, curLevel.Id, startpoint, endpoint);
                                break;

                            default:
                                Debug.Print("found something else");
                                break;

                        }

                        

                                                

                        Debug.Print(curGS.Name);

                    }
                }
                t.Commit();
            }
            
           
            TaskDialog.Show("complete", curveList.Count.ToString());
            return Result.Succeeded;
        }

        private WallType GetWallTypeByName (Document doc, string wallTypeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(WallType));

            foreach(Element curElem in collector)
            {
                WallType wallType = curElem as WallType;

                if (wallType.Name == wallTypeName)
                    return wallType;
            }
            return null;
        }

        private Level GetLevelByName (Document doc, string levelName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Level));

            foreach (Element curElem in collector)
            {
                Level level = curElem as Level;

                if (level.Name == levelName)
                    return level;
            }
            return null;
        }

        private MEPSystemType GetSystemTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(MEPSystemType));

            foreach (Element curElem in collector)
            {
                MEPSystemType curType = curElem as MEPSystemType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }

        private PipeType GetPipeTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(PipeType));

            foreach (Element curElem in collector)
            {
                PipeType curType = curElem as PipeType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }

        private DuctType GetDuctTypeByName(Document doc, string typeName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(DuctType));

            foreach (Element curElem in collector)
            {
                DuctType curType = curElem as DuctType;

                if (curType.Name == typeName)
                    return curType;
            }
            return null;
        }
    }
}
