#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace RevitAddin1
{
    [Transaction(TransactionMode.Manual)]
    public class Command01 : IExternalCommand
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

            
            string text = "Revit add-in academy";
            string fileName = doc.PathName;

            double offset = 0.05;
            double offsetCalc = offset * doc.ActiveView.Scale;
            
            XYZ curPoint = new XYZ(0,0,0);
            XYZ offsetPoint = new XYZ(0,offsetCalc,0);

           
         
     
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create text note");
            t.Start();

            int range = 100;
            for (int i = 1; i <= range; i++)
            {
                TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, "This is Line" + i.ToString(), collector.FirstElementId());
                curPoint = curPoint.Subtract(offsetPoint);
            }


            TextNote myTextNote = TextNote.Create(doc,doc.ActiveView.Id, curPoint,"this is my text note",collector.FirstElementId());

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }

        internal double Method01(double a, double b)
        {
            double c = a + b;
            Debug.Print("Got here" + c.ToString());
            return c;
        }
    }
}
