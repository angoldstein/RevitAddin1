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
    public class FizzBuzz : IExternalCommand
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

                        
            double offset = 0.05;
            double offsetCalc = offset * doc.ActiveView.Scale;
            
            XYZ curPoint = new XYZ(0,0,0);
            XYZ offsetPoint = new XYZ(0,offsetCalc,0);
                      
         
     
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create text note");
            t.Start();

            int range1 = 100;
            for (int i = 1; i <= range1; i++)
            {
                string result1 = ""; 

                if ((i % 3 == 0) && (i % 5 == 0))
                {
                    result1 = "FizzBuzz";
                }
                else if (i % 3 == 0)
                {
                    result1 = "Fizz";
                }
                else if (i % 5 == 0)
                {
                    result1 = "Buzz";
                }
                else
                {
                    result1 = i.ToString();
                }

                {
                    TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, result1, collector.FirstElementId());
                    curPoint = curPoint.Subtract(offsetPoint);
                }
            }


            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
    }
}
