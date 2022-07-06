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

            XYZ curPoint = new XYZ(0, 0, 0);
            XYZ offsetPoint = new XYZ(0, offsetCalc, 0);



            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(TextNoteType));

            Transaction t = new Transaction(doc, "Create text note");
            t.Start();

            int range = 100;
            for (int i = 1; i <= range; i++)
            {
                string result1 = CheckFizzBuzz(i);

                CreateTextNote(doc, result1, curPoint, collector.FirstElementId());
                curPoint = curPoint.Subtract(offsetPoint);

            }


            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
        internal void CreateTextNote(Document doc, string text, XYZ curPoint, ElementId id)
        {
            TextNote curNote = TextNote.Create(doc, doc.ActiveView.Id, curPoint, text, id);
        }
        internal string CheckFizzBuzz(int number)
        {
            string result1 = "";

            if ((number % 3 == 0) && (number % 5 == 0))
            {
                result1 = "FizzBuzz";
            }
            else if (number % 3 == 0)
            {
                result1 = "Fizz";
            }
            else if (number % 5 == 0)
            {
                result1 = "Buzz";
            }
            else
            {
                result1 = number.ToString();
            }
            return result1;
        }
    }
}
