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
    public class Command05Bonus : IExternalCommand
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

            string revitFile = "";

            Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            ofd.Title = "Select Revit File";
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Revit Files|*.rvt";

            if (ofd.ShowDialog() != Forms.DialogResult.OK)
                return Result.Failed;

            revitFile = ofd.FileName;

            UIDocument newUIdoc = uiapp.OpenAndActivateDocument(revitFile);
            Document newDoc = newUIdoc.Document;

            FilteredElementCollector collector = new FilteredElementCollector(newDoc);
            collector.OfCategory(BuiltInCategory.OST_IOSModelGroups);
            collector.WhereElementIsNotElementType();

            List<ElementId> groupIDlist = new List<ElementId>();
            foreach (Element curElem in collector)
            {
                groupIDlist.Add(curElem.Id);
            }

            Transform transform = null;
            CopyPasteOptions options = new CopyPasteOptions();

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Load Groups");
                ElementTransformUtils.CopyElements(newDoc, groupIDlist, doc, transform, options);
                t.Commit();
            }

            try
            {
                uiapp.OpenAndActivateDocument(doc.PathName);
                newUIdoc.SaveAndClose();
            }
            catch (Exception)
            { }
            
            TaskDialog.Show("Complete", "Loaded "+groupIDlist.Count.ToString() + " groups into current model.");
                       
            return Result.Succeeded;
        }
    }
}
