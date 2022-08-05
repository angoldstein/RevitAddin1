using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace RevitAddin1
{
    public class FurnSet
    {
        public string setType { get; set; }
        public string setName { get; set; }
        public List<string> furnList { get; private set; }

        public FurnSet(string setType, string setName, string _furnList)
        {
            this.setType = setType;
            this.setName = setName;
            furnList = GetFurnListFromString(_furnList);
        }

        private List<string> GetFurnListFromString(string list)
        {
            List<string> returnList = list.Split(',').ToList();
            List<string> returnList2 = new List<string>();

            foreach (string str in returnList)
                returnList2.Add(str.Trim());


            return returnList;
        }

        public int FurnitureCount()
        {
            return furnList.Count;
        }
    }

    public class FurnData
    {
        public string furnName { get; set; }
        public string familyName { get; set; }
        public string typeName { get; set; }
        public FamilySymbol familySymbol { get; private set; }
        public Document doc { get; set; }

        public FurnData(Document doc, string furnName, string familyName, string typeName)
        {
            this.furnName = furnName;
            this.familyName = familyName;
            this.typeName = typeName;
            this.doc = doc;
            familySymbol = GetFamilySymbol();
        }
        private FamilySymbol GetFamilySymbol()
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Family));

            foreach (Family curFam in collector)
            {
                
                if (curFam.Name == familyName)
                {
                    ISet<ElementId> famSymbolList = curFam.GetFamilySymbolIds();

                    foreach (ElementId curId in famSymbolList)
                    {
                        FamilySymbol curFS = doc.GetElement(curId) as FamilySymbol;
                        
                        if (curFS.Name == typeName)
                            return curFS;
                    }
                }
            }
            return null;
        }
    }

}
