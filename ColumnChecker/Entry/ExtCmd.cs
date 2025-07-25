using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnChecker.Data;
using ColumnChecker.Etabs;
using ColumnChecker.Revit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColumnChecker.Entry
{
    [Transaction(TransactionMode.Manual)]
    public class ExtCmd : IExternalCommand
    {
        public static UIDocument UIDoc { get; set; }
        public static Document Doc { get; set; }
        //public static ExtEventHan ExtEventHan { get; set; }

        //public static ExternalEvent ExtEvent { get; set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDoc = commandData.Application.ActiveUIDocument;
            Doc = UIDoc.Document;

            
            ManageEtabs.LinkEtabsModel();
            var data = ManageEtabs.GetDataFromEtabs();

            if(data.columns.Count > 0)
            {
                ManageRevit.CheckDataWithRevit(data.columns, Doc,data.columnArrayGroup);
            }
            


            return Result.Succeeded;
        }
    }
}
