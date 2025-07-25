using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using Autodesk.Revit.DB;
using ColumnChecker.Utils;

namespace ColumnChecker.Data
{
    public  class ColumnArrayGroup
    {
        public  List<ColumnArray> ColumnsArrays { get; set; } 

        public ColumnArrayGroup()
        {
            // Initialize the ColumnsArrays list if needed
            ColumnsArrays = new List<ColumnArray>();
        }
        
    }
}
