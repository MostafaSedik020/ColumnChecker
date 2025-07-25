using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnChecker.Data;
using ColumnChecker.Utils;
using System.IO;
using System.Windows;

namespace ColumnChecker.Revit
{
    public static class ManageRevit
    {
        public static void CheckDataWithRevit(List<Column> etabsColumns, Document doc,ColumnArrayGroup columnArrayGroup)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            var columns = collector.OfCategory(BuiltInCategory.OST_StructuralColumns)
                                 .WhereElementIsNotElementType()
                                 .ToElements()
                                 .ToList();
            //collect errors
            StringBuilder duplicates = new StringBuilder(); //for dublicates
            StringBuilder diffDimensions = new StringBuilder();// wrong section dimensions or shape
            StringBuilder diffRebar = new StringBuilder(); //different rebar diameters or number of bars
            StringBuilder undefined = new StringBuilder(); //undefined columns in ETABS such ass irregular shape
            StringBuilder missingColumns = new StringBuilder(); //columns found in etabs but missing in revit

            foreach (var column in etabsColumns)
            {
                var matchingColumns = columns.Where(c=> c.LookupParameter("ETABS Unique Name")?.AsString() == column.UniqueName)
                    .ToList();
                if(column.UniqueName == "57")
                {
                    MessageBox.Show("hi");
                }
                //check if the unique name is duplicated
                if (matchingColumns.Count > 1)
                {
                    duplicates.AppendLine($"Column: {column.UniqueName} has duplicate");
                    continue;
                }

                //check if the column exists in revit
                if (matchingColumns.Count == 0)
                {
                    missingColumns.AppendLine($"Column: {column.UniqueName} is missing in Revit");
                    continue;
                }

                foreach (var ele in matchingColumns)
                {

                    ElementId typeId = ele.GetTypeId();
                    ElementType eleType = doc.GetElement(typeId) as ElementType;

                    //check dimensions and shape
                    if (column.IsRectangle)
                    {
                        string checkShapestr = ele.LookupParameter("Family").AsValueString();
                        bool checkShape = ele.LookupParameter("Family").AsValueString().Contains("_RECTANGULAR_T");

                        if (checkShape)
                        {
                            double revitColWidth = UnitConverter.convertUnitsToMeters( eleType.LookupParameter("b").AsDouble())*1000;
                            double revitColLength = UnitConverter.convertUnitsToMeters(eleType.LookupParameter("h").AsDouble())*1000;

                            if(Math.Abs( revitColWidth - column.Width) > 0.01)
                            {
                                diffDimensions.AppendLine($"Column: {column.UniqueName} has different width in Revit, expected {column.Width} m, found {revitColWidth} m");
                            }

                            if (Math.Abs(revitColLength - column.Length) > 0.01)
                            {
                                diffDimensions.AppendLine($"Column: {column.UniqueName} has different length in Revit, expected {column.Length} m, found {revitColLength} m");
                            }


                        }
                        else
                        {
                            diffDimensions.AppendLine($"Column: {column.UniqueName} has different section shape in Revit, expected rectangular");
                        }


                    }
                    else if (column.IsCircular)
                    {
                        bool checkShape = ele.LookupParameter("Family").AsValueString().Contains("_CIRCULAR_T");
                        if (checkShape)
                        {
                            double revitColDiameter = UnitConverter.convertUnitsToMeters(eleType.LookupParameter("b").AsDouble()) * 1000;

                            if (Math.Abs(revitColDiameter - column.Length) > 0.01) //for circular columns width and length are equal
                            {
                                diffDimensions.AppendLine($"Column: {column.UniqueName} has different diameter in Revit, expected {column.Width} m, found {revitColDiameter} m");
                            }
                        }
                        else
                        {
                            diffDimensions.AppendLine($"Column: {column.UniqueName} has different section shape in Revit, expected circular");
                        }

                    }
                    else
                    {
                        undefined.AppendLine($"columns with unique name {column.UniqueName} has irregular shape,cannot be checked by software");
                    }

                    //check rebar diameter and number of bars
                    int revitRebarDia = ele.LookupParameter("Rebar: Diameter").AsInteger();
                    int revitBarsNumber = ele.LookupParameter("Rebar: No.of bars").AsInteger();

                    if (revitBarsNumber != column.BarsNumber)
                    {
                        diffRebar.AppendLine($"Column: {column.UniqueName} has different number of bars in Revit, expected {column.BarsNumber}, found {revitBarsNumber}");
                    }
                    if (revitRebarDia != column.RebarDia)
                    {
                        diffRebar.AppendLine($"Column: {column.UniqueName} has different rebar diameter in Revit, expected {column.RebarDia} mm, found {revitRebarDia} mm");
                    }
                    
                }
            }
            //write errors to file
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ColumnCheckerErrors.txt");
            if (File.Exists(filePath)) {
                File.Delete(filePath); //delete old file if exists
            }
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                if (duplicates.Length == 0 && diffDimensions.Length == 0 && diffRebar.Length == 0 && undefined.Length == 0 && missingColumns.Length == 0)
                {
                    writer.WriteLine("No issues found.");
                }
                else
                {
                    if(missingColumns.Length > 0)
                    {
                        writer.WriteLine("Missing Columns in Revit:");
                        writer.WriteLine(missingColumns.ToString());
                    }
                    if (duplicates.Length > 0)
                    {
                        writer.WriteLine("Duplicates:");
                        writer.WriteLine(duplicates.ToString());
                    }
                    if (diffDimensions.Length > 0)
                    {
                        writer.WriteLine("Different Dimensions or Shapes:");
                        writer.WriteLine(diffDimensions.ToString());
                    }
                    if (diffRebar.Length > 0)
                    {
                        writer.WriteLine("Different Rebar Diameters or Number of Bars:");
                        writer.WriteLine(diffRebar.ToString());
                    }
                    if (undefined.Length > 0)
                    {
                        writer.WriteLine("Undefined Columns:");
                        writer.WriteLine(undefined.ToString());
                    }
                }
                
                

            }
            
            //open file in default text editor
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });


        }
    }
}
