using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ColumnChecker.Data
{
    public class Column
    {
        public double Width { get; set; }
        public double Length { get; set; }
        public string MarkLabel { get; set; }
        public int MarkNumber { get; set; }
        public string ColumnMark => $"{MarkLabel}{MarkNumber}"; // Concatenates Mark and MarkNumber to form the ColumnMark
        public double ID { get; set; }
        public string BaseLevel { get; set; }
        public string TopLevel { get; set; }
        public double RebarDia { get; set; }
        public int BarsNumber { get; set; }
        public double Volume { get; set; }
        public bool IsRectangle { get; set; } = false;
        public bool IsCircular { get; set; } = false;

        //Etabs prop only
        public string EtabsName { get; set; }
        public string UniqueName { get; set; }
        public double point1X { get; set; }
        public double point1Y { get; set; }
        public double point1Z { get; set; }
        public double point2X { get; set; }
        public double point2Y { get; set; }
        public double point2Z { get; set; }
        public double Angle { get; set; }
        public string EtabsStory { get; set; }
        public Rft Rebar { get; set; }

        public bool IsSimilar(Column other)
        {
            if (other == null)
            {

                return false;
            }


            return Width == other.Width &&
                   Length == other.Length &&
                   BaseLevel == other.BaseLevel &&
                   TopLevel == other.TopLevel &&
                   RebarDia == other.RebarDia &&
                   BarsNumber == other.BarsNumber;
        }




    }
}
