using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ETABSv1;
using ColumnChecker.Data;
using ColumnChecker.Utils;
using System.Runtime.Serialization;

namespace ColumnChecker.Etabs
{
    public static class ManageEtabs
    {
        private static cSapModel mySapModel = default;
        public static void LinkEtabsModel()
        {
            //set the following flag to true to attach to an existing instance of the program 
            //otherwise a new instance of the program will be started 
            bool AttachToInstance;
            AttachToInstance = true;

            //set the following flag to true to manually specify the path to ETABS.exe
            //this allows for a connection to a version of ETABS other than the latest installation
            //otherwise the latest installed version of ETABS will be launched
            bool SpecifyPath;
            SpecifyPath = false;

            //dimension the ETABS Object as cOAPI type
            cOAPI myETABSObject = null;

            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero) 
            int ret = 0;

            //create API helper object
            cHelper myHelper = new Helper();
            //myETABSObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");


            //attach to a running instance of ETABS 
            try
            {
                //get the active ETABS object

                myETABSObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");
                mySapModel = myETABSObject.SapModel;
            }
            catch (Exception ex)
            {


            }
            //Get a reference to cSapModel to access all API classes and functions
            if (myETABSObject == null)
            {
                MessageBox.Show("ETABS Linked Failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            mySapModel = myETABSObject.SapModel;



            //Check ret value 
            if (ret == 0)
            {
                MessageBox.Show("ETABS Linked successfully.", "Note", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show("ETABS Linked Failed.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }


        }
        public static (List<Column> columns , ColumnArrayGroup columnArrayGroup) GetDataFromEtabs()
        {
            List<Column> columns = new List<Column>();
            ColumnArrayGroup columnArrayGroup = new ColumnArrayGroup();
            int ret = 0;
            if (mySapModel == null)
            {
                
                return (columns,columnArrayGroup);
            }

            //Get the list of all column names
    
            int numberNames = 0;
            string[] allFramesUniName = null;
            string[] allFramePropName = null;
            string[] storyName = null;
            string[] pointName1 = null;
            string[] pointName2 = null;
            double[] point1X = null;
            double[] point1Y = null;
            double[] point1Z = null;
            double[] point2X = null;
            double[] point2Y = null;
            double[] point2Z = null;
            double[] angle = null;
            double[] offset1X = null;
            double[] offset2X = null;
            double[] offset1Y = null;
            double[] offset2Y = null;
            double[] offset1Z = null;
            double[] offset2Z = null;
            int[] cardinalPoint = null;

            //get the section name
            int myType = 0;
            string PropName = null;
            string SAuto = null;

            //get section dimensions
            string FileName = null;
            string MatProp = null;
            double T3 = 0;
            double T2 = 0;
            int Color = 0;
            string Notes = null;
            string GUID = null;

            //get the rebar properties
            string MatPropLong = null;
            string MatPropConfine = null;
            int Pattern = 0;
            int ConfineType = 0;
            double Cover = 0;
            int NumberCBars = 0;
            int NumberR3Bars = 0;
            int NumberR2Bars = 0;
            string RebarSize = null;
            string TieSize = null;
            double TieSpacingLongit = 0;
            int Number2DirTieBars = 0;
            int Number3DirTieBars = 0;
            bool ToBeDesigned = false;

            

            //get all frames properties
            ret = mySapModel.FrameObj.GetAllFrames(
            ref numberNames, ref allFramesUniName, ref allFramePropName, ref storyName, ref pointName1, ref pointName2,
            ref point1X, ref point1Y, ref point1Z, ref point2X, ref point2Y, ref point2Z,
            ref angle, ref offset1X, ref offset2X, ref offset1Y, ref offset2Y, ref offset1Z, ref offset2Z, ref cardinalPoint
            );

            bool IsRectangle = false;
            bool IsCircular = false;

            for (int i = 0; i < numberNames; i++)
            {
                

                ret = mySapModel.PropFrame.GetTypeRebar(allFramePropName[i], ref myType);// check if the type beam or column

                if (myType == 1) //used to filter column elements only
                {
                    ret = mySapModel.PropFrame.GetCircle(allFramePropName[i], ref FileName, ref MatProp, ref T3, ref Color, ref Notes, ref GUID);// get the circlar column

                    if(ret == 0) //if it is not circular then it is rectangular column
                    {
                        IsCircular = true;
                    }
                       
                    ret = mySapModel.PropFrame.GetRectangle(allFramePropName[i], ref FileName, ref MatProp, ref T3, ref T2, ref Color, ref Notes, ref GUID);// get the rec column

                    if (ret == 0) //if it is not circular then it is rectangular column
                    {
                        IsRectangle = true;
                    }
                    //if (ObjectName[i] == "115")
                    //{
                    //    MessageBox.Show("hey");
                    //}
                    Column column = new Column
                    {
                        UniqueName = allFramesUniName[i],
                        IsCircular = IsCircular,
                        IsRectangle = IsRectangle,
                        Width = T2 * 1000,
                        Length = T3 * 1000,
                        EtabsStory = storyName[i],
                        point1X = point1X[i],
                        point1Y = point1Y[i],
                        point1Z = point1Z[i],
                        point2X = point2X[i],
                        point2Y = point2Y[i],
                        point2Z = point2Z[i],
                        Angle = MultiUtils.NormalizeAngleTo90(angle[i]),

                    };

                    columns.Add(column);

                     IsRectangle = false;// reset the flag for next column
                     IsCircular = false;
                }

            }

            foreach (var column in columns)
            {
                ret = mySapModel.FrameObj.GetSection(column.UniqueName, ref PropName, ref SAuto); //return section name e.g:B300X700
                ret = mySapModel.PropFrame.GetRebarColumn(
                    PropName,
                    ref MatPropLong, ref MatPropConfine, ref Pattern, ref ConfineType, ref Cover,
                    ref NumberCBars, ref NumberR3Bars, ref NumberR2Bars, ref RebarSize, ref TieSize,
                    ref TieSpacingLongit, ref Number2DirTieBars, ref Number3DirTieBars, ref ToBeDesigned
                );
                //if(column.EtabsId == "897")
                //{
                //    MessageBox.Show("hey");
                //}
                column.RebarDia = double.Parse(RebarSize);
                if (Pattern == 1)
                {
                    column.BarsNumber = 2*(NumberR2Bars + NumberR3Bars)-4;
                }
                else if(Pattern == 2)
                {
                    column.BarsNumber = NumberCBars;
                }
                NumberR3Bars= 0; // reset the number of bars for next column
                NumberR2Bars = 0; // reset the number of bars for next column
                NumberCBars = 0; // reset the number of bars for next column

            }
            //Create ColumnGroup e,g C1,C2,C3 but here we collect them first and
            //willbe sorted later
            int markNum = 1;
            foreach (Column column in columns)
            {

                //check if this column already has array to belong or not
                var chosenArray = columnArrayGroup.ColumnsArrays.FirstOrDefault(col => col.ColumnList.Any(x => x.point1X == column.point1X &&
                                                                                                                x.point1Y == column.point1Y &&
                                                                                                                x.Angle == column.Angle));

                if (chosenArray != null)
                {
                    columnArrayGroup.ColumnsArrays.Where(Arr => Arr.MarkNumber == chosenArray.MarkNumber).FirstOrDefault().ColumnList.Add(column);
                    columnArrayGroup.ColumnsArrays.Where(Arr => Arr.MarkNumber == chosenArray.MarkNumber).FirstOrDefault().ColumnList
                                                                                                         .Sort((a, b) => a.point2Z.CompareTo(b.point2Z));
                    column.MarkNumber = chosenArray.MarkNumber;
                    //chosenArray.ColumnList.Add(column);
                }
                else
                {
                    //if couldnt find a pairable array 
                    // it create a new one and put the column in this array 
                    // and the new array assigned in the Group
                    ColumnArray columnArray = new ColumnArray();
                    columnArray.ColumnList.Add(column);
                    columnArray.MarkNumber = markNum;
                    column.MarkNumber = markNum;
                    columnArrayGroup.ColumnsArrays.Add(columnArray);
                    markNum++;
                    
                }
            }
            var test = columnArrayGroup;

            return (columns,columnArrayGroup);
        }

    }
}
