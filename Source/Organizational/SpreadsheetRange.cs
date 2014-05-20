using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Grasshopper.Kernel.Types;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDriveGH
{
    //custom data type to represent an A1-notation cell range in the form "A1 to G5" or "A1:G5"
   public class SpreadsheetRange : GH_Goo<string> //custom data types must extend GH_Goo and also have corresponding GH_Param type, in this case Range_Param. 
    {
       public string StartCell { get; set; }
       public string EndCell { get; set; }


       public uint startColumn()
       {
           return GDriveUtil.colFromAddress(StartCell);
       }

       public uint startRow()
       {
           return GDriveUtil.rowFromAddress(StartCell);
       }

       public uint endColumn()
       {
           return GDriveUtil.colFromAddress(EndCell);
       }

       public uint endRow()
       {
           return GDriveUtil.rowFromAddress(EndCell);
       }


       

        public SpreadsheetRange()
        {
            
        }

        public SpreadsheetRange(SpreadsheetRange range)
        {
            StartCell = range.StartCell;
            EndCell = range.EndCell;
        }

        public SpreadsheetRange(string start, string end)
        {
            StartCell = start;
            EndCell = end;
        }

       

        public SpreadsheetRange(string rangeString)
        {
            SetFromString(rangeString);

        }

        public override IGH_Goo Duplicate()
        {
            return new SpreadsheetRange(this);
        }

        public  bool TestValid() //test if all cell inputs are present and valid. 
        {
            return (!String.IsNullOrEmpty(StartCell)) && (!String.IsNullOrEmpty(EndCell) && GDriveUtil.isValidAddress(StartCell) && GDriveUtil.isValidAddress(EndCell)); 
        }

        public override bool IsValid //test if all cell inputs are present and valid. 
        {
            get { return (!String.IsNullOrEmpty(StartCell)) && (!String.IsNullOrEmpty(EndCell) && GDriveUtil.isValidAddress(StartCell) && GDriveUtil.isValidAddress(EndCell)); }
        }

       //string representation user sees of spreadsheet range when hovering over a spreadsheet range input parameter 
        public override string ToString()
        {
            if (IsValid)
            {
                return String.Format("{0} to {1}", StartCell, EndCell);
            }
            else
            {
                return "Invalid Cell Range";
            }
        }

        public override string TypeDescription
        {
            get { return "This represents a range of cells in a google spreadsheet."; }
        }

        public override string TypeName
        {
            get { return "Spreadsheet Range"; }
        }

       //this method handles parsing a spreadsheet range from simple text input. 
        private bool SetFromString(string inputValue)
        {
            string ToConvert = (inputValue).ToUpperInvariant().Replace(" ", "");
            bool toMatch = LikeOperator.LikeString(ToConvert, "*TO*", CompareMethod.Binary);
            bool colonMatch = LikeOperator.LikeString(ToConvert, "*:*", CompareMethod.Binary);
            if (toMatch)
            {
                string[] splitResults = Regex.Split(ToConvert, "TO");
                StartCell = splitResults[0];
                EndCell = splitResults[1];
            }
            else if (colonMatch)
            {
                string[] splitResults = Regex.Split(ToConvert, ":");
                StartCell = splitResults[0];
                EndCell = splitResults[1];
            }

            return toMatch||colonMatch;

        }

       //enables a user to supply string input in GH and have it convert to SpreadsheetRange data.
        public override bool CastFrom(object source)
        {
            var objectValue = System.Runtime.CompilerServices.RuntimeHelpers.GetObjectValue(source);
            Type type = objectValue.GetType();
            if (objectValue is GH_String)
            {
                string objString = (objectValue as GH_String).ToString();
                return SetFromString(objString);
            }
            else
            {
                return base.CastFrom(source);
            }
        }
    }
}
