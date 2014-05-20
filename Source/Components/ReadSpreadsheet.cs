using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Grasshopper.Kernel.Data;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDriveGH.Components
{
    // This component looks up a specified spreadsheet and returns its data and corresponding addresses formatted as data trees. 
    
    public class ReadSpreadsheet : GH_Component
    {
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
         public ReadSpreadsheet()
            : base("Read From Spreadsheet", "ReadSpreadsheet", "Reads data from specified spreadsheet", Properties.Resources.AssemblyName, "Spreadsheets")
        {

        }

        //set up input parameters - auth token, name, worksheets, include blank cells, spreadsheet range, row/column order, formulas/values, and a refresh toggle. 
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddParameter(new AuthToken_Param(), "Token", "T", "Google Authentication Token, obtained with Google Authenticator component.", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "Name of the spreadsheet to query", GH_ParamAccess.item);
            pManager.AddTextParameter("Worksheet", "W", "Optional Worksheet name", GH_ParamAccess.item);
            pManager[2].Optional = true; //worksheets optional
            pManager.AddBooleanParameter("Include Blank Cells?", "B", "Set to true to include blank cells in data output.", GH_ParamAccess.item,false);
            pManager.AddParameter(new Range_Param(), "Spreadsheet Range", "SR", "Range of cells to query.", GH_ParamAccess.item);
            pManager[4].Optional = true; //spreadsheet range optional
            pManager.AddBooleanParameter("Organize by Rows or Columns", "R/C", "Set to true to organize data output by row - otherwise data is structured by column.", GH_ParamAccess.item,false);
            pManager.AddBooleanParameter("Read Formulas or Values", "F/V", "Set to true to return formulas rather than values from the spreadsheet", GH_ParamAccess.item,false);
            pManager.AddBooleanParameter("Refresh", "R", "Send a new value to this parameter to cause the spreadsheet data to refresh.", GH_ParamAccess.tree, true);

        }


        // set up output parameters - cell values, addresses, and some metadata about the sheet.
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Values", "V", "The data retrieved from the spreadsheet", GH_ParamAccess.tree);
            pManager.AddTextParameter("Cell Addresses", "CA", "The cell addresses retrieved from the spreadsheet", GH_ParamAccess.tree);
            pManager.AddTextParameter("Sheet Info", "I", "Metadata about the selected spreadsheet", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            
            
            //declare input variables to load into
            
            AuthToken authToken = new AuthToken();
            string sheetName = "";
            string worksheet = "";
            bool includeBlankCells = false;
            bool rangeSpecified = false;
            SpreadsheetRange range = new SpreadsheetRange();
            bool rowsColumns = false;
            bool formulasValues = false;

            //declare output variables
            List<string> metaData = new List<string>();
            GH_Structure<GH_String> values = new GH_Structure<GH_String>();
            GH_Structure<GH_String> addresses = new GH_Structure<GH_String>();
           
            //get data from inputs
            if (!DA.GetData<AuthToken>("Token", ref authToken))
            {
                return; //exit if no token
            }
            if (!authToken.IsValid)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The provided authentication token appears to be invalid. Try re-running the Google Authenticator component.");
                return; //exit if invalid token
            }


            if (!DA.GetData<string>("Name", ref sheetName))
            {
                return; //exit if no name provided
            }
            DA.GetData<string>("Worksheet", ref worksheet);
            DA.GetData<bool>("Include Blank Cells?", ref includeBlankCells);
            if (DA.GetData<SpreadsheetRange>("Spreadsheet Range", ref range))
            {
                rangeSpecified = true;
            }
            DA.GetData<bool>("Organize by Rows or Columns", ref rowsColumns);
            DA.GetData<bool>("Read Formulas or Values", ref formulasValues);

            



            //set up oAuth parameters
            OAuth2Parameters parameters = GDriveUtil.GetParameters(authToken);
            //establish spreadsheetservice
            SpreadsheetsService service = GDriveUtil.GetSpreadsheetsService(parameters);
            //get spreadsheet by name
            SpreadsheetEntry spreadsheet = GDriveUtil.findSpreadsheetByName(sheetName, service);

            if (spreadsheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Specified Spreadsheet not found.");
                return;
            }

            //gather spreadsheet metadata
            metaData.Add("Spreadsheet Name: " + spreadsheet.Title.Text);
            metaData.Add("Last Updated: " + spreadsheet.Updated.ToString());
            metaData.Add("Worksheets: " + worksheetList(spreadsheet.Worksheets));

            //find the specified worksheet, or first one if none specified
            WorksheetEntry worksheetEntry = null;
            worksheetEntry = GDriveUtil.findWorksheetByName(worksheet, spreadsheet);
            
            if (worksheetEntry == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Specified worksheet not found.");
                return;
            }
            

            // Fetch the cell feed of the worksheet.
            CellQuery cellQuery = new CellQuery(worksheetEntry.CellFeedLink);
            if (rangeSpecified)
            {
                if (range.TestValid())
                {
                    cellQuery.MinimumColumn = range.startColumn();
                    cellQuery.MinimumRow = range.startRow();
                    cellQuery.MaximumColumn = range.endColumn();
                    cellQuery.MaximumRow = range.endRow();
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Invalid Range Specified");
                    return;
                }
            }
           
            //passes null values if user wants the blank cells represented, otherwise they are omitted from output. 
            if (includeBlankCells)
            {
                cellQuery.ReturnEmpty = ReturnEmptyCells.yes;
            }
            //set up cell feed
            CellFeed cellFeed = service.Query(cellQuery);
          
            foreach (CellEntry cell in cellFeed.Entries) //for all the cells in the feed
            {
               
                GH_Path path = new GH_Path(DA.Iteration); //set up data path for data tree
                uint e = (rowsColumns) ? cell.Row : cell.Column; //decide whether to structure data path by row or column
                path = path.AppendElement((int)e);

                string v = (formulasValues) ? cell.InputValue : cell.Value; //decide whether to get the cell formula or the cell value as output
                values.Append(new GH_String(v), path); //save the value of the cell 

                addresses.Append(new GH_String(cell.Title.Text), path); // save the address of the cell 
            }

            //set output data
            DA.SetDataTree(0, values);
            DA.SetDataTree(1, addresses);
            DA.SetDataList("Sheet Info", metaData);
            


        }


        //method for formatting list of multiple worksheets as single string
        private string worksheetList(WorksheetFeed feed)
        {
            List<string> list = new List<string>();

            
            foreach (WorksheetEntry entry in feed.Entries)
            {
                list.Add(entry.Title.Text);
            }

            return String.Join(", ", list);
        }
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.readSpreadsheet;
            }
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{F002BC57-9357-497F-BF02-A1B97BFE9EBA}"); }
        }
    }
}
