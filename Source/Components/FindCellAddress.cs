using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Parameters;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDriveGH.Components
{

    // This component looks up a specified cell or cells in a google spreadsheet.
    // This makes it easy for the user to refer to a field name in the sheet rather
    // than manually figuring out the proper address in which to write data.
    // User can specify a cell offset one over or one down from the specified cell if desired.
    public class FindCellAddress : GH_Component
    {
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
        public FindCellAddress()
            : base("Find Cell Address", "FindCell", "Finds the address of a cell containing the specified contents", Properties.Resources.AssemblyName, "Spreadsheets") 
        { 

        }

        //set up component inputs - the token, name, worksheets, data to searh for, and cell offset. 
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddParameter(new AuthToken_Param(), "Token", "T", "Google Authentication Token, obtained with Google Authenticator component.", GH_ParamAccess.item);
            //above is custom AuthToken_Param type
            pManager.AddTextParameter("Name", "N", "Name of the spreadsheet to query", GH_ParamAccess.item);
            pManager.AddTextParameter("Worksheet", "W", "Optional Worksheet name", GH_ParamAccess.item);
            pManager[2].Optional = true; //worksheets are optional
            pManager.AddTextParameter("Data", "D", "Data to search for in spreadsheet", GH_ParamAccess.list);

            // This is a quick and dirty way to allow the user to specify offsets by name instead of the corresponding integer value,
            // even though the parameter is just taking in an integer. 
            pManager.AddIntegerParameter("Offset", "O", "An optional cell offset for address output.", GH_ParamAccess.item,0);
            Param_Integer offset = (Param_Integer)pManager[4];
            offset.AddNamedValue("No Offset", 0);
            offset.AddNamedValue("+1 Row", 1);
            offset.AddNamedValue("+1 Column", 2);


        }

        //set up output parameter - the located cell address
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Cells", "C", "First located cell containing specified data.", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            //declare input variables
            AuthToken authToken = new AuthToken();
            string sheetName = "";
            string worksheet = "";
            int offset = 0;
            List<string> data = new List<string>();
 

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
                return; //exit if no name given
            }
            DA.GetData<string>("Worksheet", ref worksheet);

            if (!DA.GetDataList("Data", data))
            {
                return; //exit if no data to search for specified
            }

            DA.GetData("Offset", ref offset);

            //declare output variables
            List<string> cellList = new List<string>(data.Count);


            //set up authentication
            OAuth2Parameters parameters = GDriveUtil.GetParameters(authToken);
            SpreadsheetsService service = GDriveUtil.GetSpreadsheetsService(parameters);
            //find the spreadsheet
            SpreadsheetEntry spreadsheet = GDriveUtil.findSpreadsheetByName(sheetName,service);
            if (spreadsheet == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Specified Spreadsheet not found.");
                return;
            }
            //find the worksheet by name
            WorksheetEntry worksheetEntry = GDriveUtil.findWorksheetByName(worksheet, spreadsheet);
            if (worksheetEntry == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Specified worksheet not found.");
                return;
            }


            // Fetch the cell feed of the worksheet.
            CellQuery cellQuery = new CellQuery(worksheetEntry.CellFeedLink);
            CellFeed cellFeed = service.Query(cellQuery);

            //for each field the user is seeking:
           for(int i=0;i<data.Count;i++){
               bool found = false;
                foreach (CellEntry cell in cellFeed.Entries) // for each cell
                {
                    if(cell.Value.ToUpperInvariant().Equals(data[i].ToUpperInvariant())){ //case insensitive match of cell data
                        CellEntry thisCell = cell;
                        switch (offset) //user specified cell offset, by 0, +1 row, or +1 column
                        {
                            case 0: //no offset
                                cellList.Add(cell.Title.Text);
                                break;
                            case 1: //+1 row offset
                                cellList.Add(GDriveUtil.addressFromCells((int)cell.Column,(int)cell.Row+1));
                                break;
                            case 2: //+1 column offset
                                cellList.Add(GDriveUtil.addressFromCells((int)cell.Column+1, (int)cell.Row));
                                break;
                        }

                        found = true;
                        break; //exit once first matching cell is found
                    }
                    
                }
                if (!found)
                {
                    // alert user cell not found, but insert a null into the output list so that list output corresponds to inputs. 
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, String.Format("{0} not found in sheet",data[i].ToString()));
                    cellList.Add(null);
                }
            }

            DA.SetDataList("Cells", cellList); //pass cell addresses to output.
        }

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.findCellAddress;
            }
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{B91CEB13-B7C3-4F2F-81A9-0D8D406BB849}"); }
        }
    }
}
