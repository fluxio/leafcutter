using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using System.Windows.Forms;
using GH_IO.Serialization;

using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDriveGH.Components
{
    public class WriteToSpreadsheet : GH_Component
    {
        // this internal boolean is controlled via user menu item - if enabled a spreadsheet with the specified name is automatically created if not already present. 
        private bool m_createNewSpreadsheets = false;
        public bool createNewSpreadsheets { 
            get
            {
               return m_createNewSpreadsheets; 
            }
            set
            {
                m_createNewSpreadsheets = value;
            }
        }
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
        public WriteToSpreadsheet()
            : base("Write To Spreadsheet", "WriteSpreadsheet", "Writes to Spreadsheet", Properties.Resources.AssemblyName, "Spreadsheets")
        {
            createNewSpreadsheets = false;
        }
        //set up input parameters - token, name, worksheets, row/column, address, data, and write toggle
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddParameter(new AuthToken_Param(), "Token", "T", "Google Authentication Token, obtained with Google Authenticator component.", GH_ParamAccess.item);
            pManager.AddTextParameter("Name", "N", "Name of the spreadsheet to query", GH_ParamAccess.item);
            pManager.AddTextParameter("Worksheet", "W", "Optional Worksheet name", GH_ParamAccess.item);
            pManager[2].Optional = true; //worksheets optional
            pManager.AddBooleanParameter("Organize by Rows or Columns", "R/C", "Set to true to organize data input by row - otherwise data is structured by column.", GH_ParamAccess.item, false);
            pManager.AddTextParameter("Address", "A", "Cell address in A1 notation at which to start writing.", GH_ParamAccess.tree);
            pManager.AddTextParameter("Data", "D", "Data to send to spreadsheet. Lists of data are sent sequentially in row/column order.", GH_ParamAccess.tree);
            pManager.AddBooleanParameter("Write", "Wr", "Set to true to activate spreadsheet writing.", GH_ParamAccess.item,false);

        }

        //set up output parameters - one for the addresses written to
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Address", "A", "Addresses of updated cells", GH_ParamAccess.tree);
        }


        protected override void SolveInstance(IGH_DataAccess DA)
        {

            //declare input variables to load into

            AuthToken authToken = new AuthToken();
            string sheetName = "";
            string worksheet = "";
            bool rowsColumns = false;

            GH_Structure<GH_String> addresses = new GH_Structure<GH_String>();
            GH_Structure<GH_String> data = new GH_Structure<GH_String>();
            bool write = false;

            //declare output variables
            GH_Structure<GH_String> addressesOut = new GH_Structure<GH_String>();

            //get data from inputs
            if (!DA.GetData<AuthToken>("Token", ref authToken))
            {
                return; //exit if no token given
            }
            if (!authToken.IsValid)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The provided authentication token appears to be invalid. Try re-running the Google Authenticator component.");
                return; //exit if token invalid
            }

            if (!DA.GetData<string>("Name", ref sheetName))
            {
                return; //exit if no name given
            }
            DA.GetData<string>("Worksheet", ref worksheet);

            DA.GetData<bool>("Organize by Rows or Columns", ref rowsColumns);
            DA.GetDataTree<GH_String>("Address", out addresses);
            DA.GetDataTree<GH_String>("Data", out data);
            DA.GetData<bool>("Write", ref write);

            if (!write) return; //exit if write is not true

            //check each specified address for validity
            foreach (GH_String address in addresses.AllData(true))
            {
                if (!GDriveUtil.isValidAddress(address.ToString()))
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Error,"Not all addresses are specified in a valid format.");
                    return;
                }
            }

            //setup auth and factory
            //set up oAuth parameters
            OAuth2Parameters parameters = GDriveUtil.GetParameters(authToken);
            SpreadsheetsService service = GDriveUtil.GetSpreadsheetsService(parameters);
        
            //find spreadsheet by name
            SpreadsheetEntry spreadsheet = GDriveUtil.findSpreadsheetByName(sheetName, service);
            if (spreadsheet == null) //if the spreadsheet is not found
            {
                if (createNewSpreadsheets) //if the user has elected to create new spreadsheets
                {
                   List<string> worksheets = new List<string>();
                   if (!String.IsNullOrEmpty(worksheet))
                   {
                       worksheets.Add(worksheet); //create a 1-item list with the worksheet name
                   }
                    spreadsheet = CreateSpreadsheet.createNewSpreadsheet(this, authToken, sheetName, worksheets, false); //create the spreadsheet
                }
                else
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Specified spreadsheet not found.");
                    return;
                }
            }
            //find worksheet by name
            WorksheetEntry worksheetEntry = GDriveUtil.findWorksheetByName(worksheet, spreadsheet);
            if (worksheetEntry == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Specified worksheet not found.");
                return;
            }

            uint maxRow = 0;
            uint maxCol = 0;
            // set up dictionary<Cell address, cell input>
            // associate each input value with a corresponding address
            Dictionary<string, string> writeMap = constructWriteMap(addresses, data, rowsColumns, ref maxRow, ref maxCol, ref addressesOut);
           
            
                //expand worksheet if necessary
                if (worksheetEntry.Rows < maxRow)
                {
                    
                        worksheetEntry.Rows = maxRow;
                        worksheetEntry.Etag = "*"; //required for compatibility with "new" google sheets (see:  http://stackoverflow.com/questions/22719170/updating-cell-in-google-spreadsheets-returns-error-missing-resource-version-id)
                        worksheetEntry.Update();
                    
                }

                if (worksheetEntry.Cols < maxCol)
                {
                    
                        worksheetEntry.Cols = maxCol;
                        worksheetEntry.Etag = "*"; //required for compatibility with "new" google sheets (see:  http://stackoverflow.com/questions/22719170/updating-cell-in-google-spreadsheets-returns-error-missing-resource-version-id)
                        worksheetEntry.Update();
                    
                }

                //set bounds of cell query
                CellQuery cellQuery = new CellQuery(worksheetEntry.CellFeedLink);
                cellQuery.MinimumColumn = 0;
                cellQuery.MinimumRow = 0;
                cellQuery.MaximumColumn = maxCol;
                cellQuery.MaximumRow = maxRow;
                cellQuery.ReturnEmpty = ReturnEmptyCells.yes;

                //retrieve cellfeed
                CellFeed cellFeed = service.Query(cellQuery);

                //convert cell entries to dictionary
               Dictionary<string,AtomEntry> cellDict = cellFeed.Entries.ToDictionary(k => k.Title.Text);

                CellFeed batchRequest = new CellFeed(cellQuery.Uri, service);
                
                //set up batchrequest
                foreach (KeyValuePair<string, string> entry in writeMap)
                {
                    AtomEntry atomEntry;
                    cellDict.TryGetValue(entry.Key, out atomEntry);
                    CellEntry batchEntry = atomEntry as CellEntry;
                    batchEntry.InputValue = entry.Value;
                    batchEntry.Etag = "*"; //required for compatibility with "new" google sheets (see:  http://stackoverflow.com/questions/22719170/updating-cell-in-google-spreadsheets-returns-error-missing-resource-version-id)
                    batchEntry.BatchData = new GDataBatchEntryData(GDataBatchOperationType.update);
                   
                    batchRequest.Entries.Add(batchEntry);
                }
                CellFeed batchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));

                // Check the results
                bool isSuccess = true;
                foreach (CellEntry entry in batchResponse.Entries)
                {
                    string batchId = entry.BatchData.Id;
                    if (entry.BatchData.Status.Code != 200)
                    {
                        isSuccess = false;
                        GDataBatchStatus status = entry.BatchData.Status;
                    }
                }

                if (!isSuccess)
                {
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "Update operations failed");
                }

            
            //output addresses
            DA.SetDataTree(0, addressesOut); //write addressesOut to first output parameter
            
        }

        //this method associates addresses with every piece of data. Responds intelligently to different cases of user input 
        private Dictionary<string, string> constructWriteMap(GH_Structure<GH_String> addresses, GH_Structure<GH_String> data, bool rowColumn, ref uint maxRow, ref uint maxCol, ref GH_Structure<GH_String> addressesOut)
        {
            Dictionary<string, string> writeMap = new Dictionary<string, string>();
            bool incrementBranches = false;
            bool incrementItems = false;


            //case 1: 1 address for entire tree
            //increment branches + increment items = true
            //case 2: 1 address per list
            //increment branches = false + increment items = true;
            //case 3: 1 address per item
            //increment branches + increment items = false

            bool validDataMatch = false;

            if (addresses.DataCount == 1) // case 1
            {
                incrementBranches = true;
                incrementItems = true;
                validDataMatch = true;
            }
            else if(addresses.Branches.Count==data.Branches.Count)
            {
                if (oneItemPerBranch(addresses))//one item in each address branch, case 2
                {
                    incrementItems = true;
                    validDataMatch = true;
                }
                else if(sameBranchLengths(data,addresses)){
                    //otherwise leave both to false in case 3
                    validDataMatch = true;
                } 

            } 

            if (!validDataMatch)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Either provide one address per data item, one address per list, or one address for the whole tree. \n In the latter two cases, sequential addresses are calculated automatically.");
                return writeMap; //return empty write map - nothing happens
            }

            for (int i_branch = 0; i_branch < data.Branches.Count; i_branch++)
            {
                List<GH_String> dataBranch = data.Branches[i_branch];
                List<GH_String> addressBranch = addresses.Branches[Math.Min(i_branch, addresses.Branches.Count - 1)];
                for (int i = 0; i < dataBranch.Count; i++)
                {
                    string dataItem = dataBranch[i].ToString();
                    string addressItem = addressBranch[Math.Min(i, addressBranch.Count - 1)].ToString();

                    int colOffset = 0;
                    int rowOffset = 0;
                    if (rowColumn) // if true, data is structured in rows - otherwise in columns
                    {
                      colOffset = (incrementItems) ? i : 0;
                        rowOffset = (incrementBranches) ? i_branch : 0;
                    }
                    else
                    {
                        colOffset = (incrementBranches) ? i_branch : 0;
                        rowOffset = (incrementItems) ? i : 0;
                    }
                    

                    long colAddress = GDriveUtil.colFromAddress(addressItem) + colOffset;
                    long rowAddress = GDriveUtil.rowFromAddress(addressItem) + rowOffset;

                    if (colAddress > maxCol) maxCol = (uint)colAddress;
                    if (rowAddress > maxRow) maxRow = (uint)rowAddress;

                    addressItem = GDriveUtil.addressFromCells(colAddress, rowAddress);
                    //list of cell addresses that corresponds to input data order
                    addressesOut.Append(new GH_String(addressItem),data.Paths[i_branch]);
                    //dictionary of addresses + data to write - prep for batch write operation.
                    try
                    {
                        writeMap.Add(addressItem, dataItem); //will throw an error if you try to write to the same cell twice - key is already in dictionary.

                    }
                    catch
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "You are trying to write into the same cell twice. Some data is being lost");
                    }
                }

            }


            return writeMap;
        }

       


        //test if all branches have 1 item
        bool oneItemPerBranch(IGH_Structure g)
        {
            bool results = true;
            for (int i = 0; i < g.PathCount;i++ )
            {
                if (g.get_Branch(i).Count != 1) results = false;
            }
            return results;
        }

        //test if all branches have exactly the same length. Assumes A and B have the same number of branches
        bool sameBranchLengths(IGH_Structure A, IGH_Structure B)
        {
            bool results = true;
            for (int i = 0; i < A.PathCount; i++)
            {
                if (A.get_Branch(i).Count != B.get_Branch(i).Count) results = false;
            }
            return results;
        }

        //this method adds the "create spreadsheet" menu item to the component menu
        protected override void AppendAdditionalComponentMenuItems(System.Windows.Forms.ToolStripDropDown menu)
        {
            ToolStripMenuItem toolStripMenuItem = GH_DocumentObject.Menu_AppendItem(menu, "Create spreadsheet if it doesn't exist", new EventHandler(menu_enableNewSpreadsheet), true, createNewSpreadsheets);
            toolStripMenuItem.ToolTipText = "When checked, the component will create a new spreadsheet if the specified one does not exist.";
        }

        //this manages the event when the user enables or disables the "create spreadsheet" menu item. 
        private void menu_enableNewSpreadsheet(object sender, EventArgs e)
        {
            RecordUndoEvent("Create New Spreadsheet");
            createNewSpreadsheets = !createNewSpreadsheets;
            ExpireSolution(true);
        }

        //write and read allow the settings to persist through document save and component copy/paste.
        public override bool Write(GH_IWriter writer)
        {
            writer.SetBoolean("createNewSpreadsheets", createNewSpreadsheets);
            return base.Write(writer);
        }
        public override bool Read(GH_IReader reader)
        {
            reader.TryGetBoolean("createNewSpreadsheets", ref m_createNewSpreadsheets);
           
            return base.Read(reader);
        }

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.writeSpreadsheet;
            }
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{97694195-3A36-4F70-83E1-7BD397523578}"); }
        }
    }
}
