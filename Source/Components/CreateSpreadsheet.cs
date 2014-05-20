using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using Google.GData.Documents;

namespace GoogleDriveGH.Components
{
   //component to create a new spreadsheet. Uses deprecated Documents List API... could probably be rewritten to take advantage of Google Drive API. 

    public class CreateSpreadsheet : GH_Component
    {
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
        public CreateSpreadsheet() 
            : base("Create New Spreadsheet", "CreateSpreadsheet", "Creates a new spreadsheet", Properties.Resources.AssemblyName, "Spreadsheets")
        {

        }
        //register component inputs - Token, Name, Overwrite, Worksheets, and Run.
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddParameter(new AuthToken_Param(), "Token", "T", "Google Authentication Token, obtained with Google Authenticator component.", GH_ParamAccess.item);
         //above parameter is of custom AuthToken_Param type
            pManager.AddTextParameter("Name", "N", "Name of spreadsheet to create", GH_ParamAccess.item);
            pManager.AddBooleanParameter("Overwrite", "O", "If true, spreadsheet with specified name will be overwritten if existing.", GH_ParamAccess.item,false);
            pManager.AddTextParameter("Worksheets", "W", "If worksheet names supplied, multiple worksheets will be created in specified spreadsheet", GH_ParamAccess.list);
            pManager[3].Optional = true; //worksheets is optional
            pManager.AddBooleanParameter("Run", "R", "Set to true to add specified spreadsheet to Google Drive", GH_ParamAccess.item,false);
        }

        //register component output - name of created spreadsheet
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.AddTextParameter("Name", "N", "On success, the name of the created spreadsheet", GH_ParamAccess.item);
           
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            //declare variables for inputs
           
            AuthToken authToken = new AuthToken();
            string sheetName = "";
            bool overwrite = false;
            List<string> worksheets = new List<string>();
             bool run = false;
         

            //get data from input parameters
            DA.GetData<bool>("Run", ref run);
            if (!DA.GetData<AuthToken>("Token", ref authToken))
            {
                return; //exit if no token
            }
            if (!authToken.IsValid)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The provided authentication token appears to be invalid. Try re-running the Google Authenticator component.");
                return; //exit if token isn't valid
            } 


            if (!DA.GetData<string>("Name", ref sheetName))
            {
                return; //exit if no name specified
            }

            DA.GetData<bool>("Overwrite", ref overwrite);
            DA.GetDataList<string>("Worksheets", worksheets);


            if (run)
            {
                //this method does all the work - creates the google spreadsheet.
                createNewSpreadsheet(this, authToken, sheetName, worksheets, overwrite);
            }

            DA.SetData("Name", sheetName); //output the sheet name


        }

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.new_spreadsheet;
            }
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{58E27A3C-E94D-4C22-88E7-D1508BC44EAC}"); }
        }

        // this method creates the spreadsheet, deletes the old one if overwrite is on and there's already on with the same name, 
        // and adds the worksheets if they've been specified. Wrapped in a static method so other components (like WriteToSpreadsheet)
        // can make use of it and create new spreadsheets. 
        public static SpreadsheetEntry createNewSpreadsheet(GH_ActiveObject activeObj, AuthToken authToken, string sheetName, List<string> worksheets, bool overwrite)
        {
            SpreadsheetEntry matchingEntry = null;
            //setup OAuth Parameters
            OAuth2Parameters parameters = GDriveUtil.GetParameters(authToken);

            // It seems clunky to need both a SpreadsheetService and a DocumentService - but
            // DocumentService is necessary to add/delete spreadsheets, and SpreadsheetService 
            // is needed to manipulate the worksheets.


            //setup auth and factory for documentsService

            GOAuth2RequestFactory requestFactory =
          new GOAuth2RequestFactory(null, "MyDocumentsListIntegration-v1", parameters);
            DocumentsService docService = new DocumentsService("MyDocumentsListIntegration-v1");
            docService.RequestFactory = requestFactory;

            //setup SpreadsheetsService
            SpreadsheetsService service = GDriveUtil.GetSpreadsheetsService(parameters);

            //make spreadsheet documentquery
            Google.GData.Documents.SpreadsheetQuery dQuery = new Google.GData.Documents.SpreadsheetQuery();

            DocumentsFeed dFeed = docService.Query(dQuery);


            //if user has opted to overwrite, find first matching spreadsheet and delete. If no matching spreadsheet found, nothing happens.
            if (overwrite)
            {

                foreach (DocumentEntry entry in dFeed.Entries)
                {
                    if (entry.Title.Text.Equals(sheetName))
                    {
                        docService.Delete(entry);
                        break;
                    }
                }
            }

            //create new spreadsheet object 
            DocumentEntry dEntry = new DocumentEntry();
            dEntry.Title.Text = sheetName;
            dEntry.Categories.Add(DocumentEntry.SPREADSHEET_CATEGORY);
            docService.Insert(DocumentsListQuery.documentsBaseUri, dEntry);


            //find the spreadsheet we just created as a SpreadsheetEntry

            matchingEntry = GDriveUtil.findSpreadsheetByName(sheetName, service);
            //if worksheets specified, add worksheets
            if (worksheets.Count > 0)
            {
               
                if (matchingEntry == null) //this shouldn't ever happen, since we just created a new spreadsheet called sheetName.
                {
                    activeObj.AddRuntimeMessage(GH_RuntimeMessageLevel.Error, "Something went wrong with spreadsheet creation.");
                    return null;
                }

                WorksheetFeed wsFeed = matchingEntry.Worksheets;

                //first, we find the existing worksheets, store them, add new ones, and then delete the old.  

                List<WorksheetEntry> entriesToDelete = new List<WorksheetEntry>();
                foreach (WorksheetEntry entry in wsFeed.Entries)
                {
                    entriesToDelete.Add(entry);
                }

                // find the dimensions of the first worksheet, to use for other worksheet creation
                uint rows = ((WorksheetEntry)wsFeed.Entries[0]).Rows;
                uint cols = ((WorksheetEntry)wsFeed.Entries[0]).Cols;

                for (int i = 0; i < worksheets.Count; i++)
                {
                    {
                        string wsName = worksheets[i];

                        WorksheetEntry entry = new WorksheetEntry(rows, cols, wsName);
                        service.Insert(wsFeed, entry);
                    }
                }

                foreach (WorksheetEntry entryToDelete in entriesToDelete)
                {
                    entryToDelete.Delete();
                }

            }

            //end if worksheets...



            return matchingEntry;
            
        }


    }
}
