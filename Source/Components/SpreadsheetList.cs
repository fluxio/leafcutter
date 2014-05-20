using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace GoogleDriveGH.Components
{
    // This component grabs a list of the user's spreadsheets. An optional filter permits the return of a limited set

    public class SpreadsheetList : GH_Component
    {
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
         public SpreadsheetList()
            : base("Spreadsheet List", "SpreadsheetList", "Returns a list of spreadsheets from Google Drive", Properties.Resources.AssemblyName, "Spreadsheets")
        {

        }
        //set up input parameters - token, filter, and refresh trigger
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddParameter(new AuthToken_Param(), "Token", "T", "Google Authentication Token, obtained with Google Authenticator component.", GH_ParamAccess.item);
            //description of filter parameter
            string desc = string.Concat(new string[]
				{
					"Allows you to filter spreadsheets from your drive.",
					"\n",
                    "Filter is case sensitive.",
                    "\n",
					"The following wildcards are allowed:",
					"\n",
					"? (any single character)",
					"\n",
					"* (zero or more characters)",
					"\n",
					"# (any single digit [0-9])",
					"\n",
					"[chars] (any single character in chars)",
					"\n",
					"[!chars] (any single character not in chars)"
				});
            pManager.AddTextParameter("Filter", "F", desc, GH_ParamAccess.item);
            pManager[1].Optional = true; //filter is optional
            pManager.AddBooleanParameter("Refresh", "R", "Send a new value to this parameter to cause the list of spreadsheets to refresh.", GH_ParamAccess.tree, true);
        
            //refresh parameter value is never actually used, but any new data passed into it will trigger a new solveinstance. 
        }

        //set up output parameter - list of spreadsheets
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.Register_StringParam("Spreadsheets", "S", "A list of spreadsheets on your Google Drive account.",GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
       
            //declare input variables
            bool useFilter = false;
            string nameFilter = "";
            AuthToken authToken = new AuthToken();

            //get data from user inputs
            if (DA.GetData<string>("Filter", ref nameFilter)) useFilter = true;
            if(!DA.GetData<AuthToken>("Token", ref authToken))
            {
                return;
            }
            if (!authToken.IsValid)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "The provided authentication token appears to be invalid. Try re-running the Google Authenticator component.");
                return;
            }

            //declare output variables
            List<string> spreadsheetList = new List<string>();

            OAuth2Parameters parameters = GDriveUtil.GetParameters(authToken);
            SpreadsheetsService service = GDriveUtil.GetSpreadsheetsService(parameters);
            SpreadsheetQuery query = new SpreadsheetQuery();

            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = service.Query(query);

            // Iterate through all of the spreadsheets returned 
            foreach(var entry in feed.Entries){
                string name = entry.Title.Text;

                bool addSpreadsheet = (useFilter) ? LikeOperator.LikeString(name, nameFilter, CompareMethod.Binary) : true; //test for matching spreadsheet name if filter has been specified 
                if (addSpreadsheet) spreadsheetList.Add(name);

            }

            

            DA.SetDataList("Spreadsheets", spreadsheetList); //pass the spreadsheet list to the output parameter. 
        }

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.spreadsheetList;
            }
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{3302AEA8-FB15-4BD0-AE15-05482608074D}"); }
        }
    }
}
