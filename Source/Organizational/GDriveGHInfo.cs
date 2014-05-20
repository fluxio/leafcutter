using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;

namespace GoogleDriveGH
{

    //information about the assembly that shows up in the Grasshopper about window. 
    public class GDriveGHInfo: GH_AssemblyInfo
    {
        public GDriveGHInfo()
        {

        }

        public override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.assemblyIcon;
            }
        }

        

        public override string Name
        {
            get
            {
                return "Google Drive for GH ";
            }
        }

        public override string AuthorName
        {
            get
            {
                return "Andrew Heumann for Flux"; //we can word this however you like
            }
        }

        public override string AssemblyDescription
        {
            get
            {
                return "This library enables an interface for GH with Google Spreadsheets."+"\n"+Properties.Resources.license;
            }
        }

        public override string Version
        {
            get
            {
                return "Beta 1";
            }
        }

        public override string AuthorContact
        {
            get
            {
                return "andrew@heumann.com";
            }
        }

    }

    //adds the small icon in the toolbar for the leafcutter tab.
    public class AssemblyPriority : GH_AssemblyPriority
    {
        public AssemblyPriority()
        {

        }
        public override GH_LoadingInstruction PriorityLoad()
        {
            Grasshopper.Instances.ComponentServer.AddCategoryIcon(Properties.Resources.AssemblyName, Properties.Resources.assemblyIcon_16);
            return GH_LoadingInstruction.Proceed;
        }
    }
}
