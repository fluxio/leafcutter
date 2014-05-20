using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.GUI;

namespace GoogleDriveGH
{
    //Custom Param type so that components can receive/output AuthToken data
    public class AuthToken_Param : GH_Param<AuthToken>
    {
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
        public AuthToken_Param()
            : base(new GH_InstanceDescription("Google Auth Token", "GAT", "Holds a Google Authentication Token", "Params"))
        {
        }
        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{CAEF4281-1205-415F-B0A5-D07D0CA6B7B2}"); }
        }

        //Don't show an authToken param type in the grasshopper toolbar - only use it internal to components. 
        public override GH_Exposure Exposure
        {
            get
            {
                return GH_Exposure.hidden;
            }
        }

        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.GAuthParam;
            }
        }
    }
}
