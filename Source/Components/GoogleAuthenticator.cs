using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Google.GData.Client;
using Google.GData.Spreadsheets;

namespace GoogleDriveGH.Components
{
    // This component handles authenticating the user with Google.
    // On run, user is presented with dialog box that sends user to
    // webpage to get authentication token, which is then pasted back
    // in GH prompt. Authentication is stored in a custom settings file
    // that lives in the user's Grasshopper folder. (Typically
    // C:/Users/USERNAME/AppData/Roaming/Grasshopper.) Data remains until
    // manually cleared by the user, or replaced with new authentication 
    // data, which may be necessary occasionally if the auth token has
    // expired.


    public class GoogleAuthenticator : GH_Component
    {
        private string PREFS_TOKEN = "Token"; //The string identifying the token in the prefs file
        private string PREFS_REFTOKEN = "Refresh Token"; //the string identifying the refresh token in the prefs file
        private string PREFS_FILENAME = "GAuthToken"; //file name for the prefs xml file

        private bool loggedIn;

        private string accessCode; // the access code used in the OAuth handshake
        private string accessToken; //the google auth token
        private string refreshToken; //the refresh token so the user doesn't need to reauth every time

        private GH_SettingsServer server; //GH_SettingsServer handles writing and reading from the preferences xml file.

        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
         public GoogleAuthenticator() 
            : base("Google Authenticator", "GAuth", "Use this component to authenticate with Google to gain access to your spreadsheets", Properties.Resources.AssemblyName, "Spreadsheets")
        {
             //on component creation, set up the settings server and initialize variables
             server = new GH_SettingsServer(PREFS_FILENAME);
            
             accessCode = "";
             accessToken = server.GetValue(PREFS_TOKEN, "");
             refreshToken = server.GetValue(PREFS_REFTOKEN, "");
             loggedIn = !String.IsNullOrEmpty(accessToken) && !String.IsNullOrEmpty(refreshToken);
             UpdateMenu();
        }
        //set up component inputs - booleans for authenticate and clear
        protected override void RegisterInputParams(GH_Component.GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Authenticate", "A", "Set to true to launch the authentication dialog", GH_ParamAccess.item,false);
            pManager.AddBooleanParameter("Clear Credentials", "C", "Set to true to clear stored authentication credentials and log out.", GH_ParamAccess.item, false);
        }

        //set up component output - uses a custom param of type AuthToken_Param to contain the token. 
        protected override void RegisterOutputParams(GH_Component.GH_OutputParamManager pManager)
        {
            pManager.RegisterParam(new AuthToken_Param(),"Token", "T", "The auth token to use for all other spreadsheet components");
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            //set up input variables
            bool run = false;
            bool clear = false;
            
            //load data from input parameters
            DA.GetData<bool>("Clear Credentials", ref clear);
            DA.GetData<bool>("Authenticate", ref run);

            if (clear)
            {
                //clear the stored credentials by writing empty strings to prefs file
                accessCode = "";
                accessToken = "";
                SaveAuth(server, "", "");
                loggedIn = false;
                UpdateMenu(); //update status menu at bottom of component
                return;
            }

            if (run)
            {
                accessToken = performOAuth(ref refreshToken); //perform the authorization procedure
                //save the received credentials to the settings XML file
                SaveAuth(server, accessToken, refreshToken);
                loggedIn = true;
                UpdateMenu(); //update status menu at bottom of component
            }

            DA.SetData("Token", new AuthToken(accessToken,refreshToken));
            //write auth token to parameter output
        }


        //this method handles displaying a window to the user and performing the OAuth 2.0 handshake.
        private string performOAuth(ref string refreshToken)
        {
            
            ////////////////////////////////////////////////////////////////////////////
            // Set up the OAuth 2.0 object
            ////////////////////////////////////////////////////////////////////////////

            // OAuth2Parameters holds all the parameters related to OAuth 2.0.
            OAuth2Parameters parameters = GDriveUtil.GetParameters();

            // Get the authorization url.  The user of your application must visit
            // this url in order to authorize with Google.  If you are building a
            // browser-based application, you can redirect the user to the authorization
            // url.
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
        

            AuthForm authForm = new AuthForm(); //form displaying prompt + field for user to paste access code
            //TODO: figure out how to avoid the infinite scroll at edge of window problem when dialog box is created
            Grasshopper.GUI.GH_WindowsFormUtil.CenterFormOnCursor(authForm, true);
            authForm.ShowURL(authorizationUrl, Grasshopper.Instances.DocumentEditor);
            accessCode = authForm.token; //retrieve user access code
            parameters.AccessCode = accessCode; 
            //generate token from access code through OAuth handshake
            OAuthUtil.GetAccessToken(parameters);
            accessToken = parameters.AccessToken;
            refreshToken = parameters.RefreshToken;
            return accessToken;
        }
        //set component icon
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                return Properties.Resources.auth;
            }
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{0E052CA0-1EAE-4277-8658-845331BA9DC1}"); }
        }

        //this method saves the token information to the settings XML file. 
        private void SaveAuth(GH_SettingsServer server, string token, string reftoken)
        {
            server.SetValue(PREFS_TOKEN, token);
            server.SetValue(PREFS_REFTOKEN, reftoken);
            server.WritePersistentSettings();
        }

        //set the Message (black text tab at bottom of component) to indicate whether or not the user is currently logged into google.
        private void UpdateMenu()
        {
            if (loggedIn)
            {
                Message = "Logged In";
            }
            else
            {
                Message = "Not Logged In";
            }
        }
    }
}
