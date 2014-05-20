using System;
using System.Collections.Generic;

using System.Text;
using Grasshopper.Kernel.Types;

namespace GoogleDriveGH
{
    //Class AuthToken wraps a string so that the token is not visible to users when passed from component to component. 
    public class AuthToken : GH_Goo<string> //custom data types must extend GH_Goo, and must also have a corresponding param type, in this case AuthToken_Param.
    {
        public string Token { get; set; }
        public string RefreshToken { get; set; }
        public AuthToken()
        {
            RefreshToken = "";
            Token = "";
        }
        public AuthToken(string _token,string _refreshToken)
        {
            Token = _token;
            RefreshToken = _refreshToken;
        }

        //text representation of the param which will show when passed to a text panel or when user hovers over a token input/output. 
        public override string ToString()
        {
            if (!IsValid)
            {
                return "Empty Google Authentication Token.";
            }
            else
            {
                return "Google Authentication Token";
            }
        }


        public override IGH_Goo Duplicate()
        {
            return new AuthToken(Token,RefreshToken);
        }

        public override bool IsValid //test for a valid token - invalid tokens will be stored as an empty string, such as when the user clears authentication.
        {
            get { return !String.IsNullOrEmpty(Token); }
        }

        public override string TypeDescription
        {
            get {return "An authentication token for Google account obtained with the Google Authenticator Component"; }
        }

        public override string TypeName
        {
            get { return "Google Auth Token"; }
        }
    }
   
}
