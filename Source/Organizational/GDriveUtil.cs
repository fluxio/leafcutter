using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Client;
using Google.GData.Spreadsheets;
using System.Text.RegularExpressions;

namespace GoogleDriveGH
{
    //utility class with helpful functions used by several components
    class GDriveUtil
    {
        public GDriveUtil()
        {
        }

        //get OAuth parameters
        public static OAuth2Parameters GetParameters()
        {
            ////////////////////////////////////////////////////////////////////////////
            //  Set up the OAuth 2.0 object
            ////////////////////////////////////////////////////////////////////////////

            // OAuth2Parameters holds all the parameters related to OAuth 2.0.
            OAuth2Parameters parameters = new OAuth2Parameters();

            // Set your OAuth 2.0 Client Id (which you can register at
            // https://code.google.com/apis/console).
            parameters.ClientId = Properties.Resources.ClientID; //you must register your own client ID and put it in a text file in resources called ClientID.txt.

            // Set your OAuth 2.0 Client Secret, which can be obtained at
            // https://code.google.com/apis/console.
            parameters.ClientSecret = Properties.Resources.ClientSecret; //you must register your own client secret and put it in a text file in resources called ClientSecret.txt.

            // Set your Redirect URI, which can be registered at
            // https://code.google.com/apis/console.
            parameters.RedirectUri = Properties.Resources.REDIRECT_URI;

            ////////////////////////////////////////////////////////////////////////////
            // STEP 3: Get the Authorization URL
            ////////////////////////////////////////////////////////////////////////////

            // Set the scope for this particular service.
            parameters.Scope = Properties.Resources.SCOPE;
            return parameters;
        }

        //gets the OAuth parameters from a given Auth Token
        public static OAuth2Parameters GetParameters(AuthToken authToken)
        {
            OAuth2Parameters parameters = GetParameters();

            parameters.AccessToken = authToken.Token;
            parameters.RefreshToken = authToken.RefreshToken;

            return parameters;
        }




        public static SpreadsheetsService GetSpreadsheetsService(OAuth2Parameters parameters)
        {
            GOAuth2RequestFactory requestFactory =
     new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
            SpreadsheetsService service = new SpreadsheetsService("MySpreadsheetIntegration-v1");
            service.RequestFactory = requestFactory;
            return service;
        }

        //finds a spreadsheet by name, or null if not found. 
        public static SpreadsheetEntry findSpreadsheetByName(string sheetName, SpreadsheetsService service)
        {
            SpreadsheetQuery query = new SpreadsheetQuery();

            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = service.Query(query);
            foreach (SpreadsheetEntry entry in feed.Entries)
            {
                // return matching spreadsheet
                if (entry.Title.Text.Equals(sheetName))
                {
                    return entry;
                }
            }

            return null;
        }

        //returns the first worksheet entry matching the specified name, or the first one in the sheet if no name is specified. 
        public static WorksheetEntry findWorksheetByName(string worksheet, SpreadsheetEntry spreadsheet)
        {
            WorksheetFeed wsFeed = spreadsheet.Worksheets;
            WorksheetEntry worksheetEntry = null;
            if (String.IsNullOrEmpty(worksheet)) //if no worksheet specified, use the first
            {
                worksheetEntry = (WorksheetEntry)wsFeed.Entries[0];
            }
            else
            {
                foreach (WorksheetEntry entry in wsFeed.Entries)
                {

                    if (entry.Title.Text.Equals(worksheet)) //find first matching worksheet
                    {
                        worksheetEntry = entry;
                        break;
                    }

                }
            }
            return worksheetEntry;
        }


        //overload for uint input
        static public string addressFromCells(uint col, uint row)
        {
            return addressFromCells((int)col, (int)row);
        }

        //overload for long input
        static public string addressFromCells(long col, long row)
        {
            return addressFromCells((int)col, (int)row);
        }

        //translates a column and row number to an A1-style address
        static public string addressFromCells(int col, int row)
        {
            return GetExcelColumnName(col) + row.ToString();
        }

        //gets the row number from a cell address
        static public uint rowFromAddress(string address)
        {
            uint result = 0;
            uint.TryParse(Regex.Replace(address, "[^.0-9]", ""), out result);
            return result;

        }


        //gets the column number from a cell address
        static public uint colFromAddress(string address)
        {
            string cell = Regex.Replace(address, @"[\d-]", String.Empty);
            return (uint)ExcelColumnNameToNumber(cell);

        }


        public static bool isValidAddress(string address)
        {
            return Regex.IsMatch(address, "^[a-zA-Z]+[1-9]+([0-9]+)?$");
        }

        //converts a column number into an A1-style column name (A, BG, etc)
        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        //converts an A1-style column name (A, BG, etc) into a column number
        private static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) return 0;

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
    }

}
