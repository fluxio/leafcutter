# Leafcutter - Google Drive Access for Grasshopper
----------------------------------------------------

## Dev Environment
#############################

Initial development on Windows 7 with Visual Studio 2010.  
Installer package created with CreateInstall Free.



## Files and Directories
#############################

- Installer Package/				Binaries and other release files
	- LeafCutterExamples/			Sample Grasshopper file with examples
	- LeafCutterFiles/				dlls and other dependencies for the compiled Grasshopper plugin (redundant?)
	- LeafCutterInstall/			Compiled installer package
	- GoogleDriveGH.sln 			Visual Studio 2010 Solution File
	- GoogleDriveGH.suo 			Visual Studio 2010 Solution Options
	- installerImg.jpg 				Source image for installer
	- LeafcutterInstallSetup.ci 	CreateInstall project file -- will require editing for new local builds
	- LeafcutterInstallSetup.log 	CreateInstall log file

- Source/							Visual Studio 2010 C# Source Code
	- bin/							dlls and other dependencies for the compiled Grasshopper plugin
		-Release/					compiled dlls and other dependencies for the Released plugin (used by CreateInstall)
	- Components/					source code for individual plugin components
	- obj/							compiler object files 
	- Organizational/				source code for custom data types and helper functions - anything that isn't a component
	- Properties/					standard Visual Studio Properties folder - assembly info + resources file
	- Resources/					Image files for the plugin components
	- GoogleDriveGH.csproj 			Visual Studio 2010 Project
	- GoogleDriveGH.csproj.user		Visual Studio 2010 Project User Options


## Dependencies
#############################

Requires DLLs in a standard Rhino + Grasshopper installation, along with some system DLLs.  You can import the DLLs into your VS 2010 project by using the Reference Paths pane of the Project Solution.


#### Grasshopper+Rhino
#############################

* GH_IO.dll
* GH_Util.dll
* Grasshopper.dll
* RhinoCommon.dll
 
#### System/.NET
#############################

* Microsoft.VisualBasic.dll
* System.dll
* System.Data.dll
* System.Drawing.dll
* System.Windows.Forms.dll
 
#### Google Data API SDK (included)
#############################

* Google.GData.AccessControl.dll
* Google.GData.Client.dll
* Google.GData.Documents.dll
* Google.GData.Extensions.dll
* Google.GData.Spreadsheets.dll
* Newtonsoft.Json.dll

Please note that you will need to get your own Google Developer API account and set up a Client ID + Client Secret.
You can obtain these codes here: https://console.developers.google.com/project


## Preparing Your Project
#############################

* Install Visual Studio 2010
* Install CreateInstall Free
* Import the Grasshopper DLL Assemblies into your VS 2010 project by using the Reference Paths pane of the Project Solution.
* Obtain Google Developer API account and place ClientID and ClientSecret in text files in the appropriate directories.  (VS2010 will point you to the right directory in the build error messages.)
* Build GH Plugin assemblies
* (optional) Edit CreateInstall .ci file and create an installer for your project.  Edit the information to reflect your identity, etc.


## Credits
#############################

Product management by Marc Syp, Application Engineer at Flux Factory, Inc.
Initial development by Andrew Heumann for Flux Factory, Inc. (flux.io)


## Additional Documentation
#############################

See the Examples.gh file for detailed explanations + examples of how to use Leafcutter in Grasshopper. 	


## Known Issues
#############################

* If multiple spreadsheets in the user's drive have the same name, it will only be possible to access the first one.
* In certain conditions auth window will create a temporary ongoing pan/scroll condition in the user's grasshopper canvas. 


## Future Improvements
#############################

* Embed Project DLLs into GHA file for easier install.
* Implement custom data type for spreadsheet reference, rather than simple string name reference. 
* Implement low level thread to check for updates to sheet, to allow "live updating" in GH


## License (see license.txt)
#############################

The MIT License (MIT)

Copyright (c) 2014 Flux Factory, Inc.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.