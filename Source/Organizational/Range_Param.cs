using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Types;
using Grasshopper.GUI;

using Rhino.Input.Custom;

namespace GoogleDriveGH
{
    //Custom Param type so that components can receive/output SpreadsheetRange data
    public class Range_Param : GH_PersistentParam<SpreadsheetRange>
    {
        //class constructor extending GH_Component - this is where the component name, nickname, description, category, and subcategory are defined. 
        public Range_Param()
            : base(new GH_InstanceDescription("Spreadsheet Range", "Range", "Defines a 2D range of cells",Properties.Resources.AssemblyName,"Spreadsheets"))
        {
        }

        //unique component id
        public override Guid ComponentGuid
        {
            get { return new Guid("{272BDDAE-F4C8-4F8E-8F21-753E08119ACA}"); }
        }
        //don't show param object in grasshopper toolbar - just use it internal to components 
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
                
                return Properties.Resources.RangeParamIcon;
            }
        }

        //This method allows the user to manually enter a spreadsheet range
        protected override ToolStripMenuItem Menu_CustomSingleValueItem()
        {
            string text = "";
            if (PersistentDataCount == 1)
            {
                SpreadsheetRange range = PersistentData.get_FirstItem(false);
                if (range != null)
                {
                    text = range.ToString();
                }
            }
            ToolStripMenuItem cItem = new ToolStripMenuItem(string.Format("Set {0}", this.TypeName));
           ToolStripTextBox dItem = GH_DocumentObject.Menu_AppendTextItem(cItem.DropDown, text, new GH_MenuTextBox.KeyDownEventHandler(this.Menu_SingleRangeValueKeyDown), new GH_MenuTextBox.TextChangedEventHandler(this.Menu_SingleRangeValueTextChanged), true, 200, true);
    
            return cItem;
        }

        //replace all standard menu items and add custom "Set one spreadsheet range" input. 
        public override void AppendAdditionalMenuItems(ToolStripDropDown menu)
        {
            this.Menu_AppendWireDisplay(menu);
            this.Menu_AppendDisconnectWires(menu);
            this.Menu_AppendPrincipalParameter(menu);
            this.Menu_AppendReverseParameter(menu);
            this.Menu_AppendFlattenParameter(menu);
            this.Menu_AppendGraftParameter(menu);
            this.Menu_AppendSimplifyParameter(menu);
            GH_DocumentObject.Menu_AppendSeparator(menu);
           
            if (this.Kind == GH_ParamKind.output)
            {
                return;
            }
            System.Windows.Forms.ToolStripMenuItem itemCustomSingle = this.Menu_CustomSingleValueItem();
            if (itemCustomSingle == null)
            {
                this.Menu_AppendPromptOne(menu);
            }
            else
            {
                itemCustomSingle.Enabled &= (this.SourceCount == 0);
                menu.Items.Add(itemCustomSingle);
            }
        
            GH_DocumentObject.Menu_AppendSeparator(menu);
            this.Menu_AppendDestroyPersistent(menu);
            this.Menu_AppendInternaliseData(menu);
            this.Menu_AppendExtractParameter(menu);
        }

        protected void Menu_SingleRangeValueTextChanged(GH_MenuTextBox sender, string text)
        {
      
                sender.TextBoxItem.ForeColor = System.Drawing.SystemColors.WindowText;
            
        }

        protected void Menu_SingleRangeValueKeyDown(GH_MenuTextBox sender, System.Windows.Forms.KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case System.Windows.Forms.Keys.Return:
                    {
                        e.Handled = true;
                        this.PersistentData.Clear();
                        string text = sender.Text;
                        if (text.Length > 0)
                        {
                                this.PersistentData.Append(new SpreadsheetRange(text));
                        }
                        this.OnObjectChanged(GH_ObjectEventType.PersistentData);
                        if (System.Windows.Forms.Control.ModifierKeys == System.Windows.Forms.Keys.Shift || System.Windows.Forms.Control.ModifierKeys == System.Windows.Forms.Keys.Control)
                        {
                            sender.CloseEntireMenuStructure();
                        }
                        this.ExpireSolution(true);
                        break;
                    }
                case System.Windows.Forms.Keys.Escape:
                    sender.CloseEntireMenuStructure();
                    break;
            }
        }



        //param must implement prompt_plural, even if the interface for it is never shown. 
        protected override GH_GetterResult Prompt_Plural(ref List<SpreadsheetRange> values)
        {
            System.Collections.Generic.List<SpreadsheetRange> nList = new System.Collections.Generic.List<SpreadsheetRange>();
            while (true)
            {
                SpreadsheetRange iInterval = null;
                switch (this.Prompt_Singular(ref iInterval))
                {
                    case GH_GetterResult.accept:
                        if (nList.Count == 0)
                        {
                            return GH_GetterResult.cancel;
                        }
                        values.AddRange(nList);
                        return GH_GetterResult.success;
                    case GH_GetterResult.success:
                        nList.Add(iInterval);
                        break;
                    case GH_GetterResult.cancel:
                        return GH_GetterResult.cancel;
                }
            }
        
           
        }

        //this method defines the commandline request to the user if (s)he selects "Set one Spreadsheet Range"

        protected override GH_GetterResult Prompt_Singular(ref SpreadsheetRange value)
        {
            string dA;
            string dB;
            if (value != null)
            {
                SpreadsheetRange value2 = new SpreadsheetRange(value);
                dA = value2.StartCell;
                dB = value2.EndCell;
            }
            else
            {
                dA = "";
                dB = "";
            }

            switch (this.Prompt_Cell("Starting cell of range", ref dA, dA))
            {
                case GH_GetterResult.accept:
                    return GH_GetterResult.accept;
                case GH_GetterResult.success:
                    
                    switch (this.Prompt_Cell("Ending cell of range", ref dB, dB))
                    {
                        case GH_GetterResult.accept:
                        case GH_GetterResult.success:
                            {
                                SpreadsheetRange value2 = new SpreadsheetRange(dA, dB);
                                value = new SpreadsheetRange(value2);
                                return GH_GetterResult.success;
                            }
                        default:
                            return GH_GetterResult.cancel;
                    }
                default:
                    return GH_GetterResult.cancel;
            }
        }

        protected GH_GetterResult Prompt_Cell(string prompt, ref string value, string @default = "")
        {
            GetString gn = new GetString();
            gn.SetCommandPrompt(prompt);
            
                gn.AcceptNothing(true);
              
            switch ((int)gn.Get())
            {
                case 2:
                    value = @default;
                    return GH_GetterResult.accept;
                case 4:
                    value = gn.StringResult();
                    return GH_GetterResult.success;
            }
            return GH_GetterResult.cancel;
        }
    }
}
