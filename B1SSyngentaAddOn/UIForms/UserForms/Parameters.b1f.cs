using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace B1SSyngentaAddOn.UIForms
{
    [FormAttribute("B1SSyngentaAddOn.UIForms.UserForms.Parameters", "UIForms/UserForms/Parameters.b1f")]
    class Parameters : UserFormBase
    {
        public Parameters()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_1").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_5").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Item_6").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_9").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_10").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_2").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_3").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Folder Folder0;

        private void OnCustomInitialize()
        {
            this.EditText0.DataBind.SetBound(true, "@" + Program.CFGTable, "Code");
            this.EditText1.DataBind.SetBound(true, "@" + Program.CFGTable, "U_B1S_AddOnVersion");

            this.ComboBox0.DataBind.SetBound(true, "@" + Program.CFGTable, "U_B1S_BPSeries");
            this.ComboBox1.DataBind.SetBound(true, "@" + Program.CFGTable, "U_B1S_BPGroup");
            this.ComboBox2.DataBind.SetBound(true, "@" + Program.CFGTable, "U_B1S_BPPropRuralEnd");

            this.EditText0.Item.Visible = false;

        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.ComboBox ComboBox2;
    }
}
