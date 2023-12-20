using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace B1SSyngentaAddOn
{
    class Menu
    {
        public void AddMenuItems()
        {
            
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "SyngentaExt";
            oCreationPackage.String = "Syngenta Extension";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception e)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("SyngentaExt");
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "B1SSyngentaAddOn.Parameters";
                oCreationPackage.String = "Parâmetros";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            

            try
            {
                SAPbouiCOM.Form form = Application.SBO_Application.Forms.ActiveForm;
                if (pVal.MenuUID == "6913" && (form.TypeEx == "50101" || form.TypeEx == "50102"))
                {
                    BubbleEvent = false;
                    return;
                }

                if (pVal.BeforeAction && pVal.MenuUID == "B1SSyngentaAddOn.Parameters")
                {
                    UIForms.Parameters activeForm = new UIForms.Parameters();
                    activeForm.UIAPIRawForm.DataSources.UserDataSources.Item("Pane").Value = "1";

                    SAPbouiCOM.ComboBox cmbSeriesPN = (SAPbouiCOM.ComboBox)activeForm.UIAPIRawForm.Items.Item("Item_12").Specific;
                    SAPbouiCOM.ComboBox cmbGrupoPN = (SAPbouiCOM.ComboBox)activeForm.UIAPIRawForm.Items.Item("Item_13").Specific;


                    while (cmbSeriesPN.ValidValues.Count > 0) { cmbSeriesPN.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index); }
                    while (cmbGrupoPN.ValidValues.Count > 0) { cmbGrupoPN.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index); }
                   

                    SAPbobsCOM.Recordset recordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    recordset.DoQuery(@"SELECT ""Series"", ""SeriesName"" FROM NNM1 WHERE ""ObjectCode"" = '2' AND ""DocSubType"" = 'S' AND ""IsManual"" = 'N'");
                    while (!recordset.EoF) { cmbSeriesPN.ValidValues.Add(recordset.Fields.Item(0).Value.ToString(), recordset.Fields.Item(1).Value.ToString()); recordset.MoveNext(); }

                    recordset = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    recordset.DoQuery(@"SELECT ""GroupCode"", ""GroupName"" FROM OCRG WHERE ""GroupType"" = 'S'");
                    while (!recordset.EoF) { cmbGrupoPN.ValidValues.Add(recordset.Fields.Item(0).Value.ToString(), recordset.Fields.Item(1).Value.ToString()); recordset.MoveNext(); }

                   
                    activeForm.UIAPIRawForm.DataBrowser.BrowseBy = "Item_2";

                    activeForm.Show();

                    //Criar thread para o B1 clicar no Menu depois
                    System.Threading.Thread Menu;
                    Menu = new System.Threading.Thread(() => ClickOK());

                    Menu.Start();
                }
            }
            catch (Exception ex)
            {
                //Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
            
        }

        public static bool ClickOK()
        {
            bool ret = true;

            try
            {
                System.Threading.Thread.Sleep(200);
                Application.SBO_Application.Menus.Item("1291").Activate();
            }
            catch (Exception)
            {
                ret = false;
            }

            return ret;
        }

    }
}
