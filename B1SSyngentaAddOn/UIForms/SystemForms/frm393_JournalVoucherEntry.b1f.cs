﻿using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace B1SSyngentaAddOn.UIForms.SystemForms
{
    [FormAttribute("393", "UIForms/SystemForms/frm393_JournalVoucherEntry.b1f")]
    class frm393_JournalVoucher : SystemFormBase
    {
        private SAPbouiCOM.EditText edit_BatchId;
        private SAPbouiCOM.EditText edit_TransId;
        private SAPbouiCOM.Matrix mtx_Lines;
        private SAPbouiCOM.EditText edit_refdate;
        private SAPbouiCOM.EditText edit_duedate;
        private SAPbouiCOM.EditText edit_taxdate;
        private SAPbouiCOM.EditText edit_memo;
        private SAPbouiCOM.EditText edit_project;
        private SAPbouiCOM.EditText edit_refOne;
        private SAPbouiCOM.EditText edit_refTwo;
        private SAPbouiCOM.EditText edit_refThree;

        private SAPbouiCOM.ComboBox cmb_series;
        private SAPbouiCOM.ComboBox cmb_ecdType;
        private SAPbouiCOM.ComboBox cmb_indicator;
        private SAPbouiCOM.ComboBox cmb_transcode;
        private SAPbouiCOM.ComboBox cmb_matriz;

        private SAPbouiCOM.CheckBox chk_cambio;
        private SAPbouiCOM.CheckBox chk_estorno;
        private SAPbouiCOM.CheckBox chk_comp;

        private SAPbouiCOM.Button btn_main;

        public frm393_JournalVoucher()
        {
            

        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.edit_BatchId = this.GetSpecificItem<SAPbouiCOM.EditText>("139");
            this.edit_TransId = this.GetSpecificItem<SAPbouiCOM.EditText>("5");
            this.edit_refdate = this.GetSpecificItem<SAPbouiCOM.EditText>("6");
            this.edit_duedate = this.GetSpecificItem<SAPbouiCOM.EditText>("102");
            this.edit_taxdate = this.GetSpecificItem<SAPbouiCOM.EditText>("97");
            this.edit_memo = this.GetSpecificItem<SAPbouiCOM.EditText>("10");
            this.edit_project = this.GetSpecificItem<SAPbouiCOM.EditText>("26");
            this.edit_refOne = this.GetSpecificItem<SAPbouiCOM.EditText>("7");
            this.edit_refTwo = this.GetSpecificItem<SAPbouiCOM.EditText>("8");
            this.edit_refThree = this.GetSpecificItem<SAPbouiCOM.EditText>("540002023");
            this.cmb_ecdType = this.GetSpecificItem<SAPbouiCOM.ComboBox>("1980000004");
            this.cmb_transcode = this.GetSpecificItem<SAPbouiCOM.ComboBox>("9");
            this.cmb_series = this.GetSpecificItem<SAPbouiCOM.ComboBox>("137");
            this.cmb_indicator = this.GetSpecificItem<SAPbouiCOM.ComboBox>("93");
            this.cmb_matriz = this.GetSpecificItem<SAPbouiCOM.ComboBox>("1320002034");
            this.chk_cambio = this.GetSpecificItem<SAPbouiCOM.CheckBox>("82");
            this.chk_estorno = this.GetSpecificItem<SAPbouiCOM.CheckBox>("99");
            this.chk_comp = this.GetSpecificItem<SAPbouiCOM.CheckBox>("95");
            this.btn_main = this.GetSpecificItem<SAPbouiCOM.Button>("1");
            this.mtx_Lines = this.GetSpecificItem<SAPbouiCOM.Matrix>("76");
            
            this.OnCustomInitialize();
        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>


        public override void OnInitializeFormEvents()
        {
            try { this.LoadAfter += this.PreLcm_LoadAfter; } catch { }
            try { this.DataLoadAfter += this.LoadDataAfter; } catch { }
        }

        private void LoadDataAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                string transId = edit_TransId.Value.ToString();
                string batchId = edit_BatchId.Value.ToString();
                string identificadorRh = GetSdrIntRhValue(batchId, transId);

                //caso seja nulo significa que o pre lcm e da integração
                if (String.IsNullOrWhiteSpace(identificadorRh) || identificadorRh == "0")
                    ChangeFormState(true);
                else
                    ChangeFormState(false);
            }
            catch { }
        }

        private void PreLcm_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            string transId = edit_TransId.Value.ToString();
            string batchId = edit_BatchId.Value.ToString();
            string identificadorRh = GetSdrIntRhValue(batchId, transId);

            //caso seja nulo significa que o pre lcm e da integração
            if (String.IsNullOrWhiteSpace(identificadorRh))
                return;

            ChangeFormState(false);

            //DisableForm();
        }

        private string GetSdrIntRhValue(string batchId, string transId)
        {
            string ret_val = "";
            string query = $"SELECT COALESCE(\"U_SDR_IntRh\",0) \"U_SDR_IntRh\" FROM OBTF WHERE TO_VARCHAR(\"BatchNum\") = '{batchId}' AND TO_VARCHAR(\"TransId\") = '{transId}' ";

            SAPbobsCOM.Recordset oRec = ((SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset));

            oRec.DoQuery(query);

            if (oRec.RecordCount == 0) ret_val = "";

            if (oRec.Fields.Item("U_SDR_IntRh").Value == null) ret_val = "";

            ret_val = oRec.Fields.Item("U_SDR_IntRh").Value.ToString();

            if(ret_val == "0")
                ret_val = "";

            //Application.SBO_Application.MessageBox("GetSdrIntRhValue: '" + ret_val + "'");

            return ret_val;
        }

        private void DisableForm()
        {
            cmb_series.Item.Enabled = false;
            edit_refdate.Item.Enabled = false;
            edit_duedate.Item.Enabled = false;
            edit_taxdate.Item.Enabled = false;
            edit_memo.Item.Enabled = false;
            cmb_indicator.Item.Enabled = false;
            edit_project.Item.Enabled = false;
            cmb_transcode.Item.Enabled = false;
            edit_refOne.Item.Enabled = false;
            edit_refTwo.Item.Enabled = false;
            edit_refThree.Item.Enabled = false;
            cmb_ecdType.Item.Enabled = false;
            cmb_matriz.Item.Enabled = false;
            btn_main.Item.Enabled = false;
            mtx_Lines.Item.Enabled = false;
            chk_cambio.Item.Enabled = false;
            chk_estorno.Item.Enabled = false;
            chk_comp.Item.Enabled = false;
        }

        private void Btn_main_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            string transId = edit_TransId.Value.ToString();
            string batchId = edit_BatchId.Value.ToString();
            string identificadorRh = GetSdrIntRhValue(batchId, transId);

            //caso seja nulo significa que o pre lcm e da integração
            if (String.IsNullOrWhiteSpace(identificadorRh))
                return;

            Application.SBO_Application.SetStatusBarMessage("Não é permitido alterar pré-lançamentos inseridos pela integração RH de forma manual.");
            BubbleEvent = false;
            btn_main.Item.Enabled = false;
        }

        private void OnCustomInitialize()
        {
            try { this.btn_main.ClickBefore += Btn_main_ClickBefore; } catch { }
        }

        private void ChangeFormState(bool newState)
        {
            try
            {
                cmb_series.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refdate.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_taxdate.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_duedate.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_memo.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_indicator.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_project.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_transcode.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refOne.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refTwo.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                edit_refThree.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_ecdType.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                cmb_matriz.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                btn_main.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                mtx_Lines.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                chk_cambio.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                chk_estorno.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

            try
            {
                chk_comp.Item.Enabled = newState;
            }
            catch (Exception)
            {
                // Tratamento de exceção
            }

        }

    }
}
