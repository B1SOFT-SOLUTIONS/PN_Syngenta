using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace B1SSyngentaAddOn.DAO
{
    static class LogicalCreation
    {
        static SAPDAO.SAPDAO Helper = new SAPDAO.SAPDAO();
        public static void criaTabelas(SAPbobsCOM.Company oCompany)
        {
            if (!Helper.TableExist(oCompany, Program.CFGTable))
                Helper.AddTableToDB(oCompany, Program.CFGTable, "B1S: Syngenta Extension", SAPbobsCOM.BoUTBTableType.bott_MasterData);

            if (!Helper.TableExist(oCompany, "B1S_EXT_DEPARTMENT"))
                Helper.AddTableToDB(oCompany, "B1S_EXT_DEPARTMENT", "B1S: Departamento", SAPbobsCOM.BoUTBTableType.bott_NoObjectAutoIncrement);

            if (!Helper.TableExist(oCompany, "SDR_INTEGRACAO_RH"))
                Helper.AddTableToDB(oCompany, "SDR_INTEGRACAO_RH", "[SEIDOR] - INTEGRAÇÃO RH", BoUTBTableType.bott_Document);

            if (!Helper.TableExist(oCompany, "SDR_INTEGRACAO_RH_L"))
                Helper.AddTableToDB(oCompany, "SDR_INTEGRACAO_RH_L", "[SEIDOR] - INTEGRAÇÃO RH LINES", BoUTBTableType.bott_DocumentLines);

            if (!Helper.TableExist(oCompany, "SDR_LOG_INTEGRACAO"))
                Helper.AddTableToDB(oCompany, "SDR_LOG_INTEGRACAO", "[SEIDOR] - LOG INTEGRAÇÃO", BoUTBTableType.bott_NoObjectAutoIncrement);

        }

        public static void criaCampos(SAPbobsCOM.Company oCompany)
        {

            SAPbobsCOM.UserFieldsMD oUserField = null;
            SAPbobsCOM.UserFieldsMD userFieldsMD = null;

            #region //OWST
            if (!Helper.FieldExist(oCompany, "U_B1S_EXT_Depart", "OWST"))
                Helper.AddFieldToTable(oCompany, "OWST", "B1S_EXT_Depart", "Departamento", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, null, "B1S_EXT_DEPARTMENT");
            #endregion

            #region //OWTM
            userFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
            userFieldsMD.ValidValues.Value = "Y";
            userFieldsMD.ValidValues.Description = "Sim";
            userFieldsMD.ValidValues.Add();
            userFieldsMD.ValidValues.Value = "N";
            userFieldsMD.ValidValues.Description = "Não";
            userFieldsMD.ValidValues.Add();

            if (!Helper.FieldExist(oCompany, "U_B1S_EXT_Justif", "OWTM"))
                Helper.AddFieldToTable(oCompany, "OWTM", "B1S_EXT_Justif", "Necessita Jusitificativa", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, userFieldsMD.ValidValues, "", "N");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(userFieldsMD);
            userFieldsMD = null;

            userFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(BoObjectTypes.oUserFields);
            userFieldsMD.ValidValues.Value = "Y";
            userFieldsMD.ValidValues.Description = "Sim";
            userFieldsMD.ValidValues.Add();
            userFieldsMD.ValidValues.Value = "N";
            userFieldsMD.ValidValues.Description = "Não";
            userFieldsMD.ValidValues.Add();

            if (!Helper.FieldExist(oCompany, "U_B1S_EXT_HomeApproval", "OWTM"))
                Helper.AddFieldToTable(oCompany, "OWTM", "B1S_EXT_HomeApproval", "Aprovar Pela Home", BoFieldTypes.db_Alpha, 1, BoFldSubTypes.st_None, userFieldsMD.ValidValues, "", "N");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(userFieldsMD);
            userFieldsMD = null;

            if (!Helper.FieldExist(oCompany, "U_B1S_EXT_Depart", "OWTM"))
                Helper.AddFieldToTable(oCompany, "OWTM", "B1S_EXT_Depart", "Departamento", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, null, "B1S_EXT_DEPARTMENT");

            if (!Helper.FieldExist(oCompany, "U_B1S_EXT_Descricao", "OWTM"))
                Helper.AddFieldToTable(oCompany, "OWTM", "B1S_EXT_Descricao", "Descricao", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, SAPbobsCOM.BoFldSubTypes.st_None);
            #endregion

            #region //OCRD
            if (!Helper.FieldExist(oCompany, "U_SD_persId", "OCRD"))
                Helper.AddFieldToTable(oCompany, "OCRD", "SD_persId", "SOLO PersonId", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, null);

            if (!Helper.FieldExist(oCompany, "U_SD_CpersId", "OCRD"))
                Helper.AddFieldToTable(oCompany, "OCRD", "SD_CpersId", "SOLO Cred PersonId", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, null);

            if (!Helper.FieldExist(oCompany, "U_SD_CardCodeC", "OCRD"))
                Helper.AddFieldToTable(oCompany, "OCRD", "SD_CardCodeC", "PN Vinculado", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, null);

            if (!Helper.FieldExist(oCompany, "U_B1S_SFID", "OCRD"))
                Helper.AddFieldToTable(oCompany, "OCRD", "B1S_SFID", "SF ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, null);

            if (!Helper.FieldExist(oCompany, "U_SOL_ConcCredito", "OCRD"))
                Helper.AddFieldToTable(oCompany, "OCRD", "SOL_ConcCredito", "Concessão do Crédito", SAPbobsCOM.BoFieldTypes.db_Date, 8, SAPbobsCOM.BoFldSubTypes.st_None, null);
            #endregion

            #region //OCPR
            if (!Helper.FieldExist(oCompany, "U_B1S_SFID", "OCRD"))
                Helper.AddFieldToTable(oCompany, "OCRD", "B1S_SFID", "SF ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, null);
            #endregion

            #region
            //CRD1
            if (!Helper.FieldExist(oCompany, "U_B1S_SFID", "CRD1"))
                Helper.AddFieldToTable(oCompany, "CRD1", "B1S_SFID", "SF ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, SAPbobsCOM.BoFldSubTypes.st_None, null);

            if (!Helper.FieldExist(oCompany, "U_SD_CardCodeC", "CRD1"))
                Helper.AddFieldToTable(oCompany, "CRD1", "SD_CardCodeC", "PN Vinculado", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, SAPbobsCOM.BoFldSubTypes.st_None, null);
            #endregion

            #region //DOC HEADER
            if (!Helper.FieldExist(oCompany, "U_B1S_SFID", "OINV"))
                Helper.AddFieldToTable(oCompany, "OINV", "B1S_SFID", "SF ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            #endregion

            #region //DOC LINE
            if (!Helper.FieldExist(oCompany, "U_B1S_SFID", "INV1"))
                Helper.AddFieldToTable(oCompany, "INV1", "B1S_SFID", "SF ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            #endregion 

            #region //ITEMS
            if (!Helper.FieldExist(oCompany, "U_B1S_SFID", "OITM"))
                Helper.AddFieldToTable(oCompany, "OITM", "B1S_SFID", "SF ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            #endregion

            #region //OJDT

            if (!Helper.FieldExist(oCompany, "U_SDR_IntRh", "OJDT"))
                Helper.AddFieldToTable(oCompany, "OJDT", "SDR_IntRh", "Identificador INT RH", BoFieldTypes.db_Numeric, 10);

            #endregion

            #region //SDR_INTEGRACAO_RH

            if (!Helper.FieldExist(oCompany, "U_empresa", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "empresa", "Empresa", BoFieldTypes.db_Alpha, 10);

            if (!Helper.FieldExist(oCompany, "U_serie", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "serie", "Serie", BoFieldTypes.db_Alpha, 10);

            if (!Helper.FieldExist(oCompany, "U_dataLancamento", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "dataLancamento", "Data Lancamento", BoFieldTypes.db_Date, 8);

            if (!Helper.FieldExist(oCompany, "U_dataVencimento", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "dataVencimento", "Data Vencimento", BoFieldTypes.db_Date, 8);

            if (!Helper.FieldExist(oCompany, "U_dataDocumento", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "dataDocumento", "Data Documento", BoFieldTypes.db_Date, 8);

            if (!Helper.FieldExist(oCompany, "U_observacao", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "observacao", "Observacao", BoFieldTypes.db_Alpha, 150);

            if (!Helper.FieldExist(oCompany, "U_indicador", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "indicador", "Indicador", BoFieldTypes.db_Alpha, 150);

            if (!Helper.FieldExist(oCompany, "U_projeto", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "projeto", "Projeto", BoFieldTypes.db_Alpha, 100);

            if (!Helper.FieldExist(oCompany, "U_codigoTransacao", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "codigoTransacao", "Codigo Transacao", BoFieldTypes.db_Alpha, 150);

            if (!Helper.FieldExist(oCompany, "U_referencia", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "referencia", "Referencia", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_referencia2", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "referencia2", "Referencia 2", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_referencia3", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "referencia3", "Referencia 3", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_ECD", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "ECD", "ECD", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_status", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "status", "Status", BoFieldTypes.db_Alpha, 10);

            if (!Helper.FieldExist(oCompany, "U_DataInicio", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "DataInicio", "Data Inicio", BoFieldTypes.db_Date, 8);

            if (!Helper.FieldExist(oCompany, "U_DataFim", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "DataFim", "Data Fim", BoFieldTypes.db_Date, 8);

            if (!Helper.FieldExist(oCompany, "U_Mensagem", "@SDR_INTEGRACAO_RH"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH", "Mensagem", "Mensagem", BoFieldTypes.db_Memo, 250);

            #endregion

            #region //SDR_INTEGRACAO_RH_L

            if (!Helper.FieldExist(oCompany, "U_tipoConta", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "tipoConta", "Tipo Conta", BoFieldTypes.db_Alpha, 150);

            if (!Helper.FieldExist(oCompany, "U_contaContabil", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "contaContabil", "Conta Contabil", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_valorCredito", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "valorCredito", "Valor Credito", BoFieldTypes.db_Numeric, 10);

            if (!Helper.FieldExist(oCompany, "U_valorDebito", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "valorDebito", "Valor Debito", BoFieldTypes.db_Numeric, 10);

           if (!Helper.FieldExist(oCompany, "U_observacao", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "observacao", "Observacao", BoFieldTypes.db_Alpha, 250);
            
            if (!Helper.FieldExist(oCompany, "U_referencia", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "referencia", "Referencia", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_referencia2", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "referencia2", "Referencia 2", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_referencia3", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "referencia3", "Referencia 3", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_dataVencimento", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "dataVencimento", "Data Vencimento", BoFieldTypes.db_Date, 8);

            if (!Helper.FieldExist(oCompany, "U_centroCusto", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "centroCusto", "Centro Custo", BoFieldTypes.db_Alpha, 250);

            if (!Helper.FieldExist(oCompany, "U_filial", "@SDR_INTEGRACAO_RH_L"))
                Helper.AddFieldToTable(oCompany, "@SDR_INTEGRACAO_RH_L", "filial", "Filial", BoFieldTypes.db_Alpha, 250);

            #endregion

            #region //SDR_LOG_INTEGRACAO

            if (!Helper.FieldExist(oCompany, "U_CNPJ", "@SDR_LOG_INTEGRACAO"))
                Helper.AddFieldToTable(oCompany, "@SDR_LOG_INTEGRACAO", "CNPJ", "CNPJ", BoFieldTypes.db_Alpha, 25);

            if (!Helper.FieldExist(oCompany, "U_Desc", "@SDR_LOG_INTEGRACAO"))
                Helper.AddFieldToTable(oCompany, "@SDR_LOG_INTEGRACAO", "Desc", "Descricao", BoFieldTypes.db_Alpha, 25);

            if (!Helper.FieldExist(oCompany, "U_Sys_Destino", "@SDR_LOG_INTEGRACAO"))
                Helper.AddFieldToTable(oCompany, "@SDR_LOG_INTEGRACAO", "Sys_Destino", "Destino", BoFieldTypes.db_Alpha, 25);

            if (!Helper.FieldExist(oCompany, "U_Data", "@SDR_LOG_INTEGRACAO"))
                Helper.AddFieldToTable(oCompany, "@SDR_LOG_INTEGRACAO", "Data", "Data", BoFieldTypes.db_Alpha, 25);

            if (!Helper.FieldExist(oCompany, "U_Hora", "@SDR_LOG_INTEGRACAO"))
                Helper.AddFieldToTable(oCompany, "@SDR_LOG_INTEGRACAO", "Hora", "Hora", BoFieldTypes.db_Alpha, 25);

            if (!Helper.FieldExist(oCompany, "U_Erro", "@SDR_LOG_INTEGRACAO"))
                Helper.AddFieldToTable(oCompany, "@SDR_LOG_INTEGRACAO", "Erro", "Erro", BoFieldTypes.db_Alpha, 25);

            #endregion

            //CONFIG
            if (!Helper.FieldExist(oCompany, "U_B1S_AddOnVersion", "@" + Program.CFGTable))
                Helper.AddFieldToTable(oCompany, "@" + Program.CFGTable, "B1S_AddOnVersion", "Versão do addOn", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
           

            // Pane 1
            if (!Helper.FieldExist(oCompany, "U_B1S_BPSeries", "@" + Program.CFGTable))
                Helper.AddFieldToTable(oCompany, "@" + Program.CFGTable, "B1S_BPSeries", "Série PN", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            

            if (!Helper.FieldExist(oCompany, "U_B1S_BPGroup", "@" + Program.CFGTable))
                Helper.AddFieldToTable(oCompany, "@" + Program.CFGTable, "B1S_BPGroup", "Grupo PN", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            

           
        }

        public static void criaRegistroUDO(SAPbobsCOM.Company oCompany)
        {
            UDORegister(Program.CFGTable);
        }
        static void UDORegister(string TableName, string[] ChildTables = null)
        {
            //if (ObjCreation.RegisterUDONoChildrenIfNotExists("DMTAXCFG", SAPbobsCOM.BoUDOObjType.boud_MasterData))
            //    Msg.ShortMessageBar("REGISTRO UDO: [DMTAXCFG] registrado/atualizado com sucesso.", false);
            SAPbobsCOM.Recordset oUDO = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oUDO.DoQuery(@"SELECT * FROM OUDO WHERE ""Code"" = '" + TableName + "'");

            if (oUDO.RecordCount == 0)
            {
                //Config
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO);
                SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.Code = TableName;
                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.Name = TableName;
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.TableName = TableName;

                if (ChildTables != null)
                {
                    foreach (var item in ChildTables)
                    {
                        oUserObjectMD.ChildTables.Add();
                        oUserObjectMD.ChildTables.TableName = item.ToString();

                    }
                }


                int udoAdd = oUserObjectMD.Add();
                if (udoAdd != 0)
                {
                    string err = Program.oCompany.GetLastErrorDescription();
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(err, SAPbouiCOM.BoMessageTime.bmt_Short, true);

                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUDO);



        }
        public static void criaInsercaoRegistros(SAPbobsCOM.Company oCompany)
        {
            insereRegistroConf(oCompany, Program.CFGTable, "U_B1S_AddOnVersion", Program.addOnVersion);
        }
        public static void insereRegistroConf(SAPbobsCOM.Company oCompany, string TableName, string VersionFieldName, string addOnVersion)
        {
            SAPbobsCOM.Recordset oCheck = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCheck.DoQuery(@"SELECT 1 FROM ""@" + TableName + @""" WHERE ""Code"" = '1'");

            if (oCheck.RecordCount == 0)
            {

                //Instancia o Compaby Service
                SAPbobsCOM.CompanyService sCmp;
                //Pega a conexão atual
                sCmp = oCompany.GetCompanyService();

                //Instancia os Serviços de UDO
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralDataMAIN;
                //SAPbobsCOM.GeneralDataCollection oGeneralDataCHILD;
                //SAPbobsCOM.GeneralData oGeneralDataCHILDLines;
                SAPbobsCOM.GeneralDataParams oGeneralParams;

                try
                {
                    oGeneralService = sCmp.GetGeneralService(TableName);
                    oGeneralDataMAIN = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oGeneralDataMAIN.SetProperty("Code", "1");
                    oGeneralDataMAIN.SetProperty(VersionFieldName, addOnVersion);
                    oGeneralService.Add(oGeneralDataMAIN);

                    //Define o UDO
                    oGeneralService = sCmp.GetGeneralService(TableName);


                    //Dados do UDO
                    oGeneralDataMAIN = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                    oGeneralParams.SetProperty("Code", "1");
                    oGeneralDataMAIN = oGeneralService.GetByParams(oGeneralParams);

                    //Atualiza o UDO
                    oGeneralService.Update(oGeneralDataMAIN);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralDataMAIN);

                }
                catch (Exception er)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Erro ao criar/atualizar registro de configuração. Motivo: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            else
            {
                SAPbobsCOM.Recordset UpdateVersion = (SAPbobsCOM.Recordset)Program.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                UpdateVersion.DoQuery(@"UPDATE ""@" + TableName + @""" SET """ + VersionFieldName + @""" = '" + addOnVersion + @"' WHERE ""Code"" = 1");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UpdateVersion);
            }
        }

        public static bool SetFirstCodeOnMasterDataTable(SAPbobsCOM.Company oCompany, string TableName)
        {
            SAPbobsCOM.Recordset oCheck = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCheck.DoQuery(@"SELECT 1 FROM ""@" + TableName + @""" WHERE ""Code"" = '1'");
            bool ret = false;

            if (oCheck.RecordCount == 0)
            {

                //Instancia o Compaby Service
                SAPbobsCOM.CompanyService sCmp;
                //Pega a conexão atual
                sCmp = oCompany.GetCompanyService();

                //Instancia os Serviços de UDO
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralDataMAIN;
                //SAPbobsCOM.GeneralDataCollection oGeneralDataCHILD;
                //SAPbobsCOM.GeneralData oGeneralDataCHILDLines;
                SAPbobsCOM.GeneralDataParams oGeneralParams;

                try
                {
                    oGeneralService = sCmp.GetGeneralService(TableName);
                    oGeneralDataMAIN = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oGeneralDataMAIN.SetProperty("Code", "1");
                    oGeneralService.Add(oGeneralDataMAIN);

                    //Define o UDO
                    oGeneralService = sCmp.GetGeneralService(TableName);


                    //Dados do UDO
                    oGeneralDataMAIN = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                    oGeneralParams.SetProperty("Code", "1");
                    oGeneralDataMAIN = oGeneralService.GetByParams(oGeneralParams);

                    //Atualiza o UDO
                    oGeneralService.Update(oGeneralDataMAIN);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralDataMAIN);

                    ret = true;
                }
                catch (Exception er)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Erro ao criar/atualizar registro de configuração. Motivo: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    ret = false;
                }
            }

            return ret;

        }

        public static bool SetFirstCodeOrUpdateVersionValueOnConfigTable(SAPbobsCOM.Company oCompany, string TableName, string VersionFieldName, string addOnVersion)
        {
            SAPbobsCOM.Recordset oCheck = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oCheck.DoQuery(@"SELECT 1 FROM ""@" + TableName + @""" WHERE ""Code"" = '1'");

            if (oCheck.RecordCount == 0)
            {

                //Instancia o Compaby Service
                SAPbobsCOM.CompanyService sCmp;
                //Pega a conexão atual
                sCmp = oCompany.GetCompanyService();

                //Instancia os Serviços de UDO
                SAPbobsCOM.GeneralService oGeneralService;
                SAPbobsCOM.GeneralData oGeneralDataMAIN;
                //SAPbobsCOM.GeneralDataCollection oGeneralDataCHILD;
                //SAPbobsCOM.GeneralData oGeneralDataCHILDLines;
                SAPbobsCOM.GeneralDataParams oGeneralParams;

                try
                {
                    oGeneralService = sCmp.GetGeneralService(TableName);
                    oGeneralDataMAIN = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oGeneralDataMAIN.SetProperty("Code", "1");
                    oGeneralDataMAIN.SetProperty(VersionFieldName, addOnVersion);
                    oGeneralService.Add(oGeneralDataMAIN);

                    //Define o UDO
                    oGeneralService = sCmp.GetGeneralService(TableName);


                    //Dados do UDO
                    oGeneralDataMAIN = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                    oGeneralParams.SetProperty("Code", "1");
                    oGeneralDataMAIN = oGeneralService.GetByParams(oGeneralParams);

                    //Atualiza o UDO
                    oGeneralService.Update(oGeneralDataMAIN);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralDataMAIN);

                    return true;
                }
                catch (Exception er)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Erro ao criar/atualizar registro de configuração. Motivo: " + er.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return false;
                }
            }
            else
            {
                SAPbobsCOM.Recordset UpdateVersion = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                UpdateVersion.DoQuery(@"UPDATE ""@" + TableName + @""" SET """ + VersionFieldName + @""" = '" + addOnVersion + @"' WHERE ""Code"" = 1");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UpdateVersion);
                return true;
            }

        }

        public static void CriaAutoriz(SAPbobsCOM.Company oCompany)
        {

            int RetVal, ErrCode;
            string ErrMsg;

            Console.WriteLine("Criando permissões na base: " + oCompany.CompanyDB);

            //RetVal = oCompany.Connect();

            //if (RetVal != 0)
            //{
            //    oCompany.GetLastError(out ErrCode, out ErrMsg);
            //    throw new Exception("Erro conectando ao SAP: " + ErrMsg);
            //}

            //Console.WriteLine("Criando campos na base: " + oCompany.CompanyDB);

            var intHelper = new SAPDAO.SAPDAO();

            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                intHelper = new SAPDAO.SAPDAO();
            }
            else
            {
                //intHelper = new DAO_SQL();
            }
            UserPermissionTree userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

            if (!intHelper.PermExists(oCompany, "SD_APPHeader"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);
                Console.WriteLine("Permissão: SD_APPHeader");
                userPermissionTree.Name = "APP Aprovações";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.PermissionID = "SD_APPHeader";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }


            //Pedido
            if (!intHelper.PermExists(oCompany, "SD_APPPedido"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);
                Console.WriteLine("Permissão: SD_APPPedido");
                userPermissionTree.Name = "Pedido";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPHeader";
                userPermissionTree.PermissionID = "SD_APPPedido";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPCliForn"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);
                Console.WriteLine("Permissão: SD_APPCliForn");

                userPermissionTree.Name = "Cliente / Fornecedor";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPCliForn";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }



            if (!intHelper.PermExists(oCompany, "SD_APPNrPedido"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Nr Pedido";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPNrPedido";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPTotPedido"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Total do Pedido";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPTotPedido";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }


            if (!intHelper.PermExists(oCompany, "SD_APPFilial"))
            {

                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Filial";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPFilial";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }


            if (!intHelper.PermExists(oCompany, "SD_APPSolicit"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Solicitante";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPSolicit";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }


            if (!intHelper.PermExists(oCompany, "SD_APPMargPed"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Margem do Pedido";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPMargPed";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPVendedor"))
            {

                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Vendedor";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPVendedor";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPGrpFamCli"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Grupo Familiar do Cliente";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPGrpFamCli";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPCredAt"))
            {
                //Disponivel
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Crédito Atual";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPCredAt";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPCredAp"))
            {

                //Limite Aprovad - CreditLimit
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Crédito Aprovado";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPCredAp";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPCredExc"))
            {
                //Excedido é Calculado
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Crédito Excedido";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPCredExc";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPRating"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Rating";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPRating";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPDtVctoCr"))
            {

                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Data de vcto. Crédito";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPDtVctoCr";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPDtConcCred"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Data de concessão Crédito";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPDtConcCred";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPCiclo"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Ciclo";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPCiclo";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPDtPagto"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Data de Pagamento";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPDtPagto";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPVlTotalPed"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Total do Pedido";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPPedido";
                userPermissionTree.PermissionID = "SD_APPVlTotalPed";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPItem"))
            {
                //Item
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Item";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPHeader";
                userPermissionTree.PermissionID = "SD_APPItem";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPProduto"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Produto";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPItem";
                userPermissionTree.PermissionID = "SD_APP" + userPermissionTree.Name;
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPMargem"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Margem";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPItem";
                userPermissionTree.PermissionID = "SD_APP" + userPermissionTree.Name;
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPDesconto"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Desconto";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPItem";
                userPermissionTree.PermissionID = "SD_APP" + userPermissionTree.Name;
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPQtde"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Quantidade";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPItem";
                userPermissionTree.PermissionID = "SD_APPQtde";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPValUnit"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Valor Unitário";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPItem";
                userPermissionTree.PermissionID = "SD_APPValUnit";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }

            if (!intHelper.PermExists(oCompany, "SD_APPDtEntrega"))
            {
                userPermissionTree = (UserPermissionTree)oCompany.GetBusinessObject(BoObjectTypes.oUserPermissionTree);

                userPermissionTree.Name = "Data de Entrega";
                userPermissionTree.Options = BoUPTOptions.bou_FullNone;
                userPermissionTree.ParentID = "SD_APPItem";
                userPermissionTree.PermissionID = "SD_APPDtEntrega";
                RetVal = userPermissionTree.Add();
                if (RetVal != 0)
                {
                    oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception("Erro criando autorização no SAP: " + ErrMsg);
                }
            }



        }
    }
}
