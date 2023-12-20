using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Globalization;

namespace B1SSyngentaAddOn.DAO.SAPDAO
{
    class SAPDAO
    {
        //public static DAO_HANA singleton;
        private Company _company;
        //private SAPbouiCOM.Application _SBO_Application;
        IFormatProvider usCulture = new CultureInfo("en-US");

        public Company company
        {
            set { _company = value; }
        }

        public bool PermExists(SAPbobsCOM.Company company, string PermissaoID)
        {
            UserPermissionTree userPermissionTree = (UserPermissionTree)company.GetBusinessObject(BoObjectTypes.oUserPermissionTree);
            return userPermissionTree.GetByKey(PermissaoID);

        }
        public void AddTableToDB(SAPbobsCOM.Company oCompany, string TblName, string tableDesc, SAPbobsCOM.BoUTBTableType TblType)
        {

            int RetVal, ErrCode;
            string ErrMsg;
            //            string ErrCode, ErrMsg;
            SAPbobsCOM.UserTablesMD oUserTable = (SAPbobsCOM.UserTablesMD)oCompany.GetBusinessObject(BoObjectTypes.oUserTables);
            oUserTable.TableName = TblName;
            oUserTable.TableDescription = tableDesc;
            oUserTable.TableType = TblType;
            RetVal = oUserTable.Add();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Erro criando tabela [" + TblName + "] no SAP: " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
            }
            else
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Tabela  [" + TblName + "] criada com sucesso no SAP: ", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }


        }

        // Adds a User field to a user table in the DB        
        public void AddFieldToTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize)
        {
            AddFieldToTable(oCompany, TblName, FldName, FldDesc, fType, fSize, SAPbobsCOM.BoFldSubTypes.st_None, null, "");
        }

        public void AddFieldToTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType)
        {
            AddFieldToTable(oCompany, TblName, FldName, FldDesc, fType, fSize, subType, null, "");
        }

        public void AddFieldToTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue)
        {
            AddFieldToTable(oCompany, TblName, FldName, FldDesc, fType, fSize, subType, validValue, "", "");
        }

        public void AddFieldToTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue, String LinkedTable)
        {
            AddFieldToTable(oCompany, TblName, FldName, FldDesc, fType, fSize, subType, validValue, LinkedTable, "");
        }

        public void AddFieldToTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue, String LinkedTable, String DefaultValue)
        {

            int RetVal = -10, ErrCode = -100;
            string ErrMsg;

            SAPbobsCOM.UserFieldsMD oUserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            oUserField = null;
            oUserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            oUserField.TableName = TblName;
            oUserField.Name = FldName;
            oUserField.Description = FldDesc;
            oUserField.Type = fType;
            oUserField.SubType = subType;
            oUserField.DefaultValue = DefaultValue;
            if (validValue != null)
            {
                for (int i = 0; i < validValue.Count - 1; i++)
                {
                    validValue.SetCurrentLine(i);
                    oUserField.ValidValues.Value = validValue.Value;
                    oUserField.ValidValues.Description = validValue.Description;
                    oUserField.ValidValues.Add();
                }
            }
            if (LinkedTable != "")
                oUserField.LinkedTable = LinkedTable;

            if (fSize > 0)
                oUserField.EditSize = fSize;

            RetVal = oUserField.Add();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Erro criando campo no SAP: " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
                GC.Collect();
            }
            else
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(string.Format("Campo [{0}] com sucesso na tabela [{1}].", oUserField.Name, oUserField.TableName), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();

            }




        }

        public void AddFieldToTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue, String LinkedTable, String DefaultValue, UDFLinkedSystemObjectTypesEnum SystemObject)
        {

            int RetVal = -10, ErrCode = -100;
            string ErrMsg;

            SAPbobsCOM.UserFieldsMD oUserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            oUserField = null;
            oUserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            oUserField.TableName = TblName;
            oUserField.Name = FldName;
            oUserField.Description = FldDesc;
            oUserField.Type = fType;
            oUserField.SubType = subType;
            oUserField.DefaultValue = DefaultValue;
            oUserField.LinkedSystemObject = (UDFLinkedSystemObjectTypesEnum)SystemObject;
            if (validValue != null)
            {
                for (int i = 0; i < validValue.Count - 1; i++)
                {
                    validValue.SetCurrentLine(i);
                    oUserField.ValidValues.Value = validValue.Value;
                    oUserField.ValidValues.Description = validValue.Description;
                    oUserField.ValidValues.Add();
                }
            }
            if (LinkedTable != "")
                oUserField.LinkedTable = LinkedTable;

            if (fSize > 0)
                oUserField.EditSize = fSize;

            RetVal = oUserField.Add();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("Erro criando campo no SAP: " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short);
                GC.Collect();
            }
            else
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(string.Format("Campo [{0}] com sucesso na tabela [{1}].", oUserField.Name, oUserField.TableName), SAPbouiCOM.BoMessageTime.bmt_Short, false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();

            }




        }
        public void UpdateFieldInTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize)
        {
            UpdateFieldInTable(oCompany, TblName, FldName, FldDesc, fType, fSize, SAPbobsCOM.BoFldSubTypes.st_None, null, "");
        }

        public void UpdateFieldInTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType)
        {
            UpdateFieldInTable(oCompany, TblName, FldName, FldDesc, fType, fSize, subType, null, "");
        }

        public void UpdateFieldInTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue)
        {
            UpdateFieldInTable(oCompany, TblName, FldName, FldDesc, fType, fSize, subType, validValue, "", "");
        }

        public void UpdateFieldInTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue, String LinkedTable)
        {
            UpdateFieldInTable(oCompany, TblName, FldName, FldDesc, fType, fSize, subType, validValue, LinkedTable, "");
        }

        public void UpdateFieldInTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue, String LinkedTable, String DefaultValue)
        {

            int RetVal, ErrCode, FieldId;
            string ErrMsg;

            string CharUserTable = "";

            if (TableExist(oCompany, TblName))
                CharUserTable = "@";

            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";

            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT \"FieldID\" FROM \"CUFD\" WHERE \"TableID\" = '{2}{0}' AND \"AliasID\" = '{1}'", TblName, FldName, CharUserTable);
            else
                sql = String.Format("SELECT FieldID FROM CUFD WHERE TableID = {2}'{0}' AND AliasID = '{1}'", TblName, FldName, CharUserTable);

            try
            {
                rs.DoQuery(sql);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro atualizando tabela no SAP: " + ex.Message);
            }
            finally
            {
                FieldId = (int)rs.Fields.Item(0).Value;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();
            }

            SAPbobsCOM.UserFieldsMD oUserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            oUserField.GetByKey(CharUserTable + TblName, FieldId);
            //oUserField.TableName = TblName;
            //oUserField.Name = FldName;
            oUserField.Description = FldDesc;
            //oUserField.Type = fType;
            //oUserField.SubType = subType;
            oUserField.DefaultValue = DefaultValue;
            if (validValue != null)
            {

                for (int i = 0; i < validValue.Count - 1; i++)
                {
                    validValue.SetCurrentLine(i);
                    if (i >= oUserField.ValidValues.Count)
                    {
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = validValue.Value;
                        oUserField.ValidValues.Description = validValue.Description;
                    }
                    else
                    {
                        oUserField.ValidValues.SetCurrentLine(i);
                        //oUserField.ValidValues.Value = validValue.Value;
                        oUserField.ValidValues.Description = validValue.Description;
                    }

                }
            }
            if (LinkedTable != "")
                oUserField.LinkedTable = LinkedTable;

            if (fSize > 0)
                oUserField.EditSize = fSize;

            RetVal = oUserField.Update();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                throw new Exception("Erro atualizando tabela no SAP: " + ErrMsg);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);

            GC.Collect();
        }

        public void UpdateFieldInTable(SAPbobsCOM.Company oCompany, string TblName, String FldName, string FldDesc, SAPbobsCOM.BoFieldTypes fType, int fSize,
            SAPbobsCOM.BoFldSubTypes subType, SAPbobsCOM.ValidValuesMD validValue, String LinkedTable, String DefaultValue, UDFLinkedSystemObjectTypesEnum SystemObject)
        {

            int RetVal, ErrCode, FieldId;
            string ErrMsg;

            string CharUserTable = "";

            if (TableExist(oCompany, TblName))
                CharUserTable = "@";

            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";

            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT \"FieldID\" FROM \"CUFD\" WHERE \"TableID\" = '{2}{0}' AND \"AliasID\" = '{1}'", TblName, FldName, CharUserTable);
            else
                sql = String.Format("SELECT FieldID FROM CUFD WHERE TableID = {2}'{0}' AND AliasID = '{1}'", TblName, FldName, CharUserTable);

            try
            {
                rs.DoQuery(sql);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro atualizando tabela no SAP: " + ex.Message);
            }
            finally
            {
                FieldId = (int)rs.Fields.Item(0).Value;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();
            }

            SAPbobsCOM.UserFieldsMD oUserField = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            oUserField.GetByKey(CharUserTable + TblName, FieldId);
            //oUserField.TableName = TblName;
            //oUserField.Name = FldName;
            oUserField.Description = FldDesc;
            //oUserField.Type = fType;
            //oUserField.SubType = subType;
            oUserField.DefaultValue = DefaultValue;
            oUserField.LinkedSystemObject = (UDFLinkedSystemObjectTypesEnum)SystemObject;
            if (validValue != null)
            {

                for (int i = 0; i < validValue.Count - 1; i++)
                {
                    validValue.SetCurrentLine(i);
                    if (i >= oUserField.ValidValues.Count)
                    {
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = validValue.Value;
                        oUserField.ValidValues.Description = validValue.Description;
                    }
                    else
                    {
                        oUserField.ValidValues.SetCurrentLine(i);
                        //oUserField.ValidValues.Value = validValue.Value;
                        oUserField.ValidValues.Description = validValue.Description;
                    }

                }
            }
            if (LinkedTable != "")
                oUserField.LinkedTable = LinkedTable;

            if (fSize > 0)
                oUserField.EditSize = fSize;

            RetVal = oUserField.Update();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                throw new Exception("Erro atualizando tabela no SAP: " + ErrMsg);
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);

            GC.Collect();
        }

        // TableExist Check whether a User Table already exists
        public bool TableExist(SAPbobsCOM.Company oCompany, string TblName)
        {

            SAPbobsCOM.UserTable oUserTable = null;
            int count;
            count = oCompany.UserTables.Count;
            for (int i = 0; i < count; i++)
            {
                oUserTable = oCompany.UserTables.Item(i);
                if (oUserTable.TableName == TblName)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                    oUserTable = null;
                    GC.Collect();
                    return true;
                }

            }

            if (oUserTable != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
            }

            return false;

        }


        public bool QueryCategoryExist(SAPbobsCOM.Company oCompany, string QueryCategoryName)
        {
            bool RetValue = false;
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT count(*) FROM \"OQCN\" WHERE \"CatName\" = '{0}'", QueryCategoryName);
            else
                sql = String.Format("SELECT count(*) FROM OQCN WHERE CatName = '{0}'", QueryCategoryName);
            try
            {
                rs.DoQuery(sql);
            }
            catch
            {
                RetValue = false;
                return RetValue;
            }
            finally
            {
                if ((int)rs.Fields.Item(0).Value == 1)
                {
                    RetValue = true;
                }
                else
                {
                    RetValue = false;

                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            GC.Collect();
            return RetValue;
        }

        public void AddQueryCategory(SAPbobsCOM.Company oCompany, string QueryCategoryName)
        {
            int RetVal, ErrCode;
            string ErrMsg;

            SAPbobsCOM.QueryCategories QueryCategory = (SAPbobsCOM.QueryCategories)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);

            QueryCategory.Name = QueryCategoryName;
            QueryCategory.Permissions = "YYYYYYYYYYYYYYY";

            RetVal = QueryCategory.Add();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(QueryCategory);
                throw new Exception("Erro criando Categoria de Consultas no SAP: " + ErrMsg);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(QueryCategory);
            GC.Collect();
        }

        public bool UserQueriesExist(SAPbobsCOM.Company oCompany, string QueryCategoryName, string UserQueriesName)
        {
            bool RetValue = false;
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT count(*) FROM \"OUQR\" T0 INNER JOIN \"OQCN\" T1 ON T0.\"QCategory\" = T1.\"CategoryId\" WHERE T0.\"QName\" = '{0}' AND T1.\"CatName\" = '{1}'", UserQueriesName, QueryCategoryName);
            else
                sql = String.Format("SELECT count(*) FROM OUQR T0 INNER JOIN OQCN T1 ON T0.QCategory = T1.CategoryId WHERE T0.QName = '{0}' AND T1.CatName = '{1}'", UserQueriesName, QueryCategoryName);

            try
            {
                rs.DoQuery(sql);
            }
            catch
            {
                RetValue = false;
                return RetValue;
            }
            finally
            {
                if ((int)rs.Fields.Item(0).Value == 1)
                {
                    RetValue = true;
                }
                else
                {
                    RetValue = false;

                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                GC.Collect();
            }
            return RetValue;
        }

        public void AddUserQueries(SAPbobsCOM.Company oCompany, string QueryCategoryName, string UserQueriesName, string QueryTXT)
        {
            int RetVal, ErrCode, CategoryId;
            string ErrMsg;

            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT \"CategoryId\" FROM \"OQCN\" WHERE \"CatName\" = '{0}'", QueryCategoryName);
            else
                sql = String.Format("SELECT CategoryId FROM OQCN WHERE CatName = '{0}'", QueryCategoryName);

            try
            {
                rs.DoQuery(sql);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro criando Consulta no SAP: " + ex.Message);
            }
            finally
            {
                CategoryId = (int)rs.Fields.Item(0).Value;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

            SAPbobsCOM.UserQueries UserQuery = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);

            UserQuery.QueryCategory = CategoryId;
            UserQuery.QueryDescription = UserQueriesName;
            UserQuery.Query = QueryTXT;

            RetVal = UserQuery.Add();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserQuery);
                throw new Exception("Erro criando Consulta no SAP: " + ErrMsg);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(UserQuery);
            GC.Collect();
        }

        public string GetFormID_UserTable(SAPbobsCOM.Company oCompany, string UserTableName)
        {
            string ReturnStr = "";

            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT '11'+CAST(\"TblNum\" AS VARCHAR) FROM \"OUTB\" WHERE \"TableName\" = '{0}'", UserTableName);
            else
                sql = String.Format("SELECT '11'+CAST(TblNum AS VARCHAR(MAX)) FROM OUTB WHERE TableName = '{0}'", UserTableName);
            try
            {
                rs.DoQuery(sql);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro buscando código da tabela no SAP: " + ex.Message);
            }
            finally
            {
                ReturnStr = rs.Fields.Item(0).Value.ToString();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            GC.Collect();
            return ReturnStr;
        }

        public void UpdateUserQueries(SAPbobsCOM.Company oCompany, string QueryCategoryName, string UserQueriesName, string QueryTXT)
        {
            int RetVal, ErrCode, CategoryId, IntrnalKey;
            string ErrMsg;

            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT T0.\"IntrnalKey\", T0.\"QCategory\" FROM \"OUQR\" T0 INNER JOIN \"OQCN\" T1 ON T0.\"QCategory\" = T1.\"CategoryId\" WHERE T0.\"QName\" = '{0}' AND T1.\"CatName\" = '{1}'", UserQueriesName, QueryCategoryName);
            else
                sql = String.Format("SELECT T0.IntrnalKey, T0.QCategory FROM OUQR T0 INNER JOIN OQCN T1 ON T0.QCategory = T1.CategoryId WHERE T0.QName = '{0}' AND T1.CatName = '{1}'", UserQueriesName, QueryCategoryName);
            try
            {
                rs.DoQuery(sql);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro atualizando Consulta no SAP: " + ex.Message);
            }
            finally
            {
                CategoryId = (int)rs.Fields.Item(1).Value;
                IntrnalKey = (int)rs.Fields.Item(0).Value;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

            SAPbobsCOM.UserQueries UserQuery = (SAPbobsCOM.UserQueries)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);

            if (!UserQuery.GetByKey(IntrnalKey, CategoryId))
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserQuery);
                throw new Exception("Erro atualizando Consulta no SAP: " + ErrMsg);
            }

            UserQuery.Query = QueryTXT;

            RetVal = UserQuery.Update();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserQuery);
                throw new Exception("Erro atualizando Consulta no SAP: " + ErrMsg);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(UserQuery);
            GC.Collect();
        }

        public bool FormattedSearchesExist(SAPbobsCOM.Company oCompany, string FormID, string ItemID, string ColID)
        {
            bool RetValue = false;
            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT COUNT(*) FROM \"CSHS\" WHERE \"FormID\" = '{0}' AND \"ItemID\" = '{1}' AND \"ColID\" = '{2}'", FormID, ItemID, ColID);
            else
                sql = String.Format("SELECT COUNT(*) FROM CSHS WHERE FormID = '{0}' AND ItemID = '{1}' AND ColID = '{2}'", FormID, ItemID, ColID);
            try
            {
                rs.DoQuery(sql);
            }
            catch
            {
                RetValue = false;
                return RetValue;
            }
            finally
            {
                if ((int)rs.Fields.Item(0).Value == 1)
                {
                    RetValue = true;
                }
                else
                {
                    RetValue = false;

                }
                GC.Collect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
            return RetValue;
        }

        public void AddFormattedSearches(SAPbobsCOM.Company oCompany, string FormID, string ItemID, string ColID, BoFormattedSearchActionEnum Action, BoYesNoEnum ForceRefresh, BoYesNoEnum ByField, BoYesNoEnum Refresh, string FieldID, string QueryCategoryName, string UserQueriesName)
        {
            int RetVal, ErrCode, IntrnalKey;
            string ErrMsg;

            Recordset rs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("SELECT T0.\"IntrnalKey\" FROM \"OUQR\" T0 INNER JOIN \"OQCN\" T1 ON T0.\"QCategory\" = T1.\"CategoryId\" WHERE T0.\"QName\" = '{0}' AND T1.\"CatName\" = '{1}'", UserQueriesName, QueryCategoryName);
            else
                sql = String.Format("SELECT T0.IntrnalKey FROM OUQR T0 INNER JOIN OQCN T1 ON T0.QCategory = T1.CategoryId WHERE T0.QName = '{0}' AND T1.CatName = '{1}'", UserQueriesName, QueryCategoryName);
            try
            {
                rs.DoQuery(sql);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro criando Consulta Formatada no SAP: " + ex.Message);
            }
            finally
            {
                IntrnalKey = (int)rs.Fields.Item(0).Value;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }

            SAPbobsCOM.FormattedSearches FormattedSearch = (SAPbobsCOM.FormattedSearches)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);

            FormattedSearch.FormID = FormID;
            FormattedSearch.ItemID = ItemID;
            FormattedSearch.ColumnID = ColID;
            FormattedSearch.Action = Action;
            FormattedSearch.ForceRefresh = ForceRefresh;
            FormattedSearch.ByField = ByField;
            FormattedSearch.Refresh = Refresh;
            FormattedSearch.FieldID = FieldID;
            FormattedSearch.QueryID = IntrnalKey;

            RetVal = FormattedSearch.Add();

            if (RetVal != 0)
            {
                oCompany.GetLastError(out ErrCode, out ErrMsg);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(FormattedSearch);
                throw new Exception("Erro criando Consulta Formatada no SAP: " + ErrMsg);
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(FormattedSearch);
            GC.Collect();
        }

        // Clears table
        public void ClearTable(SAPbobsCOM.Company oCompany, string TableName)
        {

            string Query = "";

            SAPbobsCOM.Recordset oRecSet = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (oCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                Query = "delete from \"" + TableName + "\"";
            else
                Query = "delete from [" + TableName + "]";

            oRecSet.DoQuery(Query);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecSet);
            GC.Collect();
        }

        // Check whether a UDO already exists
        public bool UDOExist(SAPbobsCOM.Company oCompany, string UDOCode)
        {

            SAPbobsCOM.UserObjectsMD UserObjectMD = (SAPbobsCOM.UserObjectsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            if (!UserObjectMD.GetByKey(UDOCode))
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserObjectMD);
                UserObjectMD = null;
                GC.Collect();
                return false;
            }
            else
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(UserObjectMD);
                UserObjectMD = null;
                GC.Collect();
                return true;
            }

        }

        public bool FieldExist(Company company, string fldName)
        {

            Field oUserFields = null;
            SAPbobsCOM.Documents doc = (SAPbobsCOM.Documents)company.GetBusinessObject(BoObjectTypes.oOrders);

            int count;
            count = doc.UserFields.Fields.Count;
            for (int i = 0; i < count; i++)
            {
                oUserFields = doc.UserFields.Fields.Item(i);
                if (oUserFields.Name == fldName)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFields);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                    oUserFields = null;
                    doc = null;
                    GC.Collect();
                    return true;
                }
            }

            if (oUserFields != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFields);
                oUserFields = null;
                GC.Collect();
            }

            if (doc != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                oUserFields = null;
                GC.Collect();
            }

            return false;
        }

        public bool FieldExist(Company company, string fldName, string tableName)
        {
            Recordset rs = (Recordset)company.GetBusinessObject(BoObjectTypes.BoRecordset);
            String sql = "";
            if (company.DbServerType == BoDataServerTypes.dst_HANADB)
                sql = String.Format("select top 1 \"{0}\" from \"{1}\"", fldName, tableName);
            else
                sql = String.Format("select top 1 {0} from [{1}]", fldName, tableName);
            try
            {
                rs.DoQuery(sql);
            }
            catch
            {
                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;
                GC.Collect();
            }
            return true;
        }
    }
}
