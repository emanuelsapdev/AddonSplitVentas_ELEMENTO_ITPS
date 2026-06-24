using Addon_AutoDivSalesOrd.Common;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Services
{
    public class InfraDataService
    {
        public class ValidValueOption
        {
            public string Value { get; set; }
            public string Description { get; set; }
        }

        public static void CreateUserTable(string name, string desc, BoUTBTableType type)
        {
            try
            {

                var oTableMd = (UserTablesMD)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oUserTables);
                var exists = oTableMd.GetByKey(name);
                if (!exists)
                {
                    oTableMd.TableName = name;  // sin @
                    oTableMd.TableDescription = desc;
                    oTableMd.TableType = type;

                    int ret = oTableMd.Add();
                    if (ret != 0)
                    {
                        string err = ConnectionSDK.DIAPI.GetLastErrorDescription();
                        throw new Exception($"Error creando tabla {name}: {err}");
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTableMd);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static void CreateUserField(string tableName,
                                   string fieldName,
                                   string desc,
                                   BoFieldTypes type,
                                   int size = 0,
                                   BoFldSubTypes subType = BoFldSubTypes.st_None,
                                   string linkedTable = null,
                                   string linkedUDO = null,
                                   UDFLinkedSystemObjectTypesEnum linkedSystemObject = UDFLinkedSystemObjectTypesEnum.ulNone,
                                   IEnumerable<ValidValueOption> validValues = null,string defaultValue = null)
        {
            if (FieldExists(tableName, fieldName))
                return;

            var oFieldMd = (UserFieldsMD)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.oUserFields);

            try
            {
                oFieldMd.TableName   = tableName;    // sin @ para tablas UDO / estándar
                oFieldMd.Name        = fieldName;    // sin U_
                oFieldMd.Description = desc;
                oFieldMd.Type        = type;
                oFieldMd.SubType     = subType;

                if (size > 0)
                    oFieldMd.EditSize = size;

                if (!string.IsNullOrEmpty(linkedTable))
                {
                    oFieldMd.LinkedTable = linkedTable;
                }
                else if (!string.IsNullOrEmpty(linkedUDO))
                {
                    oFieldMd.LinkedUDO = linkedUDO;
                }
                else if (linkedSystemObject != UDFLinkedSystemObjectTypesEnum.ulNone)
                {
                    oFieldMd.LinkedSystemObject = linkedSystemObject;
                }

                if (!string.IsNullOrEmpty(oFieldMd.DefaultValue)) {                     
                    oFieldMd.DefaultValue = defaultValue;
                }

                if (validValues != null)
                {
                    var values = validValues
                        .Where(v => v != null && !string.IsNullOrWhiteSpace(v.Value))
                        .ToList();

                    for (int i = 0; i < values.Count; i++)
                    {
                        if (i > 0)
                            oFieldMd.ValidValues.Add();

                        oFieldMd.ValidValues.SetCurrentLine(i);
                        oFieldMd.ValidValues.Value = values[i].Value;
                        oFieldMd.ValidValues.Description = string.IsNullOrWhiteSpace(values[i].Description)
                            ? values[i].Value
                            : values[i].Description;
                    }
                }

                

                int ret = oFieldMd.Add();
                if (ret != 0)
                    throw new Exception($"Error creando campo {tableName}.U_{fieldName}: {ConnectionSDK.DIAPI.GetLastErrorDescription()}");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFieldMd);
            }
        }

        private static bool FieldExists(string tableName, string fieldName)
        {
            var rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
            rs.DoQuery($@"
            SELECT 1 
              FROM CUFD 
             WHERE ""TableID"" = '{tableName.Replace("'", "''")}'  
               AND ""AliasID"" = '{fieldName.Replace("'", "''")}'");

            bool exists = !rs.EoF;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            return exists;
        }

        public static void CreateViewIfNotExists(string viewName, string viewBody)
        {
            var rs = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
            try
            {
                rs.DoQuery($@"SELECT 1 FROM SYS.VIEWS WHERE ""VIEW_NAME"" = '{viewName}' AND ""SCHEMA_NAME"" = CURRENT_SCHEMA");    

                if (!rs.EoF)
                    return;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                rs = null;

                var rsCreate = (Recordset)ConnectionSDK.DIAPI.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    rsCreate.DoQuery($"CREATE VIEW \"{viewName}\" AS {viewBody}");
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCreate);
                }
            }
            finally
            {
                if (rs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
    }
}
