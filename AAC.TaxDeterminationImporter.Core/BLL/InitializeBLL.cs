using SBO.Hub;
using SAPbobsCOM;
using AAC.TaxDeterminationImporter.Core.DAO;
using System.Runtime.InteropServices;

namespace AAC.TaxDeterminationImporter.Core.BLL
{
    public class InitializeBLL
    {
        public static void Initialize()
        {
            Scripts.SetResourceManager();
            SetupDatabaseFunctions();
            EventFilterBLL.SetDefaultEvents();
        }

        private static void SetupDatabaseFunctions()
        {
            if (SBOApp.Company.DbServerType != BoDataServerTypes.dst_HANADB)
                return;

            ExecuteDDL(HanaFunctions.FN_TAXCODE_EXPORT);
            ExecuteDDL(HanaFunctions.FN_GET_KEYFIELDDESCRIPTION);
        }

        private static void ExecuteDDL(string sql)
        {
            var rs = SBOApp.Company.GetBusinessObject(BoObjectTypes.BoRecordset) as Recordset;
            try
            {
                rs.DoQuery(sql);
            }
            finally
            {
                Marshal.ReleaseComObject(rs);
            }
        }
    }
}
