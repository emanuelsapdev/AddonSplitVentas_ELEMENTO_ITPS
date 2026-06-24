using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Windows.Forms;

namespace Addon_AutoDivSalesOrd.Common
{
    public class ConnectionSDK
    {
        protected static SAPbouiCOM.Application _UIAPI;
        protected static SAPbobsCOM.Company _DIAPI;
        public static SAPbouiCOM.Application UIAPI => _UIAPI ?? throw new Exception(Constants.Messages.UiApiNotDefined);
        public static SAPbobsCOM.Company DIAPI => _DIAPI ?? throw new Exception(Constants.Messages.DiApiNotDefined);

        public static bool Connected => _UIAPI != null && _DIAPI.Connected;


        public static void Singlenton()
        {
            _UIAPI = GetApplication();

            if (_UIAPI != null)
            {
                _DIAPI = _UIAPI.Company.GetDICompany();

            }
        }



        public static SAPbouiCOM.Application GetApplication()
        {
            SboGuiApi api = new SboGuiApi()
            {
                AddonIdentifier = Constants.Addon.Identifier
            };

            string[] commands = Environment.GetCommandLineArgs();
            string strConnection = null;

            // SAP B1 passes the connection string as the first argument after the executable.
            // Environment.GetCommandLineArgs(): [0]=exe path, [1]=connection string (when started by SAP).
            if (commands != null && commands.Length > 1)
            {
                strConnection = commands[1];
            }

            if (string.IsNullOrWhiteSpace(strConnection))
            {
                throw new ArgumentException(Constants.Messages.MissingSapConnectionString, nameof(strConnection));
            }

            try
            {
                api.Connect(strConnection);
            }
            catch (Exception ex)
            {
                throw new Exception(Constants.Messages.UiApiConnectErrorPrefix + ex.Message, ex);
            }

            SAPbouiCOM.Application app;
            try
            {
                app = api.GetApplication();
            }
            catch (Exception ex)
            {
                throw new Exception(Constants.Messages.UiApiGetApplicationErrorPrefix + ex.Message, ex);
            }

            if (app == null)
                throw new Exception(Constants.Messages.UiApiApplicationNull);

            return app;
        }
    }
}
