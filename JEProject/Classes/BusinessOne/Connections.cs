using System;
using System.Windows.Forms;
using System.Diagnostics;
using SAPbobsCOM;

namespace JEProject.Classes.BusinessOne
{
    class Connections
    {
        //Sap Connection
        public static SAPbobsCOM.Company oCompany;
        public static SAPbobsCOM.Company SapConnection()
        {
            try
            {
                int ResCode = 0;
                string ErrMsg = null;

                Connections.AppRunning();

                oCompany = new SAPbobsCOM.Company();
                oCompany.Server = Properties.Settings.Default.ServerName;
                oCompany.DbServerType = DataBaseType();
                oCompany.DbUserName = Properties.Settings.Default.SQLUserName;
                oCompany.DbPassword = Properties.Settings.Default.SQLPassword;
                oCompany.CompanyDB = (string)Environment.GetCommandLineArgs().GetValue(1);
                oCompany.UserName = Properties.Settings.Default.SAPUserName;
                oCompany.Password = Properties.Settings.Default.SAPPassword;

                ResCode = oCompany.Connect();

                if (ResCode != 0)
                {
                    oCompany.GetLastError(out ResCode, out ErrMsg);
                    //MessageBox.Show($@"{ResCode}: {ErrMsg}");
                    //System.Windows.Forms.Application.Exit();
                    System.Environment.Exit(0);
                }
                else
                {
                    //MessageBox.Show($@"Connected successfully to DB: {oCompany.CompanyDB}");
                    return oCompany;
                    //Console.Write("Connected successfully to DB: {0}", oCompany.CompanyDB);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return oCompany;
        }

        public static BoDataServerTypes DataBaseType()
        {
            string sql_server = Properties.Settings.Default.SQLVersion;

            if (!sql_server.Contains("MSSQL"))
                throw new Exception($@"Unsupported database type of ${Properties.Settings.Default.SQLVersion}");

            if (BoDataServerTypes.dst_MSSQL2005.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2005;

            if (BoDataServerTypes.dst_MSSQL2008.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2008;

            if (BoDataServerTypes.dst_MSSQL2012.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2012;

            if (BoDataServerTypes.dst_MSSQL2014.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2014;

            if (BoDataServerTypes.dst_MSSQL2016.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2016;

            if (BoDataServerTypes.dst_MSSQL2017.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2017;

            if (BoDataServerTypes.dst_MSSQL2019.ToString().Contains(sql_server))
                return BoDataServerTypes.dst_MSSQL2019;

            throw new Exception($@"Database type: ${sql_server} is not supported.");

        }

        //Check if the process is already running. If yes, exit.
        private static void AppRunning()
        {
            try
            {
                int intCounter = 0;
                foreach (Process proc in Process.GetProcesses())
                {
                    if (proc.ProcessName.StartsWith(Application.ProductName))
                        intCounter += 1;
                    if (intCounter > 1)
                    {
                        MessageBox.Show($@"Another instance of {Application.ProductName} is running. Please try again later.");
                        System.Environment.Exit(0);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
