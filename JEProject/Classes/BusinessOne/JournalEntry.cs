using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using System.Windows.Forms;

namespace JEProject.Classes.BusinessOne
{
    class JournalEntry
    {
        private static int ResCode = 0;
        private static int ErrCode = 0;
        private static string ErrMsg = null;

        public static void Update(SAPbobsCOM.Company _Company)
        {
            try
            {
                //All JE that have lines with NULL Projects and 1 at least with Project
                string JE = $@" SELECT T1.TransId 
                                FROM OJDT T0
                                INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId
                                WHERE ISNULL(T1.Project, '') = ''
                                AND T1.TransId IN ( SELECT DISTINCT S1.TransId
					                                FROM JDT1 S1
					                                WHERE ISNULL(S1.Project, '') != '')
                                GROUP BY T1.TransId ";

                SAPbobsCOM.Recordset rs;
                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(JE);
                rs.MoveFirst();

                while (rs.EoF == false)
                {
                    int TransId = 0;
                    TransId = int.Parse(rs.Fields.Item("TransId").Value.ToString());

                    try
                    {
                        SAPbobsCOM.JournalEntries oJE;
                        oJE = (SAPbobsCOM.JournalEntries)_Company.GetBusinessObject(BoObjectTypes.oJournalEntries);

                        // get Journal Entry
                        if (oJE.GetByKey(TransId) == true)
                        {
                            //All null lines for a specific TransId that whould be updated with the first not null project in JE
                            string JE_lines = $@"   SELECT T1.TransId, T1.Line_ID, T1.Project AS Null_Project, T2.Project
                                                    FROM OJDT T0
                                                    INNER JOIN JDT1 T1 ON T0.TransId = T1.TransId
                                                    INNER JOIN (SELECT TOP 1 S1.TransId, S1.Project --298
			                                                    FROM JDT1 S1
			                                                    WHERE ISNULL(S1.Project, '') != ''
			                                                    AND S1.TransId = {TransId}) T2 ON T1.TransId = T2.TransId
                                                    WHERE ISNULL(T1.Project, '') = '' ";

                            SAPbobsCOM.Recordset rsLines;
                            rsLines = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            rsLines.DoQuery(JE_lines);
                            rsLines.MoveFirst();

                            while (rsLines.EoF == false)
                            {
                                int LineID = -1;
                                LineID = int.Parse(rsLines.Fields.Item("Line_ID").Value.ToString());
                                oJE.Lines.SetCurrentLine(LineID);
                                oJE.Lines.ProjectCode = rsLines.Fields.Item("Project").Value.ToString();

                                rsLines.MoveNext();
                            }
                            rsLines = null;

                            ResCode = oJE.Update();
                            if (ResCode != 0)
                            {
                                _Company.GetLastError(out ErrCode, out ErrMsg);
                                string Delete = $@" DELETE FROM [dbo].[DDS_LOG]
                                                           WHERE Id = {TransId}
                                                           AND Addon = N'{Application.ProductName}'   ";
                                SAPbobsCOM.Recordset rsDelete;
                                rsDelete = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsDelete.DoQuery(Delete);
                                rsDelete = null;

                                string Log = $@"INSERT INTO [dbo].[DDS_LOG]
                                                               ([Id]
                                                               ,[Exception]
                                                               ,[Addon])
                                                         VALUES
                                                               ({TransId}
                                                               ,'{ErrCode} {ErrMsg.Replace("'", "")}'
                                                               ,'{Application.ProductName}')    ";
                                SAPbobsCOM.Recordset rsLog;
                                rsLog = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsLog.DoQuery(Log);
                                rsLog = null;
                            }
                            else
                            {
                                string Delete = $@" DELETE FROM [dbo].[DDS_LOG]
                                                           WHERE Id = {TransId}
                                                           AND Addon = N'{Application.ProductName}'  ";
                                SAPbobsCOM.Recordset rsDelete;
                                rsDelete = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsDelete.DoQuery(Delete);
                                rsDelete = null;

                                string Log = $@"INSERT INTO [dbo].[DDS_LOG]
                                                               ([Id]
                                                               ,[Exception]
                                                               ,[Addon])
                                                         VALUES
                                                               ({TransId}
                                                               ,'success'
                                                               ,'{Application.ProductName}')    ";
                                SAPbobsCOM.Recordset rsLog;
                                rsLog = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsLog.DoQuery(Log);
                                rsLog = null;
                            }
                        }
                    }
                    catch (Exception ex) { }

                    rs.MoveNext();
                }
                rs = null;
            }
            catch (Exception e)
            { }
        }
    }
}
