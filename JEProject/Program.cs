using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JEProject
{
    static class Program
    {
        public static SAPbobsCOM.Company oCompany;
        static void Main()
        {
            try
            {
                // DI Sap Connection
                oCompany = Classes.BusinessOne.Connections.SapConnection();

                // Update all JE that have lines with NULL Projects and 1 line at least with a not null Project with that one.
                Classes.BusinessOne.JournalEntry.Update(oCompany);

                //MessageBox.Show("The procedure completed.");
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                System.Environment.Exit(0);
            }
        }
    }
}

