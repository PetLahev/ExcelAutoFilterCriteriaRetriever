using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace AutoFilterRetriever
{
    public partial class AddinRibbon
    {
        private void AddinRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnGetFilterCriteria_Click(object sender, RibbonControlEventArgs e)
        {
            var sheet = Globals.ThisAddIn.Application.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            if (sheet == null) return;

            CriteriaFilterRetriever retriever = null;
            try
            {
                retriever = new CriteriaFilterRetriever(sheet);
                retriever.GetFilterCriteria();

                if (retriever.FilterCriteria != null)
                {
                    var message = string.Join(Environment.NewLine, retriever.FilterCriteria);
                    System.Windows.Forms.MessageBox.Show(message);
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error occurred");
            }
            finally
            {
                retriever?.Dispose();
            }
            

        }
    }
}
