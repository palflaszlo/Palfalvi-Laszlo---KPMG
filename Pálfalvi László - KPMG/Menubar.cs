using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

//------< using >------
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
//------</ using >------


namespace Pálfalvi_László___KPMG
{
    public partial class Menubar
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void MNB_adatletoltes_Click(object sender, RibbonControlEventArgs e)
        {
            //< open Excel Worksheet >
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //< open Excel Worksheet >

            //< get Cell.Value > 
            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            if(actCell.Value2 != null)
            {
                string sValue = actCell.Value2.ToString();
                string sText  = actCell.Text;
                System.Windows.Forms.MessageBox.Show(sText);
            }

            //</ get Cell.Value > 

        }
    }
}
