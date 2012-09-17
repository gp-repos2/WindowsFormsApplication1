using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Reflection;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Owc11 = Microsoft.Office.Interop.Owc11;

namespace WindowsFormsApplication1
{
 
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlBook;
            Excel.Worksheet xlSheet1, xlSheet2, xlSheet3;

            xlApp = new Excel.Application();
            xlBook = xlApp.Workbooks.Open("C:\\Documents and Settings\\PGustov\\Desktop\\Книга1.xls", 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //xlApp.Visible = true;
            //xlApp.Parent.Windows[1].Visible = true;



            while (axSpreadsheet1.Sheets.Count > xlBook.Worksheets.Count)
            {
//                Owc11.Worksheet owcSheet2 = (Owc11.Worksheet)axSpreadsheet1.Sheets[axSpreadsheet1.Sheets.Count];
//                MessageBox.Show(axSpreadsheet1.ActiveSheet.Name);
                try
                {
                    axSpreadsheet1.ActiveSheet.Delete();
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.ToString());
                }
            }

                       
            if (xlBook.Worksheets.Count > 3)
            {
                axSpreadsheet1.Sheets.Add(Type.Missing, Type.Missing, xlBook.Worksheets.Count - 3, Type.Missing);
            }

            int i=0;
            Owc11.Worksheet owcSheet;

            foreach (Excel.Worksheet xlSheet in xlBook.Worksheets)
            {
                i++;
                owcSheet = (Owc11.Worksheet)axSpreadsheet1.Sheets[i];
                owcSheet.Name = "szxcxcx" + i.ToString();
            }

            i = 0;
            foreach (Excel.Worksheet xlSheet in xlBook.Worksheets)
            {
                //MessageBox.Show(xlSheet.Name);

                i++;
                 xlSheet.Activate();
                 xlSheet.Cells.Select();
                 xlSheet.Cells.Copy(Type.Missing);

                 //owcSheet = (Owc11.Worksheet)axSpreadsheet1.Sheets.Add(axSpreadsheet1.Sheets[1], null, 1, Type.Missing);
                 
                 owcSheet=(Owc11.Worksheet)axSpreadsheet1.Sheets[i];

                 owcSheet.Name = xlSheet.Name;
                 owcSheet.Activate();
                 owcSheet.Cells.Select();
                 owcSheet.Cells.Paste();
            }


            xlApp.DisplayAlerts = false;
            xlBook.Close(Type.Missing, Type.Missing, Type.Missing);          
            xlApp.Quit();
            
            
/*            if (axSpreadsheet1.Sheets.Count > 1)
            {
                try
                {
                    ((Owc11.Worksheet)axSpreadsheet1.Sheets[1]).Delete();
                }
                catch (Exception eee)
                {
                    MessageBox.Show(eee.ToString());
                }
            }
*/


/*
Dim objexcel As Object
    ' Dim objworkbook As Object
    Dim objsheet As Object

    ' Dim objExcel As New Excel.ApplicationClass
    ' Dim objSheet As Excel.Worksheet
    Try
      'Create and open the correct excel workbook
      objExcel = CreateObject("Excel.Application")
      objexcel.Workbooks.Open("C:\Central systems\Central systems\Folders\Department Log\Reports\Queries Report 2010.xls")
      objExcel.Visible = True
      objExcel.Parent.Windows(1).Visible = True
    Catch ex As Exception
      MessageBox.Show(ex.Message)
    End Try

    'Select all the cells in the relevant worksheet
    'and copy the contents
    objsheet = objexcel.ActiveSheet
    objsheet = objexcel.Sheet("Jan" & Now.Year)
    objSheet.Cells.Select()
    objSheet.Cells.Copy()

    'Select a cell in the AxSpreadsheet control that is 
    'already declared and initialised in the class
    Me.AxSpreadsheet1.Cells.Range("A1:Z5000").Select()

    'Paste the contents
    Me.AxSpreadsheet1.Cells.Range("A1").Paste()
    Me.AxSpreadsheet1.Sheets.Add(12)
*/
        }
    }
}
