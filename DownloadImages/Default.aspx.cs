using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Excel = Microsoft.Office.Interop.Excel;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"c:\a.xlsx");
        Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
        Excel.Range xlRange = xlWorksheet.UsedRange;

        //iterate over the rows and columns and print to the console as it appears in the file
        //excel is not zero based!!
        for (int i = 2; i <= 10; i++)
        {
            if (((Excel.Range)xlRange.Cells[i, 1]).Value2 != null && ((Excel.Range)xlRange.Cells[i, 2]).Value2 != null)
            {
                string localFileName = @"C:\Images\" + ((Excel.Range)xlRange.Cells[i, 1]).Value2.ToString();
                string remoteFileUrl = ((Excel.Range)xlRange.Cells[i, 2]).Value2.ToString();
                string extnOfThefile = remoteFileUrl.Split(".".ToCharArray()).Last();

                WebClient webClient = new WebClient();
                webClient.DownloadFile(remoteFileUrl, localFileName + "." + extnOfThefile);
                ((Excel.Range)xlRange.Cells[i, 1]).Value = "Done";
            }
            //write the value to the console


            //add useful things here!   
        }
        xlWorkbook.Save();
        xlWorkbook.Close();
        
        }
    }
