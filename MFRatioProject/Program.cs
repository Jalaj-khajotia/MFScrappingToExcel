using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace MFRatioProject
{
    class Program
    {
        static void Main(string[] args)
        {
            var html = new HtmlDocument();
              Application xlApp = new Application();
            if (xlApp == null)
            {
                Console.WriteLine("not installed");
                return;
            }
            object misValue = System.Reflection.Missing.Value;
            var xlWorkBook = xlApp.Workbooks.Add(misValue);

            var xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Fund Name";
            xlWorkSheet.Cells[1, 2] = "Category";
            xlWorkSheet.Cells[1, 3] = "Total Asset";
            xlWorkSheet.Cells[1, 4] = "1 Yr Return";
            xlWorkSheet.Cells[1, 5] = "3 Yr Return";
            xlWorkSheet.Cells[1, 6] = "5 Yr Return";
            xlWorkSheet.Cells[1, 7] = "Expense Ratio";
            xlWorkSheet.Cells[1, 8] = "SD";
            xlWorkSheet.Cells[1, 9] = "Sharpe";
            xlWorkSheet.Cells[1, 10] = "Sortino";
            xlWorkSheet.Cells[1, 11] = "Beta";
            xlWorkSheet.Cells[1, 12] = "Alpha";
            xlWorkSheet.Cells[1, 13] = "VRO Rating";
            var fundList = new List<string>();
            var funds = new Dictionary<string, string>();
            
            funds.Add("DSP BlackRock Small and Mid Cap Fund - Regular Plan",
                "https://www.valueresearchonline.com/funds/fundperformance.asp?schemecode=3725");
            funds.Add("SBI Emerging Businesses Fund", "https://www.valueresearchonline.com/funds/fundperformance.asp?schemecode=2415");
            funds.Add("Mirae Asset Emerging Bluechip Fund - Regular Plan", "https://www.valueresearchonline.com/funds/fundperformance.asp?schemecode=11213");
            ////funds.Add("", "");

            var columnList = new List<string>() { "/html/body/div[2]/div/div/div[1]/div[1]/h1/span[1]/span", 
            "//*[@id='fundHead']/div[3]/div/table/tr[1]/td[2]/a",
                "//*[@id='fundHead']/div[3]/div/table/tr[2]/td[2]",
                "/html/body/div[2]/div/div/div[1]/div[10]/table/tr[2]/td[8]",
            "/html/body/div[2]/div/div/div[1]/div[10]/table/tr[2]/td[9]",
                "/html/body/div[2]/div/div/div[1]/div[10]/table/tr[2]/td[10]",
            "//*[@id='fundHead']/div[3]/div/table/tr[3]/td[2]",
                "/html/body/div[2]/div/div/div[1]/div[9]/table/tr[2]/td[3]",
                "/html/body/div[2]/div/div/div[1]/div[9]/table/tr[2]/td[4]",
            "/html/body/div[2]/div/div/div[1]/div[9]/table/tr[2]/td[5]",
                "/html/body/div[2]/div/div/div[1]/div[9]/table/tr[2]/td[6]",
             "/html/body/div[2]/div/div/div[1]/div[9]/table/tr[2]/td[7]",
            "/html/body/div[2]/div/div/div[1]/div[1]/h1/span[2]/img"};

            int columnNo = 1, rowNo=2;
            foreach (var fund in funds.Values)
            {
                html.LoadHtml(new WebClient().DownloadString(fund));

                foreach(var column in columnList)
                {
                    var node = column;
                    var a = html.DocumentNode.SelectNodes(node);
                    string result;
                    if(columnNo == 13)
                    {
                        result = a.FirstOrDefault()!=null? a.FirstOrDefault().Attributes.FirstOrDefault().Value.Split('/').Last(): "";
                    }
                    else
                    {
                        result = a.FirstOrDefault()!= null?a.FirstOrDefault().InnerText.Replace("\r\n ", "").Replace("R", "").Trim():"";
                    }
                    
                    xlWorkSheet.Cells[rowNo, columnNo++] = result;
                }
                rowNo++;columnNo = 1;

            }

            
            //var root = html.DocumentNode;


          


            xlWorkBook.SaveAs("d:\\Mutual Fund Comparision.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
             misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
              misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
