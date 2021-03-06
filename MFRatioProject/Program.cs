﻿using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO;
using System.Reflection;

namespace MFRatioProject
{
    class Program
    {
        static void Main(string[] args)
        {
            var html = new HtmlDocument();
            Application xlApp = new Application();
            var LoadEntireData = false;
            if (xlApp == null)
            {
                Console.WriteLine("not installed");
                return;
            }
            object misValue = System.Reflection.Missing.Value;
            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            //date
            //exit policy
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
            xlWorkSheet.Cells[1, 14] = "Ratio Date";
            xlWorkSheet.Cells[1, 15] = "Exit Load";
            xlWorkSheet.Cells[1, 16] = "Start Date";
            if (LoadEntireData)
            {
                xlWorkSheet.Cells[1, 17] = "PortFolio name";
                xlWorkSheet.Cells[1, 18] = "Sector";
                xlWorkSheet.Cells[1, 19] = "PE";
                xlWorkSheet.Cells[1, 20] = "%Assets";
            }
            var funds = new Dictionary<string, string>();

            XmlDataDocument xmldoc = new XmlDataDocument();
            XmlNodeList xmlnode;

            var fileName2 = Path.Combine(
           Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)
               , @"myfunds.xml");
            FileStream fs = new FileStream(fileName2, FileMode.Open, FileAccess.Read);
            xmldoc.Load(fs);
            xmlnode = xmldoc.GetElementsByTagName("mf");
            var baseUrl = "https://www.valueresearchonline.com/funds/fundperformance.asp?schemecode=";
            for (int mf = 0; mf < xmlnode.Count; mf++)
            {
                funds.Add(mf.ToString(), baseUrl + xmlnode[mf].ChildNodes.Item(0).InnerText.Trim());
            }
            fs.Close();

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
            "/html/body/div[2]/div/div/div[1]/div[1]/h1/span[2]/img",
      "/html/body/div[2]/div/div/div[1]/div[9]/div"};


            var list = new List<string>() {
            "/td[2]/a",
            "/td[3]/a",
            "/td[4]",
            "/td[7]"};

            int columnNo = 1, rowNo = 2;
            foreach (var fund in funds.Values)
            {
                html.LoadHtml(new WebClient().DownloadString(fund));

                foreach (var column in columnList)
                {
                    var node = column;
                    var a = html.DocumentNode.SelectNodes(node);
                    string result = "";
                    if (columnNo == 13)
                    {
                        result = a.FirstOrDefault() != null ? a.FirstOrDefault().Attributes.FirstOrDefault().Value.Split('/').Last() : "";
                    }
                    else if (columnNo != 1 && columnNo != 2 && columnNo != 14)
                    {
                        result = a.FirstOrDefault() != null ? a.FirstOrDefault().InnerText.Replace("\r\n ", "").Replace("R", "").Trim() : "";
                    }
                    else if (columnNo == 14)
                    {
                        result = a != null ? a.FirstOrDefault().InnerText.Replace("\r\n ", "").Split('.')[0] : "";

                    }
                    else
                    {
                        result = a.FirstOrDefault() != null ? a.FirstOrDefault().InnerText.Replace("\r\n ", "") : "";
                    }
                    xlWorkSheet.Cells[rowNo, columnNo++] = result;
                }
                var newsFundsPage = fund.Replace("fundperformance", "newsnapshot");
                html.LoadHtml(new WebClient().DownloadString(newsFundsPage));
                var ExitPoint = "//*[@id='super-container']/div[2]/div/div/div[1]/div[7]/table[2]/tr[9]/td[2]";
                var r = html.DocumentNode.SelectNodes(ExitPoint);
                xlWorkSheet.Cells[rowNo, columnNo++] = r != null ? r.FirstOrDefault().InnerText.Replace("\r\n ", "").Trim() : "";
                var fundStart = "//*[@id='super-container']/div[2]/div/div/div[1]/div[7]/table[1]/tr[3]/td[2]";
                r = html.DocumentNode.SelectNodes(fundStart);
                xlWorkSheet.Cells[rowNo, columnNo++] = r != null ? r.FirstOrDefault().InnerText.Replace("\r\n ", "").Trim() : "";
                rowNo++;
                if (LoadEntireData)
                {
                    var replaceFundName = fund.Replace("fundperformance", "portfoliovr");
                    html.LoadHtml(new WebClient().DownloadString(replaceFundName));
                    var baseURL = "//*[@id='fund-snapshot-port-holdings']/tr[3]/";
                    for (var i = 3; true; i++)
                    {
                        Regex rx = new Regex(@"\d");
                        var rs = html.DocumentNode.SelectNodes(rx.Replace(baseURL, i.ToString()) + "/td[2]/a");
                        if (rs == null)
                        {
                            break;
                        }
                        foreach (var l in list)
                        {
                            string pattern = @"\d";
                            string replacement = i.ToString();
                            Regex rgx = new Regex(pattern);
                            string result2 = rgx.Replace(baseURL, replacement);
                            var res = html.DocumentNode.SelectNodes(result2 + l);
                            xlWorkSheet.Cells[rowNo, columnNo++] = res.FirstOrDefault().InnerText;
                        }
                        rowNo++;
                        columnNo = 17;
                    }
                    rowNo++; columnNo = 1;
                }
                else
                {
                    columnNo = 1;
                }
            }

            var fileName = "MF Comp " + DateTime.Now.ToShortDateString() + ".xls";
            xlWorkBook.SaveAs("d:\\" + fileName.Replace('/', '-'), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
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
