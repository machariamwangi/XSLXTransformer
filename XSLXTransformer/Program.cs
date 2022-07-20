using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace XSLXTransformer
{

    class Program
    {
        static string mypath = "url-here";
        static string fileName = "file name(with extension here";
        static string fileextension = "*.extension";
        static void Main(string[] args)
        {
            try
            {
                var d = new DirectoryInfo(mypath);
                var files = d.GetFiles(fileextension);


                foreach (var file in files)
                {
                    var fileName = file.FullName;
                    using var package = new ExcelPackage(file, true);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (ExcelPackage p = new ExcelPackage(file))
                    {
                        var currentSheet = p.Workbook.Worksheets;
                        //var workSheet = currentSheet.First();
                        var workSheet = currentSheet.ElementAt(2);
                        OfficeOpenXml.ExcelWorksheet ws = p.Workbook.Worksheets.Add(workSheet.Name + "_English");

                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        //for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                        for (int rowIterator = 2; rowIterator <= 10; rowIterator++)
                        {
                            Console.WriteLine("=============================START===========================================");
                            ws.Cells[rowIterator, 1].Value = Translate(workSheet.Cells[rowIterator, 1].Value?.ToString());
                            ws.Cells[rowIterator, 2].Value = Translate(workSheet.Cells[rowIterator, 2].Value?.ToString());
                            ws.Cells[rowIterator, 3].Value = Translate(workSheet.Cells[rowIterator, 3].Value?.ToString());
                            ws.Cells[rowIterator, 4].Value = Translate(workSheet.Cells[rowIterator, 4].Value?.ToString());
                            ws.Cells[rowIterator, 5].Value = Translate(workSheet.Cells[rowIterator, 5].Value?.ToString());

                            Console.WriteLine("============================FINISH============================================");
                          

                        }
                    
                      p.SaveAs(new FileInfo(mypath+ file));
                    }
                }


            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
            static String Translate(String word)
            {
                var toLanguage = "en";//English
                var fromLanguage = "de";//Deutsch
                var url = $"https://translate.googleapis.com/translate_a/single?client=gtx&sl={fromLanguage}&tl={toLanguage}&dt=t&q={HttpUtility.UrlEncode(word)}";
                var webClient = new WebClient
                {
                    Encoding = Encoding.UTF8
                };
                var result = webClient.DownloadString(url);
                try
                {
                    result = result.Substring(4, result.IndexOf("\"", 4, StringComparison.Ordinal) - 4);
                    return result;
                }
                catch
                {
                    return "Error";
                }
            }
        }

    }


}
