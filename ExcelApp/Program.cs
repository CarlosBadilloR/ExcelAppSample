using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {

            List<string> tcList = new List<string>();
            var newFile = new System.IO.FileInfo("C://workbooks//sample7.xlsx");
            using (ExcelPackage package = new ExcelPackage())
            {

                var trac = package.Workbook.Worksheets.Add("Traceability");
             /*   var tcSheet = package.Workbook.Worksheets.Add("Test Cases");

    */
                string[] req = new string[20];
                for (int i = 1; i < req.Length; i++)
                {
                    req[i] = "Req_" + i.ToString();
                    
                }

                string[] tc = new string[5];
                for (int i = 1; i < tc.Length; i++)
                {
                    tc[i] = "TC" + i.ToString();

                }

                string[] steps = new string[10];
                for (int i = 1; i < steps.Length; i++)
                {

                    steps[i] = "Step " + i.ToString();
                }

                int tcRowLoop = 2;
                /*
                foreach (var test in tc)
                {
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Merge = true;
                    tcSheet.SetValue(tcRowLoop, 1, "Library");
                    tcRowLoop++;
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":C" + tcRowLoop.ToString()].Merge = true;
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":C" + tcRowLoop.ToString()].Value = "Name";
                    tcSheet.Cells["D" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Merge = true;
                    tcSheet.Cells["D" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Value = test;
                    tcList.Add("#'Test Cases'!$D$"+tcRowLoop.ToString());


                    tcRowLoop++;
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":C" + tcRowLoop.ToString()].Merge = true;
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":C" + tcRowLoop.ToString()].Value = "#";
                    tcSheet.Cells["D" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Merge = true;
                    tcSheet.Cells["D" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Value = "Actions";
                    tcRowLoop++;
                    int aux = 1;
                    foreach (var step in steps)
                    {
                        tcSheet.Cells["A" + tcRowLoop.ToString() + ":C" + tcRowLoop.ToString()].Merge = true;
                        tcSheet.Cells["A" + tcRowLoop.ToString() + ":C" + tcRowLoop.ToString()].Value = aux.ToString();
                        tcSheet.Cells["D" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Merge = true;
                        tcSheet.Cells["D" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Value = step;
                        tcRowLoop++;
                        aux++;

                    }
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Merge = true;
                    tcSheet.Cells["A" + tcRowLoop.ToString() + ":F" + tcRowLoop.ToString()].Value = "";
                    tcRowLoop++;
                }


                int reqRow = 2;
                foreach (var item in req)
                {
                    var aux = reqRow-1 + tc.Length;
                     ExcelRange Rng = trac.Cells["A" + reqRow.ToString() + ":A" + aux.ToString()];
                    Rng.Hyperlink = new Uri("#'Test Cases'!$D$17", UriKind.Relative);
                    Rng.Value = "Test";
                    trac.Cells["A" + reqRow.ToString() + ":A" + aux.ToString()].Merge = true;
                    trac.Cells["B" + reqRow.ToString() + ":B" + aux.ToString()].Merge = true;
                    trac.Cells["C" + reqRow.ToString() + ":C" + aux.ToString()].Merge = true;
                    trac.SetValue(reqRow, 1, item);
                    trac.SetValue(reqRow, 2, "user module: "+reqRow.ToString());
                    trac.SetValue(reqRow, 3, "user submodule: " + reqRow.ToString());
                    var tcRow = reqRow;
                    foreach (var test in tc)
                    {
                        trac.SetValue(tcRow, 4, test);
                        ExcelRange RngT = trac.Cells["D" + reqRow.ToString() + ":D" + reqRow.ToString()];
                        foreach (var listItem in tcList)
                        {
                            if(test.Equals(listItem))
                            RngT.Hyperlink = new Uri("#'Test Cases'!$D$17", UriKind.Relative);
                        }

                        tcRow++;
                        reqRow++;
                    }

                    
                }
            */
                trac.Cells["A1"].Value = "Requirement ID";
                trac.Cells["B1"].Value = "Section";
                trac.Cells["C1"].Value = "Subsection";
                trac.Cells["D1"].Value = "Test Cases";


                


                package.Compression = CompressionLevel.BestSpeed;
                package.SaveAs(newFile);
                package.Stream.Close();
            }

            Console.WriteLine("{0:HH.mm.ss}\tDone!!", DateTime.Now);

            ConvertToPDF.convert();



        }

    }
    }

