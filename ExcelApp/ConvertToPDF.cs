using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelApp
{
    public static class ConvertToPDF
    {


        public static void convert()
        {
            string excelLocation = "C://workbooks//sample7.xlsx";
            string outputLocation = "C://workbooks//sample7.pdf";

                try
                {
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    app.Visible = false;
                    Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Open(excelLocation, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlCorruptLoad.xlExtractData);
                    wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputLocation);
                    wkb.Close();
                    app.Quit();


                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }


            

        }

    }
}