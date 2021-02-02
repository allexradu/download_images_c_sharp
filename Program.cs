using System;
using System.IO;
using System.Net;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace down
{
    class MainClass
    {
        public static void Main(string[] args) {
            string workingDirectory = Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            string imgPath = Path.Combine(projectDirectory, "download");
            string finalExcelPath = Path.Combine(projectDirectory, "excel", "a.xlsx");

            string allex = "test.jpg";
            string[] array = allex.Split('.');
            print(array[array.Length - 1]);


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Creating an instance 
            // of ExcelPackage 
            //ExcelPackage excel = new ExcelPackage();

            FileInfo fileInfo = new FileInfo(finalExcelPath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

            // name of the sheet 
            //var workSheet = excel.Workbook.Worksheets.Add("Sheet1");

            // setting the properties 
            // of the work sheet  
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;


            // Header of the Excel sheet 
           

            // get number of rows and columns in the sheet
            int rows = workSheet.Dimension.Rows; // 20
            int columns = workSheet.Dimension.Columns; // 7

            // loop through the worksheet rows and columns
            for (int i = 2; i <= rows; i++)
            {
                using (WebClient webClient = new WebClient())
                {
                    if (workSheet.Cells[i, 2].Value != null)
                    {
                        print(i.ToString());
                        string url = workSheet.Cells[i, 2].Value.ToString();
                        string[] split_values = url.Split('/');
                        webClient.DownloadFile(url, Path.Combine(imgPath, split_values[split_values.Length - 1]));

                    }
                }

            }


            //if (File.Exists(finalExcelPath))
            //    File.Delete(finalExcelPath);

            //// Create excel file on physical disk  
            //FileStream objFileStrm = File.Create(finalExcelPath);
            //objFileStrm.Close();

            //// Write content to excel file  
            //File.WriteAllBytes(finalExcelPath, excel.GetAsByteArray());
            ////Close Excel package 
            //excel.Dispose();
            //Console.ReadKey();




            //string workingDirectory = Environment.CurrentDirectory;
            //string projectDirectory = Directory.GetParent(workingDirectory).Parent.FullName;
            //string fullImgPath = Path.Combine(projectDirectory, "download","img.jpeg");
            //string excelPath = Path.Combine(projectDirectory, "excel", "a.xlsx");


            //print(excelPath);
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            //FileInfo fileInfo = new FileInfo(excelPath);
            //ExcelPackage package = new ExcelPackage(fileInfo);
            //ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            //ExcelPackage excel = new ExcelPackage();

            //// get number of rows and columns in the sheet
            //int rows = worksheet.Dimension.Rows; // 20
            //int columns = worksheet.Dimension.Columns; // 7

            //print($"{rows} : {columns}");

            //worksheet.Cells[1, 1].Value = "test";

            //FileStream objFileStrm = File.Create(excelPath);
            //objFileStrm.Close();

            //File.WriteAllBytes(excelPath, excel.GetAsByteArray());
            //excel.Dispose();


            //// loop through the worksheet rows and columns
            //for (int i = 1; i <= rows; i++)
            //{
            //    for (int j = 1; j <= columns; j++)
            //    {

            //        string content = worksheet.Cells[i, j].Value.ToString();
            //        /* Do something ...*/
            //    }
            //}

            //using (WebClient webClient = new WebClient())
            //{
            //    webClient.DownloadFile("https://i.imgur.com/QnD1myx.jpeg", fullPath);
            //}
        }

        public static void print(string s)
        {
            Console.WriteLine(s);
        }

    }
}
