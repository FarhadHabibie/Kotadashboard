using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using KOTAdashboard.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.IO;
using Microsoft.Extensions.Hosting;
using OfficeOpenXml;
using KOTAdashboard.Data;

namespace KOTAdashboard.Controllers
{
    public class HomeController : Controller
    {
        private readonly IHostEnvironment _hostEnvironment;
        private readonly KOTAdashboardContext _context;
        private readonly ILogger<HomeController> _logger;
        public HomeController(ILogger<HomeController> logger, KOTAdashboardContext context, IHostEnvironment hostEnvironment)
        {
            _logger = logger;
            _context = context;
            _hostEnvironment = hostEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        /* public IActionResult OnPostImport(IWebHostEnvironment _hostingEnvironment)
         {
             IFormFile file = Request.Form.Files[0];
             string folderName = "Upload";
             string webRootPath = _hostingEnvironment.WebRootPath;
             string newPath = Path.Combine(webRootPath, folderName);
             StringBuilder sb = new StringBuilder();
             if (!Directory.Exists(newPath))
             {
                 Directory.CreateDirectory(newPath);
             }
             if (file.Length > 0)
             {
                 string sFileExtension = Path.GetExtension(file.FileName).ToLower();
                 ISheet sheet;
                 string fullPath = Path.Combine(newPath, file.FileName);
                 using (var stream = new FileStream(fullPath, FileMode.Create))
                 {
                     file.CopyTo(stream);
                     stream.Position = 0;
                     if (sFileExtension == ".xls")
                     {
                         HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                         sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook  
                     }
                     else
                     {
                         XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                         sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook   
                     }
                     IRow headerRow = sheet.GetRow(0); //Get Header Row
                     int cellCount = headerRow.LastCellNum;
                     sb.Append("<table class='table'><tr>");
                     for (int j = 0; j < cellCount; j++)
                     {
                         NPOI.SS.UserModel.ICell cell = headerRow.GetCell(j);
                         if (cell == null || string.IsNullOrWhiteSpace(cell.ToString())) continue;
                         sb.Append("<th>" + cell.ToString() + "</th>");
                     }
                     sb.Append("</tr>");
                     sb.AppendLine("<tr>");
                     for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File
                     {
                         IRow row = sheet.GetRow(i);
                         if (row == null) continue;
                         if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                         for (int j = row.FirstCellNum; j < cellCount; j++)
                         {
                             if (row.GetCell(j) != null)
                                 sb.Append("<td>" + row.GetCell(j).ToString() + "</td>");
                         }
                         sb.AppendLine("</tr>");
                     }
                     sb.Append("</table>");
                 }
             }
             return this.Content(sb.ToString());
         }
         */

        [HttpPost]
        public ActionResult Index(IFormFile postedFile)
        {

            if (postedFile != null)
            {
                try
                {
                    string fileExtension = Path.GetExtension(postedFile.FileName);

                    //Validate uploaded file and return error.
                    if (fileExtension != ".xls" && fileExtension != ".xlsx")
                    {
                        ViewBag.Message = "Please select the excel file with .xls or .xlsx extension";
                        return View();
                    }

                    string folderPath = _hostEnvironment.WebRootPath;
                    //Check Directory exists else create one
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }

                    //Save file to folder
                    var filePath = folderPath + Path.GetFileName(postedFile.FileName);
                    var stream = System.IO.File.Create(filePath);
                    postedFile.CopyToAsync(stream);

                    //Get file extension

                    string excelConString = "";

                    //Get connection string using extension 
                    switch (fileExtension)
                    {
                        //If uploaded file is Excel 1997-2007.
                        case ".xls":
                            excelConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                            break;
                        //If uploaded file is Excel 2007 and above
                        case ".xlsx":
                            excelConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=YES'";
                            break;
                    }

                    //Read data from first sheet of excel into datatable
                    DataTable dt = new DataTable();
                    excelConString = string.Format(excelConString, filePath);

                    using (OleDbConnection excelOledbConnection = new OleDbConnection(excelConString))
                    {
                        using (OleDbCommand excelDbCommand = new OleDbCommand())
                        {
                            using (OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter())
                            {
                                excelDbCommand.Connection = excelOledbConnection;

                                excelOledbConnection.Open();
                                //Get schema from excel sheet
                                DataTable excelSchema = GetSchemaFromExcel(excelOledbConnection);
                                //Get sheet name
                                string sheetName = excelSchema.Rows[0]["TABLE_NAME"].ToString();
                                excelOledbConnection.Close();

                                //Read Data from First Sheet.
                                excelOledbConnection.Open();
                                excelDbCommand.CommandText = "SELECT * From [" + sheetName + "]";
                                excelDataAdapter.SelectCommand = excelDbCommand;
                                //Fill datatable from adapter
                                excelDataAdapter.Fill(dt);
                                excelOledbConnection.Close();
                            }
                        }
                    }

                    //Insert records to Employee table.
                    using (var context = new KOTAdashboardContext())
                    {
                        //Loop through datatable and add employee data to employee table. 
                        foreach (DataRow row in dt.Rows)
                        {
                            context.Coronas.Add(GetCoronaFromExcelRow(row));
                        }
                        context.SaveChanges();
                    }
                    ViewBag.Message = "Data Imported Successfully.";
                }
                catch (Exception ex)
                {
                    ViewBag.Message = ex.Message;
                }
            }
            else
            {
                ViewBag.Message = "Please select the file first to upload.";
            }
            return View();
        }

        private static DataTable GetSchemaFromExcel(OleDbConnection excelOledbConnection)
        {
            return excelOledbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        }

        //Convert each datarow into employee object
        private Corona GetCoronaFromExcelRow(DataRow row)
        {
            return new Corona
            {
                Tanggal =row[0].ToString(),
                kasusbaru = int.Parse(row[1].ToString()),
                kasusimpor = int.Parse(row[2].ToString()),
                kasuslokal = int.Parse(row[3].ToString())
            };
        }
    }
}
