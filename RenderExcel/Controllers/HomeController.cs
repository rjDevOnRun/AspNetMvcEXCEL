using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebGrease.Css.Ast;

namespace RenderExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            string path = Server.MapPath("~/Uploads/");
            string xlFilePath = string.Empty;
            string xlFileExtention = string.Empty;
            DataTable dtSheet = new DataTable();
            DataSet ExcelData = new DataSet();

            if (postedFile != null)
            {
                if(!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                xlFilePath = path + Path.GetFileName(postedFile.FileName);
                xlFileExtention = Path.GetExtension(postedFile.FileName);

                postedFile.SaveAs(xlFilePath);
            }
            string connectionString = string.Empty;
            
            // Create connection string bases on file extensions
            switch (xlFileExtention)
            {
                case ".xls": // Excel 97~2003
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    }
                    break;
                case ".xlsx": // Excel 2007 ~
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    }
                    break;
            }

            // oledb Connection
            connectionString = string.Format(connectionString, xlFilePath);
            using(OleDbConnection conExcel = new OleDbConnection(connectionString))
            {
                using(OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using(OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {
                        cmdExcel.Connection = conExcel;
                        conExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = conExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                        conExcel.Close();

                        // read xl data from first sheet
                        conExcel.Open();
                        cmdExcel.CommandText = "SELECT * FROM [" + sheetName + "]";
                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dtSheet);
                        conExcel.Close();

                    }
                }
            }
            ExcelData.Tables.Add(dtSheet);
            return View(ExcelData);
        }
    }
}