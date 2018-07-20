using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Newtonsoft.Json;
using System.Diagnostics;

namespace Capstone.Controllers
{
    public class ScheduleGeneratorController : Controller
    {
        // GET: ScheduleGenerator
        public ActionResult Index()
        {
            //If for ever any reason we return here ... just go back to the main page!
            return RedirectToAction("Index", "Home", null);
        }

        [HttpPost]
        public ActionResult ImportExcel(HttpPostedFileBase excelFile)
        {
            if (excelFile == null || excelFile.ContentLength == 0)
            {
                TempData["UploadError"] = "You must upload a file.";
            }
            if (excelFile.FileName.EndsWith(".xls") || excelFile.FileName.EndsWith(".xlsx"))
            {
                TempData["UploadSuccess"] = "Upload successful!";
                // Do processing
                ProcessSchedule(excelFile);
            }
            else
            {
                TempData["UploadError"] = "Please upload a valid file.";
            }
            return RedirectToAction("Index", "Home", null);
        }

        public void ProcessSchedule(HttpPostedFileBase excelFile)
        {
            string path = Server.MapPath("~/Content/excel_storage/" + excelFile.FileName);
            if (System.IO.File.Exists(path))
                System.IO.File.Delete(path);
            excelFile.SaveAs(path);

            // Connect to recently saved excel sheet
            OleDbConnection conn = null;
            if (excelFile.FileName.EndsWith(".xls"))
                conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0; data source=" + path + "; Extended Properties=\"Excel 8.0;HDR=No;IMEX=1\";");
            if (excelFile.FileName.EndsWith(".xlsx"))
                conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=\"Excel 12.0 Xml;HDR=No;IMEX=1\";");

            conn.Open();
            DataTable data = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string excelSheet = data.Rows[0]["Table_Name"].ToString();

            OleDbCommand ExcelCommand = new OleDbCommand(@"SELECT * FROM [" + excelSheet + @"]", conn);
            OleDbDataAdapter ExcelAdapter = new OleDbDataAdapter(ExcelCommand);

            DataSet excelDataSet = new DataSet();
            ExcelAdapter.Fill(excelDataSet);
            conn.Close();
            Debug.WriteLine(JsonConvert.SerializeObject(excelDataSet, Formatting.Indented));
        }
    }
}