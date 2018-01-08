using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
//using Microsoft.Office.Interop.Excel;
using MVCMigration.Tools;
using System.IO;
using System.Data;

namespace MVCMigration.Controllers
{
    public class ExcelController : Controller
    {

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(FormCollection formCollection)
        {
            DataTable dt = null;
            if (Request.Files["MigrationSheetFile"].ContentLength > 0)  
            {
                string extension = System.IO.Path.GetExtension(Request.Files["MigrationSheetFile"].FileName).ToLower();  
                string query = null;  
                string connString = "";  
                  
                string[] validFileTypes = { ".xls", ".xlsx", ".csv" };

                string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), Request.Files["MigrationSheetFile"].FileName);  
                if (!Directory.Exists(path1))  
                {  
                    Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));  
                }  
                if (validFileTypes.Contains(extension))  
                {  
                    if (System.IO.File.Exists(path1))  
                    { System.IO.File.Delete(path1); }
                    Request.Files["MigrationSheetFile"].SaveAs(path1);  
                    if(extension==".csv")  
                    {  
                     //Not needed yet - wdavidjr   
                     //DataTable dt= Utility.ConvertCSVtoDataTable(path1);  
                     //ViewBag.Data = dt;  
                    }  
                    //Connection String to Excel Workbook  
                   else if (extension.Trim() == ".xls")  
                    {  
                        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";  
                        dt = Utility.ConvertXSLXtoDataTable(path1,connString);  
                        ViewBag.Data = dt;  
                    }  
                    else if (extension.Trim() == ".xlsx")  
                    {  
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";  
                        dt = Utility.ConvertXSLXtoDataTable(path1, connString);  
                        ViewBag.Data = dt;  
                    }  
  
                }  
                else  
                {  
                    ViewBag.Error = "Please Upload Files in .xls, .xlsx or .csv format";  
  
                }  
                  
              }

            foreach (DataRow row in dt.Rows)
            {
                string strName = row["Submitter's Name"].ToString();
                string strSupervisor = row["Submitter's Supervisor"].ToString();
                string strBusinessArea = row["Primary Business Impacted Area"].ToString();
                string strRequest = row["Type of Request"].ToString();
            }

            //Saving logic here after adding on the list - first part only (no loop of the another saving logic yet)
  
            return View("Index");  
        }  

    }
}


/*

 //Normal reading of Excel File
            /*
            if (Request != null)   
            {  
                HttpPostedFileBase file = Request.Files["UploadedFile"];  
                if ((file != null) && (file.ContentLength != 0) && !string.IsNullOrEmpty(file.FileName))   
                {  
                    string fileName = file.FileName;  
                    string fileContentType = file.ContentType;  
                    byte[] fileBytes = new byte[file.ContentLength];  
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));  
                }  
            }
            */

/*
            if (Request != null)
            {
                HttpPostedFileBase file = Request.Files["UploadedFile"];
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    var usersList = new List<Users>();
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        var currentSheet = package.Workbook.Worksheets;
                        var workSheet = currentSheet.First();
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;

                        for (int rowIterator = 2; rowIterator <= noOfRow; rowIterator++)
                        {
                            
                            var user = new Users();
                            user.FirstName = workSheet.Cells[rowIterator, 1].Value.ToString();
                            user.LastName = workSheet.Cells[rowIterator, 2].Value.ToString();
                            usersList.Add(user);
                   
*/