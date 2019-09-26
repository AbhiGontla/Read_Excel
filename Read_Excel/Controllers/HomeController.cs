using Read_Excel.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Read_Excel.Controllers
{
    public class HomeController : Controller
    {
        
        // GET: Home
        public ActionResult Index()
        {
            int id = Convert.ToInt32(TempData["ID"]);
            if(id==10)
            {
               ViewBag.ErrorMessage = "Please Upload the file";
            }
            if (id == 20)
            {
                ViewBag.ErrorMessage = "Please Upload Excel file only";
            }
            return View(new List<EmployeeModel>());
        }

        #region displaying in excel

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            if (postedFile == null)
            {
                

                TempData["ID"] = 10;
                return RedirectToAction("Index");
                //ErrorMessage = "Please upload the file";
                //return View(ErrorMessage);
            }

            try
            {
                var supportedTypes = new[] { "xls", "xlsx" };
                var fileExt = System.IO.Path.GetExtension(postedFile.FileName).Substring(1);
                if (!supportedTypes.Contains(fileExt))
                {
                    TempData["ID"] = 20;
                    return RedirectToAction("Index");

                }
                else
                {
                    List<EmployeeModel> employees = new List<EmployeeModel>();
                    string filePath = string.Empty;
                    if (postedFile != null)
                    {
                        string path = Server.MapPath("~/Uploads/");
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }

                        filePath = path + Path.GetFileName(postedFile.FileName);
                        string extension = Path.GetExtension(postedFile.FileName);
                        postedFile.SaveAs(filePath);

                        string conString = string.Empty;
                        switch (extension)
                        {
                            case ".xls": //Excel 97-03.
                                conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                                break;
                            case ".xlsx": //Excel 07 and above.
                                conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                                break;
                        }

                        conString = string.Format(conString, filePath);

                        using (OleDbConnection connExcel = new OleDbConnection(conString))
                        {
                            using (OleDbCommand cmdExcel = new OleDbCommand())
                            {
                                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                                {
                                    DataTable dt = new DataTable();
                                    cmdExcel.Connection = connExcel;

                                    //Get the name of First Sheet.
                                    connExcel.Open();
                                    DataTable dtExcelSchema;
                                    dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                    string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                    connExcel.Close();

                                    //Read Data from First Sheet.
                                    connExcel.Open();
                                    cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                    odaExcel.SelectCommand = cmdExcel;
                                    odaExcel.Fill(dt);
                                    connExcel.Close();

                                    foreach (DataRow row in dt.Rows)
                                    {
                                        employees.Add(new EmployeeModel
                                        {
                                            EmployeeId = Convert.ToInt32(row["Id"]),
                                            Name = row["Name"].ToString(),
                                            Address = row["Address"].ToString(),
                                            EmailId = row["EmailId"].ToString()
                                        });
                                    }
                                }
                            }
                        }
                    }

                    return View(employees);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return null;
        }
    }
    #endregion
}