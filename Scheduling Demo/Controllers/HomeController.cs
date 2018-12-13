﻿using Scheduling_Demo.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Web.Mvc;

namespace Scheduling_Demo.Controllers
{
    public class HomeController : Controller
    {

        

        [HttpGet]
        public ActionResult Index()
        {
            AppointmentsViewModel appts = new AppointmentsViewModel();
            //string filePath = string.Empty;
            //if (postedFile != null)
            //{
            //string path = Server.MapPath(@"C:\Users\yperez\Desktop\GroupScheduling.xlsm");
            //if (!Directory.Exists(path))
            //{
            //    Directory.CreateDirectory(path);
            //}

            //filePath = path;
            //string extension = Path.GetExtension(postedFile.FileName);
            // postedFile.SaveAs(filePath);

            string conString = string.Empty;
                

                DataTable dt = new DataTable();
                conString = ConfigurationManager.ConnectionStrings["ExcleConn"].ConnectionString;
            string test = string.Empty;

                using (OleDbConnection connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {

                        cmdExcel.Connection = connExcel;
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = "DataTable$";
                        connExcel.Close();

                        //Read Data from First DataTable.
                        connExcel.Open();
                        cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";

                        odaExcel.SelectCommand = cmdExcel;
                        odaExcel.Fill(dt);
                        connExcel.Close();
                        }
                    }
                }

            
            Appointments appointment;
            List<Appointments> apList = new List<Appointments>();

                //Insert records to database table.
                //Appointments entities = new Appointments();
                foreach (DataRow row in dt.Rows)
                {
                //test += row["PATID"].ToString();
                appointment = new Appointments();
                appointment.MRNumber = row["PATID"].ToString();
                appointment.CheckedIn = string.IsNullOrEmpty(row["Checked In"].ToString()) ? false : true;
                appointment.Date = Convert.ToDateTime(row["Group Date"]).ToString("MM/dd/yyyy");
                appointment.Facility = row["Facility"].ToString();
                appointment.GroupTopic = row["Group Topic"].ToString();
                appointment.LOC = row["LOC"].ToString();
                apList.Add(appointment);

                //    entities.MRNumber.Add(new Customer
                //    {
                //        Name = row["Name"].ToString(),
                //        Country = row["Country"].ToString()
                //    });

                }
            appts.appointments = apList;
            //entities.SaveChanges();
            //}

            // return test;
           


            return View(appts);
        }

        [HttpPost]
        public ActionResult Index(AppointmentsViewModel appts)
        {
            string conString = string.Empty;
            AppointmentsViewModel vm = new AppointmentsViewModel();
            vm = appts;


            DataTable dt = new DataTable();


            conString = ConfigurationManager.ConnectionStrings["ExcleConn"].ConnectionString;
            string test = string.Empty;

            using (OleDbConnection connExcel = new OleDbConnection(conString))
            {
                using (OleDbCommand cmdExcel = new OleDbCommand())
                {
                    using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                    {

                        cmdExcel.Connection = connExcel;
                        connExcel.Open();
                        DataTable dtExcelSchema;
                        dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string sheetName = "DataTable$";
                        connExcel.Close();
                        connExcel.Open();

                        foreach (Appointments a in appts.appointments)
                        {
                            string cmd = string.Format(@"UPDATE[{0}] SET [Checked In] = {1} WHERE PATID = {2} AND Facility = '{3}' AND [Group Topic] = '{4}'AND [Group Date] = '{5}' AND LOC = '{6}'", sheetName, a.CheckedIn, a.MRNumber, a.Facility, a.GroupTopic, a.Date, a.LOC);
                            cmdExcel.CommandText = cmd;
                            cmdExcel.ExecuteNonQuery();
                        }
                        
                        connExcel.Close();
                    }
                }
            }

            return RedirectToAction("Index");
        }
    }
}