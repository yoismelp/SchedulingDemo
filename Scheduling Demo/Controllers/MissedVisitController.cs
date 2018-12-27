using Scheduling_Demo.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Web.Mvc;

namespace Scheduling_Demo.Controllers
{
    public class MissedVisitController : Controller
    {
        // GET: MissedVisit
        [HttpGet]
        public ActionResult MissedVisit()
        {
            MissedVisitLookupViewModel mv = new MissedVisitLookupViewModel();
            

            mv.MissedVisitDate = null;

            ViewBag.CurrentDate = DateTime.Now.ToString("MM/dd/yyyy");

            return View(mv);
        }

        [HttpPost]
        public ActionResult MissedVisit(MissedVisitLookupViewModel mv)
        {
            return RedirectToAction("ShowMissedVisits", new { missedVisitDate = mv.MissedVisitDate });
        }

        public ActionResult ShowMissedVisits(DateTime missedVisitDate)
        {
            AppointmentsViewModel vm = new AppointmentsViewModel();

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
                        cmdExcel.CommandText = "SELECT * " +
                            "From [" + sheetName + "] " +
                            "WHERE ([Group Date] = @GroupDate AND @MissedVisitDate < @CurrentDate " +
                            " AND ([Checked In] IS NULL OR [Checked In] = False OR [Checked In] = 0) " +
                            ") OR " +
                            "([Group Date] = @GroupDate AND @MissedVisitDate = @CurrentDate " +
                            " AND ([Checked In] IS NULL OR [Checked In] = False OR [Checked In] = 0) " +
                            " AND ([Group Time] IS NULL OR [Group Time] < @GroupTime)" +
                            ")";

                        cmdExcel.Parameters.Add("@GroupDate", SqlDbType.DateTime);
                        cmdExcel.Parameters["@GroupDate"].Value = missedVisitDate;

                        cmdExcel.Parameters.Add("@GroupTime", SqlDbType.NVarChar);
                        cmdExcel.Parameters["@GroupTime"].Value = DateTime.Now.ToString("HH:MM");


                        cmdExcel.Parameters.Add("@CurrentDate", SqlDbType.DateTime);
                        cmdExcel.Parameters["@CurrentDate"].Value = DateTime.Now.ToString("yyyy-MM-dd");

                        cmdExcel.Parameters.Add("@MissedVisitDate", SqlDbType.DateTime);
                        cmdExcel.Parameters["@MissedVisitDate"].Value = Convert.ToDateTime(missedVisitDate).ToString("yyyy-MM-dd");

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
                appointment.CheckedIn = string.IsNullOrEmpty(row["Checked In"].ToString()) || row["Checked In"].ToString().Equals("0")
                    ? false
                    : true;
                appointment.GroupDate = Convert.ToDateTime(row["Group Date"]).ToString("MM/dd/yyyy");
                appointment.Facility = row["Facility"].ToString();
                appointment.GroupTopic = row["Group Topic"].ToString();
                appointment.LOC = row["LOC"].ToString();
                apList.Add(appointment);
            }
            vm.appointments = apList;

            return View(vm);
        }
    }
}
