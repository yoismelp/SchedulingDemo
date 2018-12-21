using Scheduling_Demo.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web.Mvc;

namespace Scheduling_Demo.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult ParameterPage()
        {
            ParameterPageModel ppm = new ParameterPageModel();
            DataAccessLayer da = new DataAccessLayer();
            ppm.GroupTopics = da.GetListOfGroupTopics();
            
            ppm.GroupDate = null;

            ViewBag.CurrentDate = DateTime.Now.ToString("MM/dd/yyyy");
            
            return View(ppm);
        }

        [HttpPost]
        public ActionResult ParameterPage(ParameterPageModel ppm)
        {
            return RedirectToAction("Index", new { Facility = ppm.SelectedFacility, GroupTopic = ppm.SelectedGroupTopic, GroupDate = ppm.GroupDate, LOC = ppm.SelectedLOC });
        }



        [HttpGet]
        public ActionResult Index(string Facility, string GroupTopic, DateTime GroupDate, string LOC)
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
                            "WHERE Facility = @Facility AND [Group Topic] = @GroupTopic AND [Group Date] = @GroupDate AND LOC = @LOC" ;

                        cmdExcel.Parameters.Add("@Facility", SqlDbType.NVarChar);
                        cmdExcel.Parameters["@Facility"].Value = Facility;

                        cmdExcel.Parameters.Add("@GroupTopic", SqlDbType.NVarChar);
                        cmdExcel.Parameters["@GroupTopic"].Value = GroupTopic;

                        cmdExcel.Parameters.Add("@GroupDate", SqlDbType.DateTime);
                        cmdExcel.Parameters["@GroupDate"].Value = GroupDate;

                        cmdExcel.Parameters.Add("@LOC", SqlDbType.NVarChar);
                        cmdExcel.Parameters["@LOC"].Value = LOC;

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

        [HttpPost]
        public ActionResult Index(AppointmentsViewModel vm)
        {
            string conString = string.Empty;
            //AppointmentsViewModel vm1 = new AppointmentsViewModel();
            //vm1.appointments.AddRange(vm.appointments);


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

                        foreach (Appointments a in vm.appointments)
                        {
                            string checkedInDate = a.CheckedIn ? DateTime.Now.ToString("MM/dd/yyyy") : string.Empty;
                            string checkedInTime = a.CheckedIn ? DateTime.Now.ToString("hh:mm tt") : string.Empty;
                            string cmd = string.Format(@"
UPDATE[{0}] 
SET [Checked In] = @CheckedIn, [Date Checked In] = '{1}', [Time] = '{2}'
WHERE PATID = @PATID
AND Facility = @Facility 
AND [Group Topic] = @GroupTopic 
AND LOC = @LOC 
AND [Group Date] = @GroupDate", sheetName, checkedInDate, checkedInTime);

                            cmdExcel.CommandText = cmd;

                            cmdExcel.Parameters.Add("@CheckedIn", SqlDbType.Bit);
                            cmdExcel.Parameters["@CheckedIn"].Value = a.CheckedIn;

                            cmdExcel.Parameters.Add("@PATID", SqlDbType.VarChar);
                            cmdExcel.Parameters["@PATID"].Value = a.MRNumber.ToString();

                            cmdExcel.Parameters.Add("@Facility", SqlDbType.NVarChar);
                            cmdExcel.Parameters["@Facility"].Value = a.Facility;

                            cmdExcel.Parameters.Add("@GroupTopic", SqlDbType.NVarChar);
                            cmdExcel.Parameters["@GroupTopic"].Value = a.GroupTopic;

                            cmdExcel.Parameters.Add("@LOC", SqlDbType.NVarChar);
                            cmdExcel.Parameters["@LOC"].Value = a.LOC;

                            cmdExcel.Parameters.Add("@GroupDate", SqlDbType.Date);
                            cmdExcel.Parameters["@GroupDate"].Value = Convert.ToDateTime(a.GroupDate);


                            cmdExcel.ExecuteNonQuery();
                        }

                        connExcel.Close();
                    }
                }
            }

            string facility = vm.appointments.Select(x => x.Facility).FirstOrDefault().ToString();
            string groupTopic = vm.appointments.Select(x => x.GroupTopic).FirstOrDefault().ToString();
            string groupDate = vm.appointments.Select(x => x.GroupDate).FirstOrDefault();
            string loc = vm.appointments.Select(x => x.LOC).FirstOrDefault();

            return RedirectToAction("Confirmation");//, new { Facility = facility, GroupTopic = groupTopic, GroupDate = groupDate, LOC = loc });
        }

        public ActionResult Confirmation()
        {
            return View();
        }
    }
}