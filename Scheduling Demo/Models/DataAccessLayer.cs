using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI.WebControls;

namespace Scheduling_Demo.Models
{
    public class DataAccessLayer
    {
        internal List<SelectListItem> GetListOfGroupTopics()
        {
            List<SelectListItem> liList = new List<SelectListItem>();

            string conString = ConfigurationManager.ConnectionStrings["ExcleConn"].ConnectionString;
            string sheetName = "DataTable$";

            DataTable dt = new DataTable();

            using (OleDbConnection connExcel = new OleDbConnection(conString))
            {
                using (OleDbCommand cmd = new OleDbCommand("SELECT DISTINCT[Group Topic] From[" + sheetName + "]", connExcel))
                {
                    connExcel.Open();
                    using (OleDbDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (!reader.IsDBNull(0))
                            {
                                SelectListItem li = new SelectListItem();
                                li.Text = reader["Group Topic"].ToString();
                                li.Value = reader["Group Topic"].ToString();
                                liList.Add(li);
                            }
                        }
                        
                    }
                }
            }

            return liList;
        }
    }
}