using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Services;
using ESL.Models;
using Login_page_project.Models;
using OfficeOpenXml;

namespace Login_page_project
    {
    /// <summary>
    /// Summary description for WebService
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class WebService : System.Web.Services.WebService
        {

        private SqlConnection connection;

        DataTable ct = new DataTable();
        DataTable dt = new DataTable();
        DataTable overall = new DataTable();
        private void Connection()
            {
            string ConnectionString = ConfigurationManager.ConnectionStrings["connection"].ConnectionString;
            connection = new SqlConnection(ConnectionString);
            }

        //method to get the year from table and pass it to dropdown
        [WebMethod]
        public List<string> GetYear()
            {
            Connection();
            SqlCommand cmd = new SqlCommand("GetYear", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            List<string> year = new List<string>();
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            ct.Clear();
            using (sd)
                {

                sd.Fill(ct);
                foreach (DataRow item in ct.Rows)
                    {
                    year.Add(item["year"].ToString());
                    }
                }
            return year;
            }

        //method to get the locations from databse abd pass it to dropdown
        [WebMethod]
        public List<string> GetLocation()
            {
            Connection();
            SqlCommand cmd = new SqlCommand("GetLocation", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            List<string> location = new List<string>();
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            ct.Clear();
            using (sd)
                {

                sd.Fill(ct);
                foreach (DataRow item in ct.Rows)
                    {
                    location.Add(item["location"].ToString());
                    }
                }
            return location;
            }

        //getting the details from the database and filter it by location to show in table
        [WebMethod]
        public DataTable GetEmployeeDetailsByLocation(string location, string table, string year)
            {


            Connection();
            SqlCommand cmd = new SqlCommand("GetEmployeeDetailsByLocation", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@location", location);
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);

            connection.Open();
            sd.Fill(dt);
            connection.Close();
            return dt;
            }


        //method to get the tcb from databse abd pass it to dropdown
        [WebMethod]
        public List<string> GetTCB()
            {
            Connection();
            SqlCommand cmd = new SqlCommand("GetTCB", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            List<string> tcb = new List<string>();
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            ct.Clear();
            using (sd)
                {

                sd.Fill(ct);
                foreach (DataRow item in ct.Rows)
                    {
                    tcb.Add(item["TCB_Assigned"].ToString());
                    }
                }

            return tcb;
            }


        //getting the details from the database and filter it by table name and location to show in table
        [WebMethod]
        public DataTable GetEmployeeDetailsByTCB(string tcb, string table, string year)
            {

            Connection();
            SqlCommand cmd = new SqlCommand("GetEmployeeDetailsByTCB", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@tcb", tcb);
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            dt.Clear();
            connection.Open();
            sd.Fill(dt);
            connection.Close();
            return dt;
            }

        //if exit is selected
        [WebMethod]
        public DataTable GetExitedDetails(string table, string year)
            {
            Connection();
            SqlCommand cmd = new SqlCommand("Exited", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            overall.Clear();
            using (sd)
                {

                sd.Fill(overall);

                }

            return overall;
            }

        //if certified billed is selected
        [WebMethod]
        public DataTable CertifiedBilled(string table, string year)
            {
            Connection();
            SqlCommand cmd = new SqlCommand("CertifiedBilled", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            overall.Clear();
            using (sd)
                {

                sd.Fill(overall);

                }

            return overall;
            }

        //if certified not billed is selected
        [WebMethod]
        public DataTable NotBilledCertified(string table, string year)
            {
            Connection();
            SqlCommand cmd = new SqlCommand("CertifiedNotBilled", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            overall.Clear();
            using (sd)
                {

                sd.Fill(overall);

                }

            return overall;
            }
        //Log
        [WebMethod]
        public List<Employee> Get()
            {
            Connection();
            string fname = "GetEmployeeDetail2";
            connection.Open();
            SqlCommand con = new SqlCommand(fname, connection);
            con.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter(con);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            List<Employee> emp = new List<Employee>();
            foreach (DataRow dr in ds.Tables[0].Rows)
                {
                emp.Add(new Employee
                {
                    FileName = Convert.ToString(dr["FileName"]),
                    Success_Count = Convert.ToString(dr["Success_Count"]),
                    Failed_Count = Convert.ToString(dr["Failed_Count"]),
                   Uploaded_By = Convert.ToString(dr["Uploaded_By"]),
                   Uploaded_Date= Convert.ToString(dr["Uploaded_Date"])
                });
                }
            connection.Close();
            return emp;
            }
        [WebMethod]
        public List<Employee> GetSearch(string Search)
            {
            Connection();
            DataSet ds = new DataSet();
            string fname = "GetSearch";

            SqlCommand con = new SqlCommand(fname, connection);
            con.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter(con);
            using (sda)
                {
                connection.Open();
                con.Parameters.AddWithValue("@search", Search);
                con.ExecuteNonQuery();
                sda.Fill(ds);
                connection.Close();

                }
            List<Employee> emp = new List<Employee>();
            foreach (DataRow dr in ds.Tables[0].Rows)
                {
                emp.Add(new Employee
                {
                    FileName = Convert.ToString(dr["FileName"]),
                    Success_Count = Convert.ToString(dr["Success_Count"]),
                    Failed_Count = Convert.ToString(dr["Failed_Count"]),
                  Uploaded_By= Convert.ToString(dr["Uploaded_By"]),
                  Uploaded_Date = Convert.ToString(dr["Uploaded_Date"])
                });
                }
            connection.Close();
            return emp;
            }

        //if not certified billed is selected
        [WebMethod]
        public DataTable BilledNotCertified(string table, string year)
            {
            Connection();
            SqlCommand cmd = new SqlCommand("NotCertifiedBilled", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            overall.Clear();
            using (sd)
                {

                sd.Fill(overall);

                }

            return overall;
            }


        //if not certified not billed is selected
        [WebMethod]
        public DataTable NotBilledNotCertified(string table, string year)
            {
            Connection();
            SqlCommand cmd = new SqlCommand("NotCertifiedNotBilled", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@table", table);
            cmd.Parameters.AddWithValue("@year", year);
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            overall.Clear();



            using (sd)
                {

                sd.Fill(overall);

                }

            return overall;
            }

        //File Upload Manager
        [WebMethod]
        public string UploadExcel(string name, HttpPostedFileBase file)
            {
            string message = "";
            if (name == "ASETable")
                {
                var userList = new List<ASETable>();
                int RowCount = 0;
                string fileName = "";
                if (HttpContext.Current.Request != null)
                    {
                    try
                        {
                        if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                            {
                            fileName = file.FileName;
                            string fileContentType = file.ContentType;
                            byte[] fileBytes = new byte[file.ContentLength];
                            var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));

                            using (var package = new ExcelPackage(file.InputStream))
                                {
                                var currentSheet = package.Workbook.Worksheets;
                                var workSheet = currentSheet.First();
                                var noOfCol = workSheet.Dimension.End.Column;
                                RowCount = workSheet.Dimension.End.Row;

                                for (int rowIterator = 2; rowIterator <= RowCount; rowIterator++)
                                    {
                                    var Emp = new ASETable();
                                    Emp.Emp_Id = Convert.ToInt32(workSheet.Cells[rowIterator, 1].Value);
                                    Emp.Employee_Name = workSheet.Cells[rowIterator, 2].Value.ToString();
                                    Emp.Batch = Convert.ToInt32(workSheet.Cells[rowIterator, 3].Value);
                                    Emp.DOJ = Convert.ToDateTime(workSheet.Cells[rowIterator, 4].Value);
                                    Emp.Location = workSheet.Cells[rowIterator, 5].Value.ToString();
                                    Emp.Status = workSheet.Cells[rowIterator, 6].Value.ToString();
                                    Emp.Educational_Qualification = workSheet.Cells[rowIterator, 7].Value.ToString();
                                    Emp.Branch = workSheet.Cells[rowIterator, 8].Value.ToString();
                                    Emp.FCA_L_PL_Prog = workSheet.Cells[rowIterator, 9].Value.ToString();
                                    Emp.FCA_L_PL_Score = Convert.ToInt32(workSheet.Cells[rowIterator, 10].Value);
                                    Emp.FCA_L_PL_Date = Convert.ToDateTime(workSheet.Cells[rowIterator, 11].Value);
                                    Emp.TCB_Assigned = workSheet.Cells[rowIterator, 12].Value.ToString();
                                    Emp.TCB_Specialization = workSheet.Cells[rowIterator, 13].Value.ToString();
                                    Emp.Practice = workSheet.Cells[rowIterator, 14].Value.ToString();
                                    Emp.FCA_SE_1 = Convert.ToInt32(workSheet.Cells[rowIterator, 15].Value);
                                    Emp.FCA_SE_2 = Convert.ToInt32(workSheet.Cells[rowIterator, 16].Value);
                                    Emp.FCA_SE_3 = Convert.ToInt32(workSheet.Cells[rowIterator, 17].Value);
                                    Emp.FCA_SE_Cleared_Date = Convert.ToDateTime(workSheet.Cells[rowIterator, 18].Value);
                                    Emp.SE_Attempts = Convert.ToInt32(workSheet.Cells[rowIterator, 19].Value);
                                    Emp.TCB_Prog_Lang = workSheet.Cells[rowIterator, 20].Value.ToString();
                                    Emp.FCA_PL1 = Convert.ToInt32(workSheet.Cells[rowIterator, 21].Value);
                                    Emp.FCA_PL2 = Convert.ToInt32(workSheet.Cells[rowIterator, 22].Value);
                                    Emp.FCA_PL3 = Convert.ToInt32(workSheet.Cells[rowIterator, 23].Value);
                                    Emp.FCA_PL_Cleared_Date = Convert.ToDateTime(workSheet.Cells[rowIterator, 24].Value);
                                    Emp.PL_Attempts = Convert.ToInt32(workSheet.Cells[rowIterator, 25].Value);
                                    Emp.Hackathon_1_Feedback = workSheet.Cells[rowIterator, 26].Value.ToString();
                                    Emp.Hackathon_2_Feedback = workSheet.Cells[rowIterator, 27].Value.ToString();
                                    Emp.Prog_Lang_Training_Start = Convert.ToDateTime(workSheet.Cells[rowIterator, 28].Value);
                                    Emp.Prog_Lang_Trg_End = Convert.ToDateTime(workSheet.Cells[rowIterator, 29].Value);
                                    Emp.SCL_End = Convert.ToDateTime(workSheet.Cells[rowIterator, 30].Value);
                                    Emp.FCA_Due_Date = Convert.ToDateTime(workSheet.Cells[rowIterator, 31].Value);
                                    Emp.Training_Stage = Convert.ToInt32(workSheet.Cells[rowIterator, 32].Value);
                                    Emp.TCB_Trg_Start = Convert.ToDateTime(workSheet.Cells[rowIterator, 33].Value);
                                    Emp.Stage_2_Start = Convert.ToDateTime(workSheet.Cells[rowIterator, 34].Value);
                                    Emp.Stage_3_Start = Convert.ToDateTime(workSheet.Cells[rowIterator, 35].Value);
                                    Emp.Stage_4_Start = Convert.ToDateTime(workSheet.Cells[rowIterator, 36].Value);
                                    Emp.KenMAP_Level_3_Date1 = Convert.ToDateTime(workSheet.Cells[rowIterator, 37].Value);
                                    Emp.KenMAP_Level_3_Date2 = Convert.ToDateTime(workSheet.Cells[rowIterator, 38].Value);
                                    Emp.Project = workSheet.Cells[rowIterator, 39].Value.ToString();
                                    Emp.Project_Start_Date = Convert.ToDateTime(workSheet.Cells[rowIterator, 40].Value);
                                    userList.Add(Emp);
                                    }
                                }
                            }
                        }
                    catch (Exception )
                        {
                        message = "Excel Columns does not match. Please upload appropriate File";
                        }
                    }
                int result = 0;
                using (ESLEntities1 ImportDataDBEntities = new ESLEntities1())
                    {
                    foreach (var item in userList)
                        {
                        ImportDataDBEntities.ASETables.Add(item);
                        }

                    try
                        {
                        result = ImportDataDBEntities.SaveChanges();
                        }
                    catch (Exception)
                        {
                        System.Diagnostics.Trace.WriteLine("Data Duplicated");
                        message = "Data Already Exists";

                        }
                    if (result > 0)
                        {
                        message = "File Uploaded Successfully!!";
                        }
                    /*else if (result == 0)
                    {
                        message = "File Upload UnSuccessful!!";
                    }*/
                    }
                int remain = RowCount - result - 1;
                Total(result, remain, fileName);
                }

            else if (name == "TCRTable")
                {
                var TCRInfoList = new List<TCRTable>();
                string fileName = "";
                int rowCount = 0;
                if (HttpContext.Current.Request != null)
                    {
                    try
                        {
                        if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                            {
                            fileName = file.FileName;
                            string fileContentType = file.ContentType;
                            byte[] fileBytes = new byte[file.ContentLength];
                            var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));



                            using (var package = new ExcelPackage(file.InputStream))
                                {
                                var currentSheet = package.Workbook.Worksheets;
                                var workSheet = currentSheet.First();
                                var noOfCol = workSheet.Dimension.End.Column;
                                rowCount = workSheet.Dimension.End.Row;
                                for (int rowIterator = 2; rowIterator <= rowCount; rowIterator++)
                                    {
                                    var Emp = new TCRTable();
                                    Emp.Emp_Id = Convert.ToInt32(workSheet.Cells[rowIterator, 1].Value);
                                    Emp.Employee_Name = workSheet.Cells[rowIterator, 2].Value.ToString();
                                    Emp.Manager_Emp_Id = Convert.ToInt32(workSheet.Cells[rowIterator, 3].Value);
                                    Emp.Manager_Name = workSheet.Cells[rowIterator, 4].Value.ToString();
                                    Emp.Band = workSheet.Cells[rowIterator, 5].Value.ToString();
                                    Emp.Org_Unit_ID = Convert.ToInt32(workSheet.Cells[rowIterator, 6].Value);
                                    Emp.Org_Unit = workSheet.Cells[rowIterator, 7].Value.ToString();
                                    Emp.Emp_Selected_Primary_TCB = workSheet.Cells[rowIterator, 8].Value.ToString();
                                    Emp.Date_Of_Hire = Convert.ToDateTime(workSheet.Cells[rowIterator, 9].Value);
                                    Emp.Ongoing_TCB = workSheet.Cells[rowIterator, 10].Value.ToString();
                                    Emp.Hired_TCB = workSheet.Cells[rowIterator, 11].Value.ToString();
                                    Emp.Hired_TCB_Specialization = workSheet.Cells[rowIterator, 12].Value.ToString();
                                    Emp.Hired_Rating = Convert.ToInt32(workSheet.Cells[rowIterator, 13].Value);
                                    Emp.Date_of_SME_Assessment = Convert.ToDateTime(workSheet.Cells[rowIterator, 14].Value);
                                    Emp.SME_Emp_No = Convert.ToInt32(workSheet.Cells[rowIterator, 15].Value);
                                    Emp.SME_Name = workSheet.Cells[rowIterator, 16].Value.ToString();
                                    Emp.SME_Assessed_Primary_TCB = workSheet.Cells[rowIterator, 17].Value.ToString();
                                    Emp.SME_Assessed_KENMAP_Rating = Convert.ToInt32(workSheet.Cells[rowIterator, 18].Value);
                                    Emp.Current_Status = workSheet.Cells[rowIterator, 19].Value.ToString();
                                    Emp.Current_TCB_Requested_For = workSheet.Cells[rowIterator, 20].Value.ToString();
                                    Emp.Current_Date_of_Request = Convert.ToDateTime(workSheet.Cells[rowIterator, 21].Value);
                                    Emp.Current_Name_of_SME = workSheet.Cells[rowIterator, 22].Value.ToString();
                                    Emp.Current_Request_Id = Convert.ToInt32(workSheet.Cells[rowIterator, 23].Value);
                                    Emp.SME_Assessed_Specialization = workSheet.Cells[rowIterator, 24].Value.ToString();
                                    Emp.Employee_Specialization_Primary_TCB = workSheet.Cells[rowIterator, 25].Value.ToString();



                                    TCRInfoList.Add(Emp);
                                    }
                                }
                            }
                        }
                    catch (Exception ex)
                        {
                        message = "Excel Columns does not match. Please upload appropriate File";
                        }
                    }

                int result = 0;
                using (ESLEntities1 ImportDataDBEntities = new ESLEntities1())
                    {
                    foreach (var item in TCRInfoList)
                        {
                        ImportDataDBEntities.TCRTables.Add(item);
                        }

                    try
                        {
                        result = ImportDataDBEntities.SaveChanges();
                        }
                    catch (Exception ex)
                        {
                        System.Diagnostics.Trace.WriteLine("Data Duplicated");
                        message = "Data Already Exists";

                        }

                    if (result > 0)
                        {
                        message = "File Uploaded Successfully!!";

                        }
                    /*else if (result == 0)
                    {
                        message = "File Upload UnSuccessful!!";
                    }*/
                    }
                int remain = rowCount - result - 1;
                Total(result, remain, fileName);
                }
                 UpdateUpload();
            return message;

            }
             public void UpdateUpload()
        {
           

            string userid = Session["Identity"].ToString() ;
            
            Connection();
            SqlCommand con = new SqlCommand("GetEmployeeDetail1", connection);
            con.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter(con);
            using (sda)
            {
                connection.Open();
                con.Parameters.AddWithValue("@userid", userid);
                con.ExecuteNonQuery();
                connection.Close();

            }
        }

        public static void Total(int result, int remain, string filename)
            {
            var FileLog = new List<FileLogTable>();
            FileLog.Add(
            new FileLogTable
            {
                FileName = filename,
                Success_Count = result,
                Failed_Count = remain,
                Uploaded_Date = DateTime.Now
            }
            );



            using (ESLEntities1 ImportDataDBEntities = new ESLEntities1())
                {
                foreach (var item in FileLog)
                    {
                    ImportDataDBEntities.FileLogTables.Add(item);
                    }
                try
                    {
                    result = ImportDataDBEntities.SaveChanges();
                    }
                catch (Exception ex)
                    {
                    System.Diagnostics.Trace.WriteLine("Data Duplicated");
                    }
                }
            }

        //Method get table name
        [WebMethod]
        public List<string> GetTable()
            {
            Connection();



            List<string> table = new List<string>();
            SqlCommand cmd = new SqlCommand("GetTableNames", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sd = new SqlDataAdapter(cmd);
            ct.Clear();
            using (sd)
                {



                sd.Fill(ct);
                foreach (DataRow item in ct.Rows)
                    {
                    table.Add(item["name"].ToString());
                    }
                }
            return table;



            }

        //Method to get Table details
        [WebMethod]
        public DataTable GetTableDetails(string select)
            {

            if (select == "ASETable")
                {
                SqlCommand cmd = new SqlCommand("GetAseDetail", connection);
                cmd.CommandType = CommandType.StoredProcedure;



                SqlDataAdapter sd = new SqlDataAdapter(cmd);
                dt.Clear();
                connection.Open();
                sd.Fill(dt);
                connection.Close();
                }
            /*else if (select == "TCBTable")
            {
                SqlCommand cmd = new SqlCommand("GetTcbDetail", connection);
                cmd.CommandType = CommandType.StoredProcedure;

 

                SqlDataAdapter sd = new SqlDataAdapter(cmd);
                dt.Clear();
                connection.Open();
                sd.Fill(dt);
                connection.Close();
            }*/
            else if (select == "TCRTable")
                {
                SqlCommand cmd = new SqlCommand("GetTcrDetail", connection);
                cmd.CommandType = CommandType.StoredProcedure;



                SqlDataAdapter sd = new SqlDataAdapter(cmd);
                dt.Clear();
                connection.Open();
                sd.Fill(dt);
                connection.Close();
                }
            return dt;
            }


        }
    }
