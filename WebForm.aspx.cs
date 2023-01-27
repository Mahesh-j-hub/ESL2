using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;
using System.Web.UI.DataVisualization.Charting;
using System.Web.UI.WebControls;
using System.Net.Mail;
using System.IO;
using System.Web.Security;
using System.Collections.Generic;
using System.Data;
using System.IO.Compression;

namespace ESL_Chart
{
    public partial class WebForm : System.Web.UI.Page
    {
        public string Email;
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (!IsPostBack)
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                    {
                        SqlCommand commandForBatches = new SqlCommand("select Count(Distinct Batch) from Tracker", connection);
                        connection.Open();
                        List<string> batches = new List<string>();
                        int BatchCount = (int)commandForBatches.ExecuteScalar();
                        for (int index = 1; index <= BatchCount; index++)
                        {
                            DDBatch.Items.Add(string.Format("Batch {0}", index));
                        }
                        DDBatch.Items.Add("All Batches");
                    }

                    foreach (int chartType in Enum.GetValues(typeof(SeriesChartType)))
                    {
                        ListItem li = new ListItem(Enum.GetName(typeof(SeriesChartType), chartType), chartType.ToString());
                        // Adding chart types to list
                        DDChartType.Items.Add(li);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
        }

        protected void GenerateGraphEvent(object sender, EventArgs e)
        {
            GridView1.Visible = false;
            GridView2.Visible = false;
            GridView3.Visible = false;
            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            para.Visible = false;
            goDown.Visible = false;

            if (DDTest.SelectedIndex == 3 && DDBatch.SelectedValue != "All Batches")
            {
                GenerateGraphForSME();
                Button2.Visible = true;
            }
            else if (DDTest.SelectedIndex != 0 && DDBatch.SelectedValue != "All Batches")
            {
                DDAttempts.Visible = false;
                Button2.Visible = true;
                GenerateGraph();
                ListItem item = DDBatch.Items.FindByValue("All Batches");
                if (item == null)
                {
                    DDBatch.Items.Add("All Batches");
                }
            }
            else if (DDBatch.SelectedValue == "All Batches" && (DDTest.SelectedIndex == 1 || DDTest.SelectedIndex == 2 || DDTest.SelectedIndex == 3))
            {
                if (DDChartType.SelectedValue != "10")
                {
                    Alert("Only Column Chart Type is allowed for All Batches");
                }
                else if (DDTest.SelectedIndex == 3)
                {
                    Chart1.Visible = false;
                    Chart2.Visible = false;
                    Chart3.Visible = true;
                    Button2.Visible = false;
                    AllBatchChartForSME();
                }
                else
                {
                    Chart1.Visible = false;
                    Chart2.Visible = false;
                    Chart3.Visible = true;
                    Button2.Visible = false;
                    AllBatchChart();
                }
            }
            else
            {
                Alert("Please select the options before preceding to Generate the Charts");
            }
        }

        // SME
        public void GenerateGraphForSME()
        {
            try
            {
                List<string> attempts = new List<string>() { "Attempt 1", "Attempt 2" };
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                {
                    Series series = Chart1.Series["Series1"];
                    Chart1.Legends.Add("legend");
                    SqlCommand command = new SqlCommand("sp_SME", conn);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@Batch", DDBatch.SelectedValue.Split(' ')[1]);

                    conn.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    int i = 0;
                    while (reader.Read())
                    {
                        int result1 = reader.GetInt32(0);
                        int result2 = reader.GetInt32(1);
                        series.Points.AddXY(attempts[i++], result1);
                        series.Points.AddXY(attempts[i++], result2);
                    }

                    // Titles for X and Y axis
                    Chart1.ChartAreas[0].AxisX.Title = "Attempts";
                    Chart1.ChartAreas[0].AxisY.Title = "Number of Employees";

                    // Setting the chart type
                    series.ChartType = (SeriesChartType)Enum.Parse(typeof(SeriesChartType), DDChartType.SelectedValue);
                    if (series.ChartType == (SeriesChartType)Enum.Parse(typeof(SeriesChartType), "Column"))
                    {
                        Chart1.Legends.RemoveAt(0);
                    }

                    series.IsValueShownAsLabel = true;
                    Chart1.Series[0].Font = new Font("Arial", 12);

                    // Setting the Chart Dimension
                    Chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = DDDimension.SelectedIndex == 2 ? true : false;
                    // Chart Title
                    Title title = Chart1.Titles.Add(DDTest.SelectedValue);
                    title.Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold);

                    // Binding the DataReader to Chart
                    Chart1.DataSource = reader;
                    Chart1.DataBind();
                }
            }
            catch (Exception ex)
            {
                Alert("Invalid Options selected");
                ExceptionLogging.LogError(ex);
            }
        }

        public void GenerateGraph()
        {
            try
            {
                List<string> attempts = new List<string>() { "Attempt 1", "Attempt 2", "Attempt 3" };
                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                {

                    Series series = Chart1.Series["Series1"];
                    Chart1.Legends.Add("legend");
                    SqlCommand command = new SqlCommand(DDTest.SelectedIndex == 1 ? "sp_FCA_PL" : "sp_FCA_SE", conn);
                    command.CommandType = System.Data.CommandType.StoredProcedure;
                    command.Parameters.AddWithValue("@Batch", DDBatch.SelectedValue.Split(' ')[1]);
                    conn.Open();

                    SqlDataReader reader = command.ExecuteReader();
                    int Count = 0;
                    while (reader.Read())
                    {
                        int result1 = reader.GetInt32(0);
                        int result2 = reader.GetInt32(1);
                        int result3 = reader.GetInt32(2);
                        series.Points.AddXY(attempts[Count++], result1);
                        series.Points.AddXY(attempts[Count++], result2);
                        series.Points.AddXY(attempts[Count++], result3);
                    }

                    // Titles for X and Y axis
                    Chart1.ChartAreas[0].AxisX.Title = "Attempts";
                    Chart1.ChartAreas[0].AxisY.Title = "Number of Employees";

                    // Setting the chart type
                    series.ChartType = (SeriesChartType)Enum.Parse(typeof(SeriesChartType), DDChartType.SelectedValue);
                    if (series.ChartType == (SeriesChartType)Enum.Parse(typeof(SeriesChartType), "Column"))
                    {
                        Chart1.Legends.RemoveAt(0);
                    }

                    // Setting the Chart Dimension
                    Chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = DDDimension.SelectedIndex == 2 ? true : false;
                    series.IsValueShownAsLabel = true;
                    Chart1.Series[0].Font = new Font("Arial", 12);

                    // Chart Title
                    Title title = Chart1.Titles.Add(DDTest.SelectedValue);
                    title.Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold);

                    // Binding the DataReader to Chart
                    Chart1.DataSource = reader;
                    Chart1.DataBind();
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
        }

        protected void DetailEvent(object sender, EventArgs e)
        {
            DetailChart();
        }

        public void DetailChart()
        {
            if (DDTest.SelectedIndex != 0)
            {
                try
                {
                    DDAttempts.Visible = true;
                    Button1.Visible = false;
                    Button4.Visible = true;

                    if (DDTest.SelectedIndex != 3)
                        GenerateGraph();
                    else
                        GenerateGraphForSME();

                    Series series = Chart2.Series["Series2"];
                    Chart2.Legends.Add("legend");
                    SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString);
                    SqlCommand cmd = new SqlCommand(SqlQuerysForBrief(), conn);
                    conn.Open();

                    SqlDataReader rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        series.Points.AddXY(rdr[1].ToString(), rdr[0]);
                    }
                    series.ChartType = (SeriesChartType)Enum.Parse(typeof(SeriesChartType), DDChartType.SelectedValue);
                    if (series.ChartType == (SeriesChartType)Enum.Parse(typeof(SeriesChartType), "Column"))
                    {
                        Chart2.Legends.RemoveAt(0);
                    }
                    Chart2.ChartAreas[0].AxisX.Title = "Particular TCB's";
                    Chart2.ChartAreas[0].AxisY.Title = "Number of Employees";
                    Random random = new Random();
                    foreach (var item in Chart2.Series[0].Points)
                    {
                        Color c = Color.FromArgb(random.Next(0, 255), random.Next(0, 255), random.Next(0, 255));
                        item.Color = c;
                    }
                    Title title = Chart2.Titles.Add(DDAttempts.SelectedValue + " Detail of " + DDTest.SelectedValue);
                    title.Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold);
                    Chart2.ChartAreas["ChartArea2"].Area3DStyle.Enable3D = DDDimension.SelectedIndex == 2 ? true : false;
                    series.IsValueShownAsLabel = true;
                    Chart2.Series[0].Font = new Font("Arial", 12);

                    Chart2.DataSource = rdr;
                    Chart2.DataBind();
                }
                catch (Exception ex)
                {
                    ExceptionLogging.LogError(ex);
                    Response.Redirect("Error.html");
                }
            }
        }

        protected void DDAttemptsSelectedIndexChanged(object sender, EventArgs e)
        {
            DetailChart();
        }

        public void AllBatchChartForSME()
        {
            try
            {
                goDown.Visible = true;
                para.Visible = true;

                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                {
                    Series series1 = Chart3.Series["Attempt 1"];
                    Series series2 = Chart3.Series["Attempt 2"];
                    Series series3 = Chart3.Series["Attempt 3"];

                    // Removing the Chart3 3rd series
                    Chart3.Series.Remove(series3);

                    SqlCommand commandForBatches = new SqlCommand("select Count(Distinct Batch) from Tracker", connection);
                    connection.Open();
                    List<string> batches = new List<string>();
                    int BatchCount = (int)commandForBatches.ExecuteScalar();
                    for (int index = 1; index <= BatchCount; index++)
                    {
                        batches.Add(string.Format("Batch {0}", index));
                    }

                    List<string> one = new List<string>();
                    List<string> two = new List<string>();

                    using (SqlCommand cmd = new SqlCommand("sp_All_Batch_MultiLevel_SME", connection))
                    {
                        // Setting Query as Command Type
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlDataReader rdr = cmd.ExecuteReader();
                        DataTable dt = new DataTable();
                        dt.Load(rdr);

                        foreach (DataRow i in dt.Rows)
                        {
                            one.Add(i["1st Attempt"].ToString());
                            two.Add(i["2nd Attempt"].ToString());
                        }

                        Chart3.Series["Attempt 1"].Points.DataBindXY(batches, one);
                        Chart3.Series["Attempt 2"].Points.DataBindXY(batches, two);

                        Legend secondLegend = new Legend("MyLegend");

                        Chart3.Legends.Add("MyLegend");
                        Chart3.Series[0].Legend = "MyLegend";
                        Chart3.Series[1].Legend = "MyLegend";

                        Chart3.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = DDDimension.SelectedIndex == 2 ? true : false;
                        series1.IsValueShownAsLabel = true;
                        series2.IsValueShownAsLabel = true;
                        Chart3.Series[0].Font = new Font("Arial", 12);
                        Chart3.Series[1].Font = new Font("Arial", 12);

                        Title title = Chart3.Titles.Add(string.Format("{0} of {1}", DDTest.SelectedValue, DDBatch.SelectedValue));
                        title.Font = new System.Drawing.Font("Arial", 10, FontStyle.Bold);
                    }
                }

                if (DDTest.SelectedIndex == 3)
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                    {
                        connection.Open();

                        // FCA_PL 1st Attempt
                        label1.Text = "SME Assessment - 1st Attempt";
                        SqlCommand cmdFCAPL1 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker where [KenMAP_Level_3_Date1] IS NOT NULL AND [KenMAP_Level_3_Date2] IS NULL AND [KenMap_Level_3_Rating] >= 3 GROUP BY TCB_Assigned UNION Select 'Total', COUNT(*) FROM Tracker where [KenMAP_Level_3_Date1] IS NOT NULL AND [KenMAP_Level_3_Date2] IS NULL AND [KenMap_Level_3_Rating] >= 3", connection);
                        SqlDataReader rdrFCAPL1 = cmdFCAPL1.ExecuteReader();
                        DataTable dtFCAPL1 = new DataTable();
                        dtFCAPL1.Load(rdrFCAPL1);
                        GridView1.DataSource = dtFCAPL1;
                        GridView1.DataBind();

                        // FCA_PL 2nd Attempt
                        label2.Text = "SME Assessment - 2nd Attempt";
                        SqlCommand cmdFCAPL2 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker where [KenMAP_Level_3_Date1] IS NOT NULL AND [KenMAP_Level_3_Date2] IS NOT NULL AND [KenMap_Level_3_Rating] >= 3 GROUP BY TCB_Assigned UNION Select 'Total', COUNT(*) FROM Tracker where [KenMAP_Level_3_Date1] IS NOT NULL AND [KenMAP_Level_3_Date2] IS NOT NULL AND [KenMap_Level_3_Rating] >= 3 ", connection);
                        SqlDataReader rdrFCAPL2 = cmdFCAPL2.ExecuteReader();
                        DataTable dtFCAPL2 = new DataTable();
                        dtFCAPL2.Load(rdrFCAPL2);
                        GridView2.DataSource = dtFCAPL2;
                        GridView2.DataBind();

                        GridView3.Visible = false;
                        label3.Visible = false;

                        GridView1.Visible = true;
                        GridView2.Visible = true;
                        label1.Visible = true;
                        label2.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
        }

        public void AllBatchChart()
        {
            try
            {
                GridView1.Visible = true;
                GridView2.Visible = true;
                GridView3.Visible = true;
                label1.Visible = true;
                label2.Visible = true;
                label3.Visible = true;
                para.Visible = true;
                goDown.Visible = true;

                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                {
                    Series series1 = Chart3.Series["Attempt 1"];
                    Series series2 = Chart3.Series["Attempt 2"];
                    Series series3 = Chart3.Series["Attempt 3"];

                    SqlCommand commandForBatches = new SqlCommand("select Count(Distinct Batch) from Tracker", connection);
                    connection.Open();
                    List<string> batches = new List<string>();
                    int BatchCount = (int)commandForBatches.ExecuteScalar();
                    for (int index = 1; index <= BatchCount; index++)
                    {
                        batches.Add(string.Format("Batch {0}", index));
                    }

                    List<string> one = new List<string>();
                    List<string> two = new List<string>();
                    List<string> three = new List<string>();
                    using (SqlCommand cmd = new SqlCommand(DDTest.SelectedIndex == 1 ? "sp_All_Batch_MultiLevel_PL" : "sp_All_Batch_MultiLevel_SE", connection))
                    {
                        // Setting Query as Command Type
                        cmd.CommandType = CommandType.StoredProcedure;
                        //connection.Open();
                        SqlDataReader rdr = cmd.ExecuteReader();
                        DataTable dt = new DataTable();
                        dt.Load(rdr);

                        foreach (DataRow i in dt.Rows)
                        {
                            one.Add(i["1st Attempt"].ToString());
                            two.Add(i["2nd Attempt"].ToString());
                            three.Add(i["3rd Attempt"].ToString());
                        }

                        Chart3.Series["Attempt 1"].Points.DataBindXY(batches, one);

                        Chart3.Series["Attempt 2"].Points.DataBindXY(batches, two);

                        Chart3.Series["Attempt 3"].Points.DataBindXY(batches, three);
                        Chart3.Series["Attempt 3"].Color = Color.DarkSeaGreen;

                        Legend secondLegend = new Legend("MyLegend");

                        this.Chart3.Legends.Add("MyLegend");
                        Chart3.Series[0].Legend = "MyLegend";
                        Chart3.Series[1].Legend = "MyLegend";
                        Chart3.Series[2].Legend = "MyLegend";

                        Chart3.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = DDDimension.SelectedIndex == 2 ? true : false;
                        Title title = Chart3.Titles.Add(string.Format("{0} of {1}", DDTest.SelectedValue, DDBatch.SelectedValue));
                        //series.IsValueShownAsLabel = true;
                        series1.IsValueShownAsLabel = true;
                        series2.IsValueShownAsLabel = true;
                        series3.IsValueShownAsLabel = true;
                        Chart3.Series[0].Font = new Font("Arial", 12);
                        Chart3.Series[1].Font = new Font("Arial", 12);
                        Chart3.Series[2].Font = new Font("Arial", 12);
                    }
                }

                if (DDTest.SelectedIndex == 1)
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                    {
                        connection.Open();

                        // FCA_PL 1st Attempt
                        label1.Text = "1st Attempt FCA Programming Language";
                        SqlCommand cmdFCAPL1 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker WHERE FCA_PL1 >= 50 GROUP BY TCB_Assigned UNION SELECT 'Total', COUNT(*) FROM Tracker WHERE FCA_PL1 >= 50", connection);
                        SqlDataReader rdrFCAPL1 = cmdFCAPL1.ExecuteReader();
                        DataTable dtFCAPL1 = new DataTable();
                        dtFCAPL1.Load(rdrFCAPL1);
                        GridView1.DataSource = dtFCAPL1;
                        GridView1.DataBind();

                        // FCA_PL 2nd Attempt
                        label2.Text = "2nd Attempt FCA Programming Language";
                        SqlCommand cmdFCAPL2 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker WHERE FCA_PL2 >= 50 GROUP BY TCB_Assigned UNION SELECT 'Total', COUNT(*) FROM Tracker WHERE FCA_PL2 >= 50", connection);
                        SqlDataReader rdrFCAPL2 = cmdFCAPL2.ExecuteReader();
                        DataTable dtFCAPL2 = new DataTable();
                        dtFCAPL2.Load(rdrFCAPL2);
                        GridView2.DataSource = dtFCAPL2;
                        GridView2.DataBind();

                        // FCA_PL 3rd Attempt
                        label3.Text = "3rd Attempt FCA Programming Language";
                        SqlCommand cmdFCAPL3 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker WHERE FCA_PL3 >= 50 GROUP BY TCB_Assigned UNION SELECT 'Total', COUNT(*) FROM Tracker WHERE FCA_PL3 >= 50", connection);
                        SqlDataReader rdrFCAPL3 = cmdFCAPL3.ExecuteReader();
                        DataTable dtFCAPL3 = new DataTable();
                        dtFCAPL3.Load(rdrFCAPL3);
                        GridView3.DataSource = dtFCAPL3;
                        GridView3.DataBind();
                    }
                }

                if (DDTest.SelectedIndex == 2)
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["conn"].ConnectionString))
                    {
                        connection.Open();

                        // FCA_SE 1st Attempt
                        label1.Text = "1st Attempt FCA Software Engineering";
                        SqlCommand cmdFCAPL1 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker WHERE FCA_SE_1 >= 50 GROUP BY TCB_Assigned UNION SELECT 'Total', COUNT(*) FROM Tracker WHERE FCA_SE_1 >= 50", connection);
                        SqlDataReader rdrFCAPL1 = cmdFCAPL1.ExecuteReader();
                        DataTable dtFCAPL1 = new DataTable();
                        dtFCAPL1.Load(rdrFCAPL1);
                        GridView1.DataSource = dtFCAPL1;
                        GridView1.DataBind();

                        // FCA_SE 2nd Attempt
                        label2.Text = "2nd Attempt FCA Software Engineering";
                        SqlCommand cmdFCAPL2 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker WHERE FCA_SE_2 >= 50 GROUP BY TCB_Assigned UNION SELECT 'Total', COUNT(*) FROM Tracker WHERE FCA_SE_2 >= 50", connection);
                        SqlDataReader rdrFCAPL2 = cmdFCAPL2.ExecuteReader();
                        DataTable dtFCAPL2 = new DataTable();
                        dtFCAPL2.Load(rdrFCAPL2);
                        GridView2.DataSource = dtFCAPL2;
                        GridView2.DataBind();

                        // FCA_SE 3rd Attempt
                        label3.Text = "3rd Attempt FCA Software Engineering";
                        SqlCommand cmdFCAPL3 = new SqlCommand("SELECT TCB_Assigned, COUNT(*) as 'Number of Engineers Cleared' FROM Tracker WHERE FCA_SE_3 >= 50 GROUP BY TCB_Assigned UNION SELECT 'Total', COUNT(*) FROM Tracker WHERE FCA_SE_3 >= 50", connection);
                        SqlDataReader rdrFCAPL3 = cmdFCAPL3.ExecuteReader();
                        DataTable dtFCAPL3 = new DataTable();
                        dtFCAPL3.Load(rdrFCAPL3);
                        GridView3.DataSource = dtFCAPL3;
                        GridView3.DataBind();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
        }


        public string SqlQuerysForBrief()
        {
            try
            {
                if (DDTest.SelectedIndex == 1)
                {
                    ListItem item = DDAttempts.Items.FindByValue("Attempt 3");
                    if (item == null) DDAttempts.Items.Add("Attempt 3");
                    if (DDAttempts.SelectedIndex == 0)
                    {
                        return string.Format("exec sp_Detail 'FCA_PL1', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else if (DDAttempts.SelectedIndex == 1)
                    {
                        return string.Format("exec sp_Detail 'FCA_PL2', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else if (DDAttempts.SelectedIndex == 2)
                    {
                        return string.Format("exec sp_Detail 'FCA_PL3', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else
                        return string.Format("exec sp_Detail 'FCA_PL2', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                }
                else if (DDTest.SelectedIndex == 3)
                {
                    DDAttempts.Items.Remove("Attempt 3");
                    if (DDAttempts.SelectedValue == "Attempt 1")
                    {
                        return string.Format("exec sp_SME_Details 1, {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else
                    {
                        return string.Format("exec sp_SME_Details 2, {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                }
                else
                {
                    if (DDAttempts.SelectedIndex == 0)
                    {
                        return string.Format("exec sp_Detail 'FCA_SE_1', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else if (DDAttempts.SelectedIndex == 1)
                    {
                        return string.Format("exec sp_Detail 'FCA_SE_2', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else if (DDAttempts.SelectedIndex == 2)
                    {
                        return string.Format("exec sp_Detail 'FCA_SE_3', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                    }
                    else
                        return string.Format("exec sp_Detail 'FCA_SE_1', {0}", DDBatch.SelectedValue.Split(' ')[1]);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
            return "";
        }

        protected void GoBackToMainCharts(object sender, EventArgs e)
        {
            if (DDBatch.SelectedValue != "All Batches")
            {
                GenerateGraph();
                Button1.Visible = true;
                Button4.Visible = false;
                DDAttempts.Visible = false;
                Button3.Visible = true;
            }
            else
            {
                Alert("Invalid Data");
            }
        }

        protected void Save(object sender, EventArgs e)
        {
            if (DDTest.SelectedIndex != 0 || DDDimension.SelectedIndex != 0 || DDBatch.SelectedIndex != 0 || DDChartType.SelectedIndex != 0)
            {
                if (DDBatch.SelectedValue != "All Batches")
                {
                    GenerateGraph();
                    if (Button4.Visible == true)
                    {
                        DetailChart();
                        SaveDetailImage();
                    }
                    else if (ConfigurationManager.AppSettings["SendMail"] == "1")
                        SendMail(Chart1);
                    else
                    {
                        CreateDirectoryToSaveImage(Chart1);
                    }
                }
                else
                {
                    AllBatchChart();
                    CreateDirectoryToSaveImage(Chart3);
                    if (ConfigurationManager.AppSettings["SendMail"] == "1")
                        SendMail(Chart3);
                }
            }
        }

        public void SaveDetailImage()
        {
            try
            {
                Random random = new Random();
                List<string> paths = new List<string>();
                if (!Directory.Exists("D:\\ChartImages"))
                {
                    Directory.CreateDirectory("D:\\ChartImages");
                    string fileName1 = string.Format("D:\\ChartImages\\{0} {1} Chart Image {2}.png", DDTest.SelectedValue, DDBatch.SelectedValue, random.Next());
                    string fileName2 = string.Format("D:\\ChartImages\\{0} {1} Chart Image {2}.png", DDTest.SelectedValue, DDBatch.SelectedValue, random.Next());
                    Chart1.SaveImage(fileName1, ChartImageFormat.Png);
                    Chart2.SaveImage(fileName2, ChartImageFormat.Png);
                    paths.Add(fileName1);
                    paths.Add(fileName2);
                }
                else
                {
                    string fileName1 = string.Format("D:\\ChartImages\\{0} {1} Chart Image {2}.png", DDTest.SelectedValue, DDBatch.SelectedValue, random.Next());
                    string fileName2 = string.Format("D:\\ChartImages\\{0} {1} Chart Image {2}.png", DDTest.SelectedValue, DDBatch.SelectedValue, random.Next());
                    Chart1.SaveImage(fileName1, ChartImageFormat.Png);
                    Chart2.SaveImage(fileName2, ChartImageFormat.Png);
                    paths.Add(fileName1);
                    paths.Add(fileName2);
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    using (ZipArchive archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                    {
                        foreach (string path in paths)
                        {
                            ZipArchiveEntry entry = archive.CreateEntry(Path.GetFileName(path));
                            using (FileStream fileStream = File.OpenRead(path))
                            using (Stream entryStream = entry.Open())
                            {
                                fileStream.CopyTo(entryStream);
                            }
                        }
                    }
                    ms.Seek(0, SeekOrigin.Begin);

                    Response.Clear();
                    Response.ContentType = "application/zip";
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + DDAttempts.SelectedValue + " Detail of " + DDTest.SelectedValue + ".zip");
                    ms.CopyTo(Response.OutputStream);
                    // Response.End();
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Clear();
                Response.Redirect("Error.html");
            }
        }

        public void CreateDirectoryToSaveImage(Chart chart)
        {
            try
            {
                if (!Directory.Exists("D:\\ChartImages"))
                {
                    Directory.CreateDirectory("D:\\ChartImages");
                    Random random = new Random();
                    string fileName = string.Format("D:\\ChartImages\\{0} {1} Chart Image {2}.png", DDTest.SelectedValue, DDBatch.SelectedValue, random.Next());

                    chart.SaveImage(fileName, ChartImageFormat.Png);

                    FileInfo file = new FileInfo(fileName);
                    Response.Clear();
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + file.Name);
                    Response.AddHeader("Content-Length", file.Length.ToString());
                    Response.ContentType = "application/octet-stream";
                    Response.WriteFile(file.FullName);
                    Response.Flush();
                    Response.Clear();
                    Response.Close();
                }
                else
                {
                    Random random = new Random();
                    string fileName = string.Format("D:\\ChartImages\\{0} {1} Chart Image {2}.png", DDTest.SelectedValue, DDBatch.SelectedValue, random.Next());

                    chart.SaveImage(fileName, ChartImageFormat.Png);

                    FileInfo file = new FileInfo(fileName);
                    Response.Clear();
                    Response.AddHeader("Content-Disposition", "attachment; filename=" + file.Name);
                    Response.AddHeader("Content-Length", file.Length.ToString());
                    Response.ContentType = "application/octet-stream";
                    Response.WriteFile(file.FullName);
                    Response.Flush();
                    Response.Clear();
                    Response.Close();
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Clear();
                Response.Redirect("Error.html");
            }
        }

        public void SendMail(Chart chart)
        {
            var Email = System.Environment.GetEnvironmentVariable("MailID");
            var Password = System.Environment.GetEnvironmentVariable("MailPassword");

            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtp = new SmtpClient("smtp.gmail.com");

                mail.From = new MailAddress(Email, "Skill Ladder");
                mail.To.Add("gururaj.kl@sasken.com");
                mail.IsBodyHtml = true;

                MemoryStream s = new MemoryStream();
                chart.SaveImage(s, ChartImageFormat.Png);
                s.Position = 0;
                Attachment attach = new Attachment(s, "chart.png");
                mail.Attachments.Add(attach);

                mail.Subject = string.Format("{0} Chart Generated", DDBatch.SelectedValue);
                mail.Body = string.Format("<h1>{0} of {1} Chart Generated</h1><p>Chart attachment is included in this mail</p>", DDTest.SelectedValue, DDBatch.SelectedValue);

                smtp.Port = 587;
                smtp.Credentials = new System.Net.NetworkCredential(Email, Password);
                smtp.EnableSsl = true;

                smtp.Send(mail);
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
        }

        public void Alert(string message)
        {
            try
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.Append("<script type = 'text/javascript'>");
                sb.Append("window.onload=function(){");
                sb.Append("alert('");
                sb.Append(message);
                sb.Append("')};");
                sb.Append("</script>");
                ClientScript.RegisterClientScriptBlock(this.GetType(), "alert", sb.ToString());
            }
            catch (Exception ex)
            {
                ExceptionLogging.LogError(ex);
                Response.Redirect("Error.html");
            }
        }

        protected void Logout(object sender, EventArgs e)
        {
            FormsAuthentication.SignOut();
            Response.Redirect("/Login/Index");
        }
    }
}