using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.SqlClient;

namespace ANALYTICS_KPI_FORMAPP
{
    public partial class Form1 : Form
    {
    
        
        //lbl_series.Text = NormTable(dt);// StDev(y).ToString() + "--" + Mean(y).ToString();
 


        public Form1()
        {
            InitializeComponent();
            Sifreler();
            MultiLableChart(" Select * from ANALYTICS_KPI_BASIS_OS ", "_BASIS_OS","ÖDEME SÖZÜ ORANI");
            MultiLableChart(" Select * from ANALYTICS_KPI_BASIS_YY ", "_BASIS_YY", "YAPILANDIRMA ORANI");
            MultiLableChart(" Select * from ANALYTICS_KPI_BASIS_AE ", "_BASIS_AE", "AKSIYON ERTELEME ORANI");
            LineChartWithTrend(" Select * from ANALYTICS_KPI_RPC_SKOR ORDER BY ACTION_DT ", "_RPC_SKOR", "ILK SIRADA YUKLENEN TELEFONUNUN ULAŞMA SKORU ORTALAMASI");

            ReturnTable();
            SendMail();
        }
        private static string GUN = DateTime.Now.ToString("yyyyMMdd");
        private static string PASS;
        private static string USR;
        public static string savePath = @"\\btprdout01\output\OUTPUT\Perakende\COLLECTION\ANALYTICS_KPI\";
        public static string RetTable = "";

        public void ReturnTable()
        {
            RetTable = " ";

            RetTable +=  NormTable("Select * from ANALYTICS_KPI_BASIS_OS")+"<hr>";
            //RetTable += NormTable("Select * from ANALYTICS_KPI_BASIS_YY") + "<hr>"; 
        }
        static void SendMail()
        {

            //Bağlantımı oluşturuyorum
            SqlConnection conn = new SqlConnection("Data Source=.;Initial Catalog=.;User ID=.;Password=.;");
            //Komutumu yazıyorum.
            SqlCommand comm = new SqlCommand(@"EXEC  ANALYTICS_KPI_SP_FIN '" + RetTable + "','" + GUN + "'", conn);
            //Bağlantımı açıyorum.
            conn.Open();
            //SqlDataReader nesnem..
            comm.ExecuteNonQuery();


            conn.Close();

            // string path = localDestnDir + "\\DENEME.sql";
        } 

        static void Sifreler()
        {
            USR = "A25318";
            //Bağlantımı oluşturuyorum
            SqlConnection conn = new SqlConnection("Data Source=.;Initial Catalog=.;User ID=.;Password=.;");
            //Komutumu yazıyorum.
            SqlCommand comm = new SqlCommand(@"  select  ConfiguredValue    FROM [CCOps].[dbo].[SSIS Configurations]   where [PackagePath]='\Package.Connections[edw.finansbank.com.tr.A25318].Properties[Password]' and [ConfigurationFilter]='IVNEOD'", conn);
            //Bağlantımı açıyorum.
            conn.Open();
            //SqlDataReader nesnem..
            SqlDataReader dr = comm.ExecuteReader();
            while (dr.Read())
            {
                PASS = dr["ConfiguredValue"].ToString().Trim();
            }
            //sqlDataReader ve SqlConnection kapatılıyor.
            dr.Close();
            conn.Close();

            // string path = localDestnDir + "\\DENEME.sql";
        }


        public static  DataTable OraDt(string cmdstr)
        {
            string constr = "Data Source=.;User Id=" + USR + ";Password=" + PASS + ";Integrated Security=no;";


            // Create the adapter with the selectCommand txt and the
            // connection string
            OracleDataAdapter adapter = new OracleDataAdapter(cmdstr, constr);

            // Create the builder for the adapter to automatically generate
            // the Command when needed
            OracleCommandBuilder builder = new OracleCommandBuilder(adapter);

            // Create and fill the DataSet using the EMP
            DataSet dataset = new DataSet();
            adapter.Fill(dataset, "EMP");

            // Get the EMP table from the dataset
            DataTable dt = dataset.Tables["EMP"];
            return dt;
        }
        private void LineChartWithTrend(string cmdstr, string Cnm, string baslik)
        {
      
            DataTable dt = OraDt(cmdstr);

            //string[] x = new string[dt.Rows.Count];
            //int[] y = new int[dt.Rows.Count];
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    x[i] = dt.Rows[i][0].ToString();
            //    y[i] = Convert.ToInt32(dt.Rows[i][1]);
            //}


            Chart chart3 = new Chart()
            {
                Width = 1200,
                Height = 300
            };



            chart3.Legends.Add(new Legend() { Name = "Legend" });
           chart3.Legends[0].Docking = Docking.Bottom;
            ChartArea chartArea = new ChartArea() { Name = "ChartArea" };
            //Remove X-axis grid lines
            // chartArea.AxisX.MajorGrid.LineWidth = 0;
            ////Remove Y-axis grid lines
            //chartArea.AxisY.MajorGrid.LineWidth = 0;


           // YUZDE chartArea.AxisY.LabelStyle.Format = "{0;0}" + "%";
            //Chart Area Back Color
            chartArea.BackColor = Color.FromName("white");
            chart3.ChartAreas.Add(chartArea);
            chart3.Palette = ChartColorPalette.BrightPastel;

            var series = new Series("Cevap Verildi");
            series.ChartType = SeriesChartType.SplineArea;
            // series.ChartType = (SeriesChartType)14;
            //string name = dt.Rows.Count.ToString();

            string name = dt.Rows[0][0].ToString();


            foreach (DataRow dr in dt.Rows)
            {
                double y = (double)PCL.Utility.DBtoMT.ToDouble(dr[1]);
                series.Points.AddXY(dr["ACTION_DT"].ToString(), y);
                 series.LegendText = dr.Table.Columns[1].ColumnName;
            }

            //series.Points.DataBindXY(new[] { "2001-2", "2002-2", "2003-1", "2004-1" }, new[] { 10, 20, 90, 15 });
            //series.Points.DataBindXY(x, y);
       //     series.IsValueShownAsLabel = true;
            //series.IsVisibleInLegend = false;
         //YUZDE   series.LabelFormat = "{0;0}" + "%";

            chart3.Series.Add(series);
            //işe yarıyor  
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = 45; // this works
            /*chart3.ChartAreas[0].AxisX.IsLabelAutoFit = true;
            chart3.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.LabelsAngleStep30;
            chart3.ChartAreas[0].AxisX.LabelStyle.Enabled = true;
            */
            // chart3.Series[0].XValueType = ChartValueType.DateTime;
            //chart3.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy/MM/dd";
      chart3.Series[0].Sort(PointSortOrder.Ascending, "X");

            //chart3.ChartAreas[0].AxisX.Interval = 1;
            //chart3.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
            //chart3.ChartAreas[0].AxisX.IntervalOffset = 1;

            //trendline

            chart3.Series.Add("TrendLine");
            chart3.Series["TrendLine"].ChartType = SeriesChartType.Line;
            chart3.Series["TrendLine"].BorderWidth = 1;
            chart3.Series["TrendLine"].Color = Color.Red;
            // Line of best fit is linear
            string typeRegression = "Linear";//"Exponential";//
            // The number of days for Forecasting
            string forecasting = "1";
            // Show Error as a range chart.
            string error = "false";
            // Show Forecasting Error as a range chart.
            string forecastingError = "false";
            // Formula parameters
            string parameters = typeRegression + ',' + forecasting + ',' + error + ',' + forecastingError;
            chart3.Series[0].Sort(PointSortOrder.Ascending, "X");
            // Create Forecasting Series.
            chart3.DataManipulator.FinancialFormula(FinancialFormula.Forecasting, parameters, chart3.Series[0], chart3.Series["TrendLine"]);
            /* */



            chart3.Titles.Add(baslik + " (" + name + " Tarihi Sonrası)");
            //chart3.Titles.Add(sira + ". Sırada \"Cevap Verildi\" Alınan Telefonların Trendi ( " + name + "  Tarihi Sonrası)");
            //chart3.ChartAreas[0].Area3DStyle.Enable3D = true;
           // chart3.SaveImage(savePath + GUN + "_CHART_TREND" + sira + ".png", ChartImageFormat.Png);
            chart3.SaveImage(savePath + GUN + Cnm + ".png", ChartImageFormat.Png);
        }
        public void MultiLableChart(string cmdstr, string Cnm,string baslik)
        {

         
            DataTable dt = OraDt(cmdstr);




            Chart chart3 = new Chart()
            {
                Width = 1200,
                Height = 300
            };
            // chart3.



            chart3.Legends.Add(new Legend() { Name = "Legend" });
            chart3.Legends[0].Docking = Docking.Bottom;
            ChartArea chartArea = new ChartArea() { Name = "ChartArea" };
            //Remove X-axis grid lines
            //chartArea.AxisX.MajorGrid.LineWidth = 0;
            ////Remove Y-axis grid lines
            //chartArea.AxisY.MajorGrid.LineWidth = 0;


            chartArea.AxisY.LabelStyle.Format = "{0;0}" + "%";
            //Chart Area Back Color
            chartArea.BackColor = Color.FromName("white");
            chart3.ChartAreas.Add(chartArea);
            chart3.Palette = ChartColorPalette.BrightPastel;

            string name = dt.Rows[0][0].ToString();


            for (int i = 1; i < dt.Columns.Count; i++)
            {
                Series series = new Series();
                foreach (DataRow dr in dt.Rows)
                {
                    double y = (double)PCL.Utility.DBtoMT.ToDouble(dr[i]);
                    series.Points.AddXY(dr["Tarih"].ToString(), y);
                    series.LegendText = dr.Table.Columns[i].ColumnName;
                }
                chart3.Series.Add(series);
                chart3.Series[0].BorderWidth = 1;
                // Chart3.Series[i].ChartType = SeriesChartType.Line;
                // series.ChartType = SeriesChartType.StackedArea;
                series.ChartType = SeriesChartType.Line;
            }
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = 45; // this works

            chart3.Titles.Add(baslik+" (" + name + " Tarihi Sonrası)");
          //  chart3.ChartAreas[0].Area3DStyle.Enable3D = true;
            chart3.SaveImage(savePath + GUN + Cnm+".png", ChartImageFormat.Png);
        }

   
    public static double[] Seri(DataTable dt, string Column)
    {
        double[] y = new double[dt.Rows.Count];
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            y[i] = PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
        }
        return y;
    }
    public static int MaxValue(DataTable dt)
    {
        return dt.Rows[dt.Rows.Count - 1].Field<int>(0);

    }
    public static double DegerValue(DataTable dt, string Column, int Tarih)
    {

        double y = new double();
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            if (dt.Rows[i][0].ToString() == Tarih.ToString())
                y = PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
        }
        return y;
    }
    public static string Trend(double v, double o, double s)
    {
        string r = "0";
        if (v > o + s)
            r = "2";
        if (v < o - s)
            r = "1";
        return r;
    }
    public static string NormTable(string cmdstr)
    {

        DataTable dt = OraDt(cmdstr);
        int MaxTarih=0;
        string rv = "<table><tr><th></th><th>Değer (" + MaxTarih.ToString() + ")</th><th>Ortalama</th><th>St. Sapma</th><th>Trend</th></tr>";
  
 
      MaxTarih = Convert.ToInt32(dt.Compute("max(Tarih)", string.Empty));
            
 
        catch (Exception ex) {
            rv = ex.ToString();

        }

        for (int i = 1; i < dt.Columns.Count; i++)
        {
            string name = dt.Columns[i].ColumnName.ToString();
            double[] s = Seri(dt, name);
            double Deger = DegerValue(dt, name, MaxTarih);
            double Ortalama = Math.Round(Mean(s), 2);
            double StandartSapma = Math.Round(StDev(s), 2);
            rv += "<tr><td>" + name + "</td><td>" + Deger + "</td><td>" + Ortalama.ToString() + "</td><td>" + StandartSapma.ToString() + "</td><td>"+Trend(Deger,Ortalama,StandartSapma)+"</td></tr>";

        }

        return rv + "</table>";
    }





    /// <summary>
    /// Calculates the mean of an array of values
    /// </summary>
    /// <param name="v">the array of values to calculate their mean</param>
    /// <returns>The mean of the array of values</returns>
    public static double Mean(double[] v)
    {
        double sum = 0.0;

        for (int i = 0; i < v.Length; i++)
        {
            sum += v[i];
        }

        return sum / v.Length;
    }

    /// <summary>
    /// Calculates the variance of an array of values
    /// </summary>
    /// <param name="v">the array of values to calculate their variance</param>
    /// <returns>The variance of the array of values</returns>
    public static double Variance(double[] v)
    {
        double mean = Mean(v);
        double sum = 0.0;

        for (int i = 0; i < v.Length; i++)
        {
            sum += (v[i] - mean) * (v[i] - mean);
        }

        int denom = v.Length - 1;
        if (v.Length <= 1)
            denom = v.Length;

        return sum / denom;
    }

    /// <summary>
    /// Calculates the standard deviation of an array of values
    /// </summary>
    /// <param name="v">the array of values to calculate their standard deviation</param>
    /// <returns>The standard deviation of the array of values</returns>
    public static double StDev(double[] v)
    {
        return Math.Sqrt(Variance(v));
    }




    }
}

