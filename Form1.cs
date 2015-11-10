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

            //ARAMA EKIPLERI
            MultiLableChart(" select * from ANALYTICS_KPI_GECIKMELI_STG ", "9", "1", "_GECIKMELI_STG", "STRATEJI", "STG", " gelen gelmeyen");
            MultiLableChart(" select * from ANALYTICS_KPI_GENELARANAN ", "9", "1", "_GENELARANAN", "ARAMA EKIPLERI", "GENEL ARANMA TAKIBI", "Yüklenenler datanın ne kadarı arandı, ne kadarına ulaşıldığı gibi genel aranma durumunu gösterir. Detay rapora <a href=\"\\\\BTPRDOUT01\\OUTPUT\\OUTPUT\\Perakende\\COLLECTION\\Pred-Prew ilişkin Raporlar\\Pred-Prev Raporu\">buradan</a> ulaşabilirsiniz.");

            MultiLableChart(" select * from ANALYTICS_KPI_REVENUE ", "9", "1", "_REVENUE", "TOPLAM FAYDA", "REVENUE", " Ulaşma X OS Oranı X Kept Oranı");


            //STG
            MultiLableChart(" select * from ANALYTICS_KPI_HAVUZLANMA_ADET ", "%", "1", "_HAVUZLANMA_ADET", "STRATEJI", "GECIKME DONEMI BAZINDA HAVUZLANAN ADET", "Gecikme dönemleri bazında strateji tarafından havuzlanma adetlerinin oranlarıdır. Detay rapora <a href=\"\\\\BTPRDOUT01\\OUTPUT\\OUTPUT\\Perakende\\COLLECTION\\Gecikme - Havuzlanma - Havuzlanmama Raporları\">buradan</a> ulaşabilirsiniz.");
            MultiLableChart(" select * from ANALYTICS_KPI_HAVUZLANMA_BAL ", "%", "1", "_HAVUZLANMA_BAL", "STRATEJI", "GECIKME DONEMI BAZINDA HAVUZLANAN BAKIYE", "Gecikme dönemleri bazında strateji tarafından havuzlanma bakiyelerinin oranlarıdır. Detay rapora <a href=\"\\\\BTPRDOUT01\\OUTPUT\\OUTPUT\\Perakende\\COLLECTION\\Gecikme - Havuzlanma - Havuzlanmama Raporları\">buradan</a> ulaşabilirsiniz.");


            MultiLableChartWithTrend(" Select * from ANALYTICS_KPI_BT_LIMIT ORDER BY ACTION_DT ", "9", "1", "_BT_LIMIT_ADET", "STRATEJI", "YY LIMIT TANIMLI ADET", "BT YY Limit tanımlı ürün adetleri.");
            MultiLableChartWithTrend(" Select * from ANALYTICS_KPI_KMH_LIMIT ORDER BY ACTION_DT ", "9", "1", "_KMH_LIMIT_ADET", "STRATEJI", "YY LIMIT TANIMLI ADET", "KMH YY Limit tanımlı ürün adetleri.");

            //YUKLENEN

            MultiLableChart(" select * from ANALYTICS_KPI_YUKLENENADET ", "9", "1", "_YUKLENENADET", "YUKLENEN DATA", "EKIP BAZINDA YUKLENEN ADET", "Ekip bazında yüklenen adetleri içerir. Detay rapora <a href=\"\\\\BTPRDOUT01\\OUTPUT\\OUTPUT\\Perakende\\COLLECTION\\Pred-Prew ilişkin Raporlar\\Pred-Prev Raporu\">buradan</a> ulaşabilirsiniz.");




            MultiLableChart(" Select * from ANALYTICS_KPI_BASIS_OS ", "%", "1", "_BASIS_OS", "OS-KEPT TAKIP", "ÖDEME SÖZÜ ORANI", "RPC olan kayıtların Ödeme Sözü alınma oranları");
            MultiLableChart(" Select * from ANALYTICS_KPI_BASIS_YY ", "%", "1", "_BASIS_YY", "OS-KEPT TAKIP", "YAPILANDIRMA ORANI", "RPC olan kayıtların Ödeme Planı alınma oranları");
            MultiLableChart(" Select * from ANALYTICS_KPI_BASIS_AE ", "%", "0", "_BASIS_AE", "OS-KEPT TAKIP", "AKSIYON ERTELEME ORANI", "RPC olan kayıtların Aksiyon Erteleme alınma oranları");
            MultiLableChart(" Select * from ANALYTICS_KPI_BEKGER_KEPT ", "%", "1", "_BEKGER_KEPT", "OS-KEPT TAKIP", "KEPT TAKIP", "Ekip bazında kept takibi");
            MultiLableChart(" Select * from ANALYTICS_KPI_BEKGER_ADET ", "%", "1", "_BEKGER_ADET", "OS-KEPT TAKIP", "BEKLENEN GERÇEKLEŞEN -ADET", "Ekip bazında beklenen gerçekleşen adetleri");
            MultiLableChart(" Select * from ANALYTICS_KPI_BEKGER_BAKIYE ", "%", "1", "_BEKGER_BAL", "OS-KEPT TAKIP", "BEKLENEN GERÇEKLEŞEN -BAKIYE", "Ekip bazında beklenen gerçekleşen bakiyeleri");

            LineChartWithTrend(" Select ACTION_DT , RPC_SKOR from ANALYTICS_KPI_RPC_SKOR ORDER BY ACTION_DT ", "%", "1", "_RPC_SKOR", "ULAŞMA TAKIP", "ILK SIRADA YUKLENEN TELEFONUNUN ULAŞMA SKORU ORTALAMASI", "Ulaşma-Fayda arama modeli kapsamında ilk sırada yüklenen telefon numaralarının skor durumu");


            SendMail();
        }
        private static string GUN = DateTime.Now.ToString("yyyyMMdd");
        private static string PASS;
        private static string USR;
        public static string savePath = @"..\ANALYTICS_KPI\";
        public static string RetTable = "";


        #region çoğul
        public static string NormTables(DataTable dt, string file, string baslik, string trend)
        {

            //     DataTable dt = OraDt(cmdstr);
            int MaxTarih = 0;
            try
            {

                MaxTarih = Convert.ToInt32(dt.Compute("max(Tarih)", string.Empty));
            }
            catch
            //
            {
                MaxTarih = Convert.ToInt32(dt.Compute("max(ACTION_DT)", string.Empty));
            }

            string rb = "";
            string rv = "";
            DataView view = new DataView(dt);
            DataTable distinctValues = view.ToTable(true, "GGS_BAND");




            for (int i = 2; i < dt.Columns.Count; i++)
            {
                for (int n = 0; n < distinctValues.Rows.Count; n++)
                {

                    string distinctName = distinctValues.Rows[n][0].ToString();
                    string Column = dt.Columns[i].ColumnName.ToString();
                    rb += Degeri(dt, Column, distinctName, MaxTarih, trend);
                }
            }

            return rb;
        }
        public static string Degeri(DataTable dt, string Column, string distinctName, int MaxTarih, string trend)
        {
            double[] s = Seris(dt, Column, distinctName);
            double Deger = DegerValues(dt, Column, MaxTarih, distinctName);
            //double DegerAylik = Math.Round(DegerValueAylik(dt, Column, MaxTarih), 2);
            double Ortalama = Math.Round(Mean(s), 2);
            double StandartSapma = Math.Round(StDev(s), 2);

            return " <table><tr><td>" + Deger + " " + Trend(Deger, Ortalama, StandartSapma, trend) + "col:" + Column + " discol:" + distinctName + " </td></tr></table> ";

        }
        public static double[] Seris(DataTable dt, string Column, string distinctName)
        {
            int Y = 0;
            int X = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][1].ToString() == distinctName)
                    Y++;
            }

            double[] y = new double[Y];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][1].ToString() == distinctName)
                {

                    y[X] = PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
                    X++;
                }
            }
            return y;
        }
        public static double DegerValues(DataTable dt, string Column, int Tarih, string distinctName)
        {

            double y = new double();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString() == Tarih.ToString() && dt.Rows[i][1].ToString() == distinctName)
                    y = PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
            }
            return y;
        }
        #endregion



        private void LineChartWithTrend(string cmdstr, string type, string trend, string Cnm, string parent, string baslik, string aciklama)
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
                series.Points.AddXY(dr[0].ToString(), y);
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

            chart3.ChartAreas[0].AxisX.Interval = 10;
            //  chart3.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
            chart3.ChartAreas[0].AxisX.IntervalOffset = 1;

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

            string file = savePath + GUN + Cnm + ".png";
            chart3.SaveImage(file, ChartImageFormat.Png);

            RetTable += NormTable(dt, file, parent, trend, baslik, aciklama);
        }
        public void MultiLableChartWithTrend(string cmdstr, string type, string trend, string Cnm, string parent, string baslik, string aciklama)
        {
            DataTable dt = OraDt(cmdstr);
            Chart chart3 = new Chart()
            {
                Width = 1200,
                Height = 300
            };
            chart3.Legends.Add(new Legend() { Name = "Legend" });
            chart3.Legends[0].Docking = Docking.Bottom;
            ChartArea chartArea = new ChartArea() { Name = "ChartArea" };

            if (type == "%")
                chartArea.AxisY.LabelStyle.Format = "{0;0}" + "%";
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
                    series.Points.AddXY(dr[0].ToString(), y);
                    if (i == 1)
                    {
                        series.LegendText = dr.Table.Columns[i].ColumnName + " (Primary)";
                        series.YAxisType = AxisType.Primary;
                    }
                    else
                    {
                        series.LegendText = dr.Table.Columns[i].ColumnName + " (Secondary)";
                        series.YAxisType = AxisType.Secondary;
                    }
                }
                chart3.Series.Add(series);
                chart3.Series[i - 1].BorderWidth = 3;
                series.ChartType = SeriesChartType.SplineArea;

            }

            //
            // chart3.Series[0].YAxisType = AxisType.Primary;
            // chart3.Series[1].YAxisType = AxisType.Secondary;

            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = 45; // this works
            chart3.ChartAreas[0].AxisX.Interval = 10;
            //  chart3.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
            chart3.ChartAreas[0].AxisX.IntervalOffset = 1;



            chart3.Series[0].Sort(PointSortOrder.Ascending, "X");

            chart3.Series.Add("TrendLine");
            chart3.Series["TrendLine"].ChartType = SeriesChartType.Line;
            chart3.Series["TrendLine"].BorderWidth = 1;
            chart3.Series["TrendLine"].Color = Color.Red;
           // chart3.Series["TrendLine"].LabelAngle=

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
            string file = savePath + GUN + Cnm + ".png";
            chart3.SaveImage(file, ChartImageFormat.Png);

            RetTable += NormTable(dt, file, parent, trend, baslik, aciklama);
        }
        public void MultiLableChart(string cmdstr, string type, string trend, string Cnm, string parent, string baslik, string aciklama)
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

            if (type == "%")
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
                chart3.Series[i - 1].BorderWidth = 3;
                // Chart3.Series[i].ChartType = SeriesChartType.Line;
                // series.ChartType = SeriesChartType.StackedArea;
                series.ChartType = SeriesChartType.Line;

            }
            chart3.ChartAreas[0].AxisX.LabelStyle.Angle = 45; // this works
            chart3.ChartAreas[0].AxisX.Interval = 10;
            chart3.ChartAreas[0].AxisX.IntervalOffset = 1;

            //    chart3.Series[0].BorderWidth = 2;
            chart3.Titles.Add(baslik + " (" + name + " Tarihi Sonrası)");
            //  chart3.ChartAreas[0].Area3DStyle.Enable3D = true;
            string file = savePath + GUN + Cnm + ".png";
            chart3.SaveImage(file, ChartImageFormat.Png);

            RetTable += NormTable(dt, file, parent, trend, baslik, aciklama);
        }

        public static string NormTable(DataTable dt, string file, string parent, string trend, string baslik, string aciklama)
        {

            //     DataTable dt = OraDt(cmdstr);
            int MaxTarih = 0;
            try
            {

                MaxTarih = Convert.ToInt32(dt.Compute("max(Tarih)", string.Empty));
            }
            catch
            //
            {
                MaxTarih = Convert.ToInt32(dt.Compute("max(ACTION_DT)", string.Empty));
            }
            string rv = "<h3>(<u>" + parent + "</u>) " + baslik + "</h3><font color=\"GRAY\">" + aciklama + "</font><table><tr><td><table class=\"NormTable\"><tr><th></th><th>Değer (" + MaxTarih.ToString() + ")</th><th>Ay Ortalaması</th><th>3 Aylık Ortalama</th><th>3-6 Aylık Ortalama</th><th>St. Sapma</th><th>Trend</th><th>Aylık Trend</th></tr>";


            for (int i = 1; i < dt.Columns.Count; i++)
            {
                string name = dt.Columns[i].ColumnName.ToString();
                //double[] s = Seri(dt, name);
                double[] s = Seri(dt, name, MaxTarih,3);
                double Deger = DegerValue(dt, name, MaxTarih);
                double DegerAylik = Math.Round(DegerValueAylik(dt, name, MaxTarih), 2);
                double Deger6Aylik = Math.Round(DegerValue6Aylik(dt, name, MaxTarih), 2);
                double Deger3Aylik = Math.Round(Mean(s), 2);
                double StandartSapma = Math.Round(StDev(s), 2);
                rv += "<tr><td>" + name + "</td><td>" + Deger + "</td><td>" + DegerAylik + "</td><td>" + Deger3Aylik.ToString() + "</td><td>" + Deger6Aylik.ToString() + "</td><td>" + StandartSapma.ToString() + "</td><td><img width=\"12\" height=\"12\" src=\"" + savePath + "trend" + Trend(Deger, Deger3Aylik, StandartSapma, trend) + ".png\"></td><td><img width=\"12\" height=\"12\" src=\"" + savePath + "trend" + Trend(DegerAylik, Deger3Aylik, StandartSapma, trend) + ".png\"></td></tr>";

            }

            return rv + "</table></td><td><img src=\"" + file + "\"></td></tr></table>";
        }

        /// <summary>
        /// 3-6 aylık ortalama. mevcut ay dahil değil, son ay-3 ay dahil değil
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Column"></param>
        /// <param name="Tarih"></param>
        /// <returns></returns>
        public static double DegerValue6Aylik(DataTable dt, string Column, int Tarih)
        {
            string newDt = Tarih.ToString().Substring(0, 6);
            int iDt = (int.Parse(newDt) - 3);
            double y = new double();
            y = 0;
            double X = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                int iDx = int.Parse(dt.Rows[i][0].ToString().Substring(0, 6));
                if (iDx < iDt)
                {
                    y += PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
                    X++;
                }
            }
            return y / X;
        }
        public static double DegerValueAylik(DataTable dt, string Column, int Tarih)
        {
            string newDt = Tarih.ToString().Substring(0, 6);
            double y = new double();
            y = 0;
            double X = 0;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString().Substring(0, 6) == newDt)
                {
                    y += PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
                    X++;
                }
            }
            return y / X;
        }

        static void SendMail()
        {

            //Bağlantımı oluşturuyorum
            SqlConnection conn = new SqlConnection("Data Source=.;Initial Catalog=**;User ID=**;Password=**;");
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
            USR = "111";
            //Bağlantımı oluşturuyorum
            SqlConnection conn = new SqlConnection("Data Source=.;Initial Catalog=*;User ID=*;Password=*;");
            //Komutumu yazıyorum.
            SqlCommand comm = new SqlCommand(@"  select  ConfiguredValue    FROM [CCOps].[dbo].[SSIS Configurations]   where [PackagePath]='\Package.Connections[edw].Properties[Password]' and [ConfigurationFilter]='IVNEOD'", conn);
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


        public static DataTable OraDt(string cmdstr)
        {
            string constr = "Data Source=***;User Id=" + USR + ";Password=" + PASS + ";Integrated Security=no;";


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
        //  private void LineChartWithTrend(string cmdstr, string Cnm, string baslik)
        //  {

        //      DataTable dt = OraDt(cmdstr);

        //      //string[] x = new string[dt.Rows.Count];
        //      //int[] y = new int[dt.Rows.Count];
        //      //for (int i = 0; i < dt.Rows.Count; i++)
        //      //{
        //      //    x[i] = dt.Rows[i][0].ToString();
        //      //    y[i] = Convert.ToInt32(dt.Rows[i][1]);
        //      //}


        //      Chart chart3 = new Chart()
        //      {
        //          Width = 1200,
        //          Height = 300
        //      };



        //      chart3.Legends.Add(new Legend() { Name = "Legend" });
        //     chart3.Legends[0].Docking = Docking.Bottom;
        //      ChartArea chartArea = new ChartArea() { Name = "ChartArea" };
        //      //Remove X-axis grid lines
        //      // chartArea.AxisX.MajorGrid.LineWidth = 0;
        //      ////Remove Y-axis grid lines
        //      //chartArea.AxisY.MajorGrid.LineWidth = 0;


        //     // YUZDE chartArea.AxisY.LabelStyle.Format = "{0;0}" + "%";
        //      //Chart Area Back Color
        //      chartArea.BackColor = Color.FromName("white");
        //      chart3.ChartAreas.Add(chartArea);
        //      chart3.Palette = ChartColorPalette.BrightPastel;

        //      var series = new Series("Cevap Verildi");
        //      series.ChartType = SeriesChartType.SplineArea;
        //      // series.ChartType = (SeriesChartType)14;
        //      //string name = dt.Rows.Count.ToString();

        //      string name = dt.Rows[0][0].ToString();


        //      foreach (DataRow dr in dt.Rows)
        //      {
        //          double y = (double)PCL.Utility.DBtoMT.ToDouble(dr[1]);
        //          series.Points.AddXY(dr["ACTION_DT"].ToString(), y);
        //           series.LegendText = dr.Table.Columns[1].ColumnName;
        //      }

        //      //series.Points.DataBindXY(new[] { "2001-2", "2002-2", "2003-1", "2004-1" }, new[] { 10, 20, 90, 15 });
        //      //series.Points.DataBindXY(x, y);
        // //     series.IsValueShownAsLabel = true;
        //      //series.IsVisibleInLegend = false;
        //   //YUZDE   series.LabelFormat = "{0;0}" + "%";

        //      chart3.Series.Add(series);
        //      //işe yarıyor  
        //      chart3.ChartAreas[0].AxisX.LabelStyle.Angle = 45; // this works
        //      /*chart3.ChartAreas[0].AxisX.IsLabelAutoFit = true;
        //      chart3.ChartAreas[0].AxisX.LabelAutoFitStyle = LabelAutoFitStyles.LabelsAngleStep30;
        //      chart3.ChartAreas[0].AxisX.LabelStyle.Enabled = true;
        //      */
        //      // chart3.Series[0].XValueType = ChartValueType.DateTime;
        //      //chart3.ChartAreas[0].AxisX.LabelStyle.Format = "yyyy/MM/dd";
        //chart3.Series[0].Sort(PointSortOrder.Ascending, "X");

        //      //chart3.ChartAreas[0].AxisX.Interval = 1;
        //      //chart3.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Months;
        //      //chart3.ChartAreas[0].AxisX.IntervalOffset = 1;

        //      //trendline

        //      chart3.Series.Add("TrendLine");
        //      chart3.Series["TrendLine"].ChartType = SeriesChartType.Line;
        //      chart3.Series["TrendLine"].BorderWidth = 1;
        //      chart3.Series["TrendLine"].Color = Color.Red;
        //      // Line of best fit is linear
        //      string typeRegression = "Linear";//"Exponential";//
        //      // The number of days for Forecasting
        //      string forecasting = "1";
        //      // Show Error as a range chart.
        //      string error = "false";
        //      // Show Forecasting Error as a range chart.
        //      string forecastingError = "false";
        //      // Formula parameters
        //      string parameters = typeRegression + ',' + forecasting + ',' + error + ',' + forecastingError;
        //      chart3.Series[0].Sort(PointSortOrder.Ascending, "X");
        //      // Create Forecasting Series.
        //      chart3.DataManipulator.FinancialFormula(FinancialFormula.Forecasting, parameters, chart3.Series[0], chart3.Series["TrendLine"]);
        //      /* */



        //      chart3.Titles.Add(baslik + " (" + name + " Tarihi Sonrası)");
        //      //chart3.Titles.Add(sira + ". Sırada \"Cevap Verildi\" Alınan Telefonların Trendi ( " + name + "  Tarihi Sonrası)");
        //      //chart3.ChartAreas[0].Area3DStyle.Enable3D = true;
        //     // chart3.SaveImage(savePath + GUN + "_CHART_TREND" + sira + ".png", ChartImageFormat.Png);
        //      chart3.SaveImage(savePath + GUN + Cnm + ".png", ChartImageFormat.Png);
        //  }
        //  public void MultiLableChart(string cmdstr, string Cnm,string baslik)
        //  {


        //      DataTable dt = OraDt(cmdstr);




        //      Chart chart3 = new Chart()
        //      {
        //          Width = 1200,
        //          Height = 300
        //      };
        //      // chart3.



        //      chart3.Legends.Add(new Legend() { Name = "Legend" });
        //      chart3.Legends[0].Docking = Docking.Bottom;
        //      ChartArea chartArea = new ChartArea() { Name = "ChartArea" };
        //      //Remove X-axis grid lines
        //      //chartArea.AxisX.MajorGrid.LineWidth = 0;
        //      ////Remove Y-axis grid lines
        //      //chartArea.AxisY.MajorGrid.LineWidth = 0;


        //      chartArea.AxisY.LabelStyle.Format = "{0;0}" + "%";
        //      //Chart Area Back Color
        //      chartArea.BackColor = Color.FromName("white");
        //      chart3.ChartAreas.Add(chartArea);
        //      chart3.Palette = ChartColorPalette.BrightPastel;

        //      string name = dt.Rows[0][0].ToString();


        //      for (int i = 1; i < dt.Columns.Count; i++)
        //      {
        //          Series series = new Series();
        //          foreach (DataRow dr in dt.Rows)
        //          {
        //              double y = (double)PCL.Utility.DBtoMT.ToDouble(dr[i]);
        //              series.Points.AddXY(dr["Tarih"].ToString(), y);
        //              series.LegendText = dr.Table.Columns[i].ColumnName;
        //          }
        //          chart3.Series.Add(series);
        //          chart3.Series[0].BorderWidth = 1;
        //          // Chart3.Series[i].ChartType = SeriesChartType.Line;
        //          // series.ChartType = SeriesChartType.StackedArea;
        //          series.ChartType = SeriesChartType.Line;
        //      }
        //      chart3.ChartAreas[0].AxisX.LabelStyle.Angle = 45; // this works

        //      chart3.Titles.Add(baslik+" (" + name + " Tarihi Sonrası)");
        //    //  chart3.ChartAreas[0].Area3DStyle.Enable3D = true;
        //      chart3.SaveImage(savePath + GUN + Cnm+".png", ChartImageFormat.Png);
        //  }

/// <summary>
/// Max Tarih değeri ile gelen tarihin içindeki ayı almaz sonraki DahilAy kadar kabul eder
/// </summary>
/// <param name="dt"></param>
/// <param name="Column"></param>
/// <param name="Tarih"></param>
/// <param name="DahilAy"></param>
/// <returns></returns>
        public static double[] Seri(DataTable dt, string Column, int Tarih,int DahilAy)
        {
            int newDt =int.Parse(Tarih.ToString().Substring(0, 6));
            int iDt = int.Parse(newDt.ToString()) - DahilAy;


            int x = 0;
            int z = 0;
            for (x = 0; x < dt.Rows.Count; x++)
            {
                if (int.Parse(dt.Rows[x][0].ToString().Substring(0, 6)) != newDt && int.Parse(dt.Rows[x][0].ToString().Substring(0, 6)) > iDt)
                    x++;
            }

            double[] y = new double[x];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if ((int.Parse(dt.Rows[i][0].ToString().Substring(0, 6)) != newDt) && (int.Parse(dt.Rows[i][0].ToString().Substring(0, 6)) > iDt))
                {
                    y[z] = PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
                    z++;
                }
            }
            return y;

        }      
        /// <summary>
        /// Max Tarih değeri ile gelen tarihin içindeki ayı almaz
        /// </summary>
        public static double[] Seri(DataTable dt, string Column, int Tarih)
        {
            int x = 0;
            int z = 0;
            for (x = 0; x < dt.Rows.Count; x++)
            {
                if (dt.Rows[x][0].ToString().Substring(0, 6) != Tarih.ToString().Substring(0, 6))
                    x++;
            }

            double[] y = new double[x];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][0].ToString().Substring(0, 6) != Tarih.ToString().Substring(0, 6))
                {
                    y[z] = PCL.Utility.DBtoMT.ToDouble(dt.Rows[i][Column]);
                    z++;
                }
            }
            return y;

        }
        /// <summary>
        /// tarih kriteri koymadan tamamını alır
        /// </summary>
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
        /// <summary>
        /// data table içerisinde tarihe karşılık gelen değer
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="Column"></param>
        /// <param name="Tarih"></param>
        /// <returns></returns>
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
        public static string Trend(double v, double o, double s, string trend)
        {
            string r = "0";


            if (v > o + s)
                r = "2";
            if (v < o - s)
                r = "1";

            if (trend == "0")
            {
                if (v > o + s)
                    r = "1";
                if (v < o - s)
                    r = "2";
            }


            return r;
        }
        // public static string NormTable(string cmdstr)
        //{

        //    DataTable dt = OraDt(cmdstr);
        //    int MaxTarih=0;
        //    string rv = "<table><tr><th></th><th>Değer (" + MaxTarih.ToString() + ")</th><th>Ortalama</th><th>St. Sapma</th><th>Trend</th></tr>";


        //  MaxTarih = Convert.ToInt32(dt.Compute("max(Tarih)", string.Empty));


        //    catch (Exception ex) {
        //        rv = ex.ToString();

        //    }

        //    for (int i = 1; i < dt.Columns.Count; i++)
        //    {
        //        string name = dt.Columns[i].ColumnName.ToString();
        //        double[] s = Seri(dt, name);
        //        double Deger = DegerValue(dt, name, MaxTarih);
        //        double Ortalama = Math.Round(Mean(s), 2);
        //        double StandartSapma = Math.Round(StDev(s), 2);
        //        rv += "<tr><td>" + name + "</td><td>" + Deger + "</td><td>" + Ortalama.ToString() + "</td><td>" + StandartSapma.ToString() + "</td><td>"+Trend(Deger,Ortalama,StandartSapma)+"</td></tr>";

        //    }

        //    return rv + "</table>";
        //}





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

