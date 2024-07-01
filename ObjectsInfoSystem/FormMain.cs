using rtp3esh_bd;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

//using lcpi.data.oledb;
using FirebirdSql.Data.FirebirdClient;

namespace ObjectsInfoSystem
{
    public partial class FormMain : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        //bool permissionAdmin; // признак прав доступа: 100 - администратор, 1 - пользователь, 2 -диспетчер-редактор
        int permissionAdmin; // признак прав доступа: 100 - администратор, 1 - пользователь, 2 -диспетчер-редактор

        public string dbconnectionString;
        public string dbconnectionStringBlackouts;
        public string dbconnectionStringIESBK;
        public string dbconnectionStringIESBK2;
        public static DataTable tablelsPropGLOBAL;

        public FormMain()
        {
            InitializeComponent();
            dbconnectionString = "Data Source=;Initial Catalog=objectsoke;User ID=;Password=";
            dbconnectionStringBlackouts = "Data Source=;Initial Catalog=RTP3ESH;User ID=rtp3esh_admin;Password=";
            dbconnectionStringIESBK = "Data Source=;Initial Catalog=iesbk;User ID=;Password=";            
            dbconnectionStringIESBK2 = "Data Source=;Initial Catalog=iesbk2;User ID=;Password=";
        }
        
        private void ribbonControl1_Merge(object sender, DevExpress.XtraBars.Ribbon.RibbonMergeEventArgs e)
        {
            ribbonControl1.SelectedPage = ribbonControl1.MergedRibbon.Pages[0];
        }

        // проекты
        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblProjectType");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // филиалы
        private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblFilial");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // типы точек учета
        private void barButtonItem5_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblMetPointType");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // типы опроса
        private void barButtonItem6_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblOprosType");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // места установки
        private void barButtonItem7_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblMetPointType");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // ручная загрузка данных
        private void barButtonItem8_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Загрузка данных";

            form1.Show();
        }
                
        // объекты ЭСХ
        private void barButtonItem9_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblObjects");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem12_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormMapGeoDraw form1 = new FormMapGeoDraw(dbconnectionString, permissionAdmin);
            form1.Text = "Электронные карты ЭСК";
            form1.MdiParent = this;            
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // типы объектоы
        private void barButtonItem13_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblObjectType");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // типы свойств объектов
        private void barButtonItem14_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblObjectPropertiesTypes");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // свойства объектов
        private void barButtonItem15_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblObjectPropertiesValues");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Панорама подрядчики
        private void barButtonItem17_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblContractor");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Панорама источники данных
        private void barButtonItem18_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblMapSource");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Панорама объекты
        private void barButtonItem19_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblPanoramaObject");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Панорама семантика
        private void barButtonItem20_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblPanoramaSemantic");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Панорама значения свойств объектов
        private void barButtonItem21_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblPanoramaObjSemValues");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Панорама координаты объектов
        private void barButtonItem22_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblPanoramaCoords");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // муниципальные образования
        private void barButtonItem23_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblMunicipality");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // филиалы
        private void barButtonItem24_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblFilial");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // Выгрузка N,X,Y,(X,Y)
        private void barButtonItem25_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.Text = "Выгрузка данных";
            
            form1.spreadsheetControl1.Hide();
            splashScreenManager1.ShowWaitForm();
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            DataSetPnrmMapSrc DataSetLoad = new DataSetPnrmMapSrc();
            DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();
            tblPanoramaCoordsTableAdapter.Fill(DataSetLoad.tblPanoramaCoords);
            DataRow[] coordrows = DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL");

            int row = 0;
            worksheet[row, 0].Value = "IDMAPSRC";
            worksheet[row, 1].Value = "PNRMPOINT";
            worksheet[row, 2].Value = "IDPNRMOBJECT";
            worksheet[row, 3].Value = "PNRMSUBJECT";
            worksheet[row, 4].Value = "X";
            worksheet[row, 5].Value = "Y";
            worksheet[row, 6].Value = "X,Y";
            worksheet[row, 7].Value = "ALT,LONGT";

            foreach (DataRow coordrow in coordrows)
            {
                row++;
                worksheet[row, 0].Value = coordrow["idmapsrc"].ToString();
                worksheet[row, 1].Value = coordrow["pnrmPOINT"].ToString();
                worksheet[row, 2].Value = coordrow["idpnrmOBJECT"].ToString();
                worksheet[row, 3].Value = coordrow["pnrmSUBJECT"].ToString();
                worksheet[row, 4].Value = coordrow["pnrmX"].ToString();
                worksheet[row, 5].Value = coordrow["pnrmY"].ToString();
                worksheet[row, 6].Value = String.Concat(worksheet[row, 4].Value.ToString(), ",", worksheet[row, 5].Value.ToString());
            }

            worksheet.Columns.AutoFit(0, 7);

            form1.spreadsheetControl1.Show();
            splashScreenManager1.CloseWaitForm();

            form1.Show();
        }

        // выгрузка X,Y столбцами
        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormCoordZonesForLoad form1 = null;
            form1 = new FormCoordZonesForLoad();

            form1.DataSetLoad = new DataSetPnrmMapSrc();
            form1.tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();
            form1.tblPanoramaCoordsTableAdapter.Fill(form1.DataSetLoad.tblPanoramaCoords);

            DataRow[] coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '1'");
            form1.textBox1.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '2'");
            form1.textBox2.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '3'");
            form1.textBox3.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '4'");
            form1.textBox4.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '5'");
            form1.textBox5.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '6'");            
            form1.textBox6.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '7'");
            form1.textBox7.Text = coordrows.Count().ToString();

            coordrows = form1.DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '8'");
            form1.textBox8.Text = coordrows.Count().ToString();

            form1.Show();
        }  // // выгрузка X,Y столбцами  

        private void FormMain_Load(object sender, EventArgs e)
        {            
            FormLogin form1 = null;
            form1 = new FormLogin();
            if (form1.ShowDialog(this) == DialogResult.OK)
            {
                barButtonItem37.Visibility = BarItemVisibility.Never; // Привязка потребителей РТП-3
                if (form1.textEditLogin.Text.Equals("admin") && form1.textEditPassword.Text.Equals(""))
                {
                    permissionAdmin = 100;

                    this.ribbonPage2.Visible = true;
                    this.ribbonPage3.Visible = true;
                    ribbonPageGroup6.Visible = true;

                    barButtonItem37.Visibility = BarItemVisibility.Always; // Привязка потребителей РТП-3

                    /*textBox1.Visible = true;
                    textBox2.Visible = true;
                    textBox3.Visible = true;*/
                }
                else if (form1.textEditLogin.Text.Equals("ods") && form1.textEditPassword.Text.Equals(""))
                {
                    permissionAdmin = 2;
                }
                else
                {
                    permissionAdmin = 1;
                }

                this.ribbonPageGroupRTP3_Admin.Visible = permissionAdmin == 100;
            }
        }

        private void barButtonItem26_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKPeriod");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem27_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKls");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem28_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKlsprop");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem29_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKlspropvalue");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem31_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKtemplate");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem32_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKlstype");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        private void barButtonItem33_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            FormGridDataView form1 = null;
            form1 = new FormGridDataView("tblIESBKotdelenie");
            form1.MdiParent = this;
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // ИЭСБК отчеты - "Отчет о работе контролеров"
        private void barButtonItem30_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Отчет о работе контролеров";
            IWorkbook workbook = form1.spreadsheetControl1.Document;            

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            {
                string otdelenieid = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();
                                
                Worksheet worksheet = workbook.Worksheets[i];
                worksheet.Name = captionotd;
                workbook.Worksheets.Add();

                // задаем смещение табличной части отчета
                int strow = 4;
                int stcol = 0;

                string peryear = "2018";
                string permonth = "11";

                worksheet[0, 0].SetValue("ФЛ без ОДПУ");
                worksheet[1, 0].SetValue(captionotd);
                worksheet[2, 0].SetValue(peryear+", "+permonth);

                worksheet[strow + 0, stcol + 0].SetValue("Вид последнего показания");
                worksheet[strow + 0, stcol + 1].SetValue("Всего л/с");
                worksheet[strow + 0, stcol + 2].SetValue("Всего ПО");
                worksheet[strow + 0, stcol + 3].SetValue("Средний ПО");
                worksheet[strow + 0, stcol + 4].SetValue("Доля л/с, %");
                worksheet[strow + 0, stcol + 5].SetValue("Доля ПО, %");

                string queryString =
                    "SELECT otdelenie.captionotd,pvLPT.propvalue AS typeinfo,COUNT(*) AS totalls, SUM(CAST(REPLACE(pvPOJ.propvalue, ',', '.') AS float)) AS totalpo"+
                    " FROM"+
                    " ([iesbk].[dbo].[tblIESBKlspropvalue] pvLPT"+
                    " LEFT JOIN[iesbk].[dbo].[tblIESBKotdelenie] otdelenie"+
                    " ON otdelenie.otdelenieid = pvLPT.otdelenieid)"+
	                " LEFT JOIN"+
                    " (SELECT pvPO.propvalue, pvPO.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvPO"+
                    " WHERE pvPO.lstypeid = 'ФЛ' AND pvPO.periodyear = '"+peryear+ "' AND pvPO.periodmonth = '" + permonth + "' AND pvPO.lspropertieid = '27' AND pvPO.otdelenieid = '" + otdelenieid+"') pvPOJ"+
                    " ON pvLPT.codeIESBK = pvPOJ.codeIESBK"+
                    " WHERE pvLPT.lstypeid = 'ФЛ' AND pvLPT.periodyear = '" + peryear + "' AND pvLPT.periodmonth = '" + permonth + "' AND pvLPT.lspropertieid = '26' AND pvLPT.otdelenieid = '" + otdelenieid+"'"+

                    " GROUP BY otdelenie.captionotd,pvLPT.propvalue"+
                    " ORDER BY pvLPT.propvalue";
                      
                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                // делаем общий подсчет---------------------------------------
                string queryString2 =
                "SELECT" +
                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '" + peryear + "' AND pv201602.periodmonth = '" + permonth + "' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = '"+otdelenieid+"') AS totalls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '" + peryear + "' AND pv201602.periodmonth = '" + permonth + "' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = '"+otdelenieid +"' AND pv201602.propvalue IS NOT NULL) AS totalpo";

                DataTable tableTOTALlsPO = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALlsPO, dbconnectionStringIESBK, queryString2);

                int totalls = Convert.ToInt32(tableTOTALlsPO.Rows[0]["totalls"]);
                double totalpo = String.IsNullOrWhiteSpace(tableTOTALlsPO.Rows[0]["totalpo"].ToString()) ? 0 :Convert.ToDouble(tableTOTALlsPO.Rows[0]["totalpo"]);
                //------------------------------------------------------------

                for (int j = 0; j < tableTOTALls.Rows.Count; j++)
                {
                    if (String.IsNullOrWhiteSpace(tableTOTALls.Rows[j]["typeinfo"].ToString())) worksheet[strow + j + 1, stcol + 0].SetValue("Расчетное(норматив)");
                    else worksheet[strow + j + 1, stcol + 0].SetValue(tableTOTALls.Rows[j]["typeinfo"].ToString());
                    worksheet[strow + j + 1, stcol + 1].SetValue(tableTOTALls.Rows[j]["totalls"]);
                    worksheet[strow + j + 1, stcol + 2].SetValue(tableTOTALls.Rows[j]["totalpo"]);

                    double totalPO = 0;
                    if (String.IsNullOrWhiteSpace(tableTOTALls.Rows[j]["totalpo"].ToString())) totalPO = 0;
                    else totalPO = Convert.ToDouble(tableTOTALls.Rows[j]["totalpo"]);
                    double vesPO = totalPO / Convert.ToInt32(tableTOTALls.Rows[j]["totalls"]);
                    worksheet[strow + j + 1, stcol + 3].SetValue(vesPO);

                    double dolyaLS = Convert.ToDouble(tableTOTALls.Rows[j]["totalls"]) / totalls*100;
                    worksheet[strow + j + 1, stcol + 4].SetValue(dolyaLS);
                    double dolyaPO = totalPO / totalpo*100;
                    worksheet[strow + j + 1, stcol + 5].SetValue(dolyaPO);

                    worksheet[strow + j + 1, stcol + 1].NumberFormat = "#####";
                    worksheet[strow + j + 1, stcol + 2].NumberFormat = "#";
                    worksheet[strow + j + 1, stcol + 3].NumberFormat = "#.##";
                    worksheet[strow + j + 1, stcol + 4].NumberFormat = "#";
                    worksheet[strow + j + 1, stcol + 5].NumberFormat = "#";
                }

                worksheet.Columns.AutoFit(0, 10);
            } // for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)

            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // ИЭСБК отчеты - "Отчет о работе контролеров"

        // отчет-"шахматка" по наличию л/с и полезного отпуска
        private void barButtonItem34_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
                        
            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Шахматка с января 2018 по ноябрь 2018";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            workbook.History.IsEnabled = false;            
            form1.spreadsheetControl1.BeginUpdate();

            // константы
            //int MAX_PERIOD_MONTH = 12;
            int MAX_PERIOD_MONTH = 11; // январь - ноябрь
            //int MAX_PERIOD_MONTH = 12; // январь - декабрь
            string periodmonthnext = "12"; // переделать

            int columns_in_period_auto = 20+1; // колонок в периоде для автоматического вывода
            int columns_in_period_manual = 4; // колонок в периоде для ручного вывода
            int columns_in_period = columns_in_period_auto + columns_in_period_manual; // общее кол-во колонок в периоде
            int FIRST_COLUMNS = 8;
            int END_COLUMNS = 1 + 3 + 8 + MAX_PERIOD_MONTH + 7;

            //int MAXPROPMAS = 9;
            int MAXCOLinWRKSH = FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + END_COLUMNS;

            /*for (int col = 0; col < MAXCOLinWRKSH; col++)
            {
                worksheet.Columns[col].Font.Name = "Arial";
                worksheet.Columns[col].Font.Size = 8;
            }*/

            DateTime dt_IESBK_MIN = Convert.ToDateTime("01.01.2015"); // левая граница имеющихся данных в OLAP-кубе

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            // загружаем л/с
            DataSetIESBKTableAdapters.tblIESBKlsTableAdapter tblIESBKlsTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlsTableAdapter();
            tblIESBKlsTableAdapter.Fill(DataSetIESBKLoad.tblIESBKls);


            // продумать выборку!!!
            // св-во 36 - "Расход ОДН по нормативу" (убрал)
            /*string queryString = "SELECT DISTINCT codeIESBK,otdelenieid " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +                                
                                "WHERE periodyear = '2016' AND (periodmonth = '01' OR periodmonth = '07') AND lspropertieid='36' AND propvalue IS NULL";*/
            string queryString = "SELECT DISTINCT codeIESBK,otdelenieid " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                 "WHERE periodyear in ('2018') ";
                                 //"AND codeIESBK = 'КНОО00001590'";
                                 //" AND otdelenieid ='" + this.textBox1.Text + "'";
                                
                                //"WHERE periodyear = '2017'";
                                //" AND otdelenieid = 'ЦО'";
                                //+ " AND otdelenieid = '" + this.textBox1.Text + "'";// AND lspropertieid='36' AND propvalue IS NULL";
                                //+ " AND (otdelenieid = 'СОС' OR otdelenieid = 'СОЗ')";// AND lspropertieid='36' AND propvalue IS NULL";
            DataTable tableTOTALls10 = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls10, dbconnectionStringIESBK, queryString);
            //-----------------------------

            // выводим заголовки столбцов -----------------------------------------------------

            // статические
            worksheet[0, 0].SetValue("№ п/п");
            worksheet[0, 1].SetValue("Отделение ИЭСБК");
            worksheet[0, 2].SetValue("Код л/с ИЭСБК");            
            worksheet[0, 3].SetValue("ФИО");                
            worksheet[0, 4].SetValue("Населенный пункт");
            worksheet[0, 5].SetValue("Улица");
            worksheet[0, 6].SetValue("Дом");
            worksheet[0, 7].SetValue("Номер квартиры");
            //worksheet[0, 8].SetValue("Состояние ЛС (на 2016 08)");

            // периодические
            // id "периодических" свойств
            //int[] propidmas = new int[] { 6, 7, 50, 24, 25, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };
            int[] propidmas = new int[] { 51, 20, 24, 25, 6, 7, 50, 26, 27, 28, 55, 29, 30, 53, 31, 54, 32, 33, 34, 35, 36 };
            int idprop_PO_in_propidmas = 7+1; // индекс идентификатора поля ПО от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_lastPOK_in_propidmas = 2+1; // индекс идентификатора поля ПослПоказаниеПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_nomerPU_in_propidmas = 4+1; // индекс идентификатора поля ЗаводскойНомерПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int[] propidmas_doublevalue = new int[] { 27, 28, 55, 29, 30, 53, 31, 54, 32, 33, 34, 35, 36 }; // id числовых полей

            string queryStringlsprop = "SELECT lspropertieid, templateid, numcolumninfile, captionlsprop " +
                                        "FROM [iesbk].[dbo].[tblIESBKlsprop]";
            DataTable tableLSprop = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

            Color Color_IS_PO_NORMATIV_NOT_PU = Color.Orange; // цвет "норматив - безприборник"
            Color Color_IS_PO_NORMATIV_YES_PU = Color.Red; // цвет "норматив - приборник"
            Color Color_IS_PO_SREDNEMES_YES_PU = Color.Blue; // цвет "среднемесячное - приборник"
            Color Color_IS_PO_RASHOD_YES_PU = Color.Green; // цвет "расход по прибору"

            for (int period_i = 1; period_i < MAX_PERIOD_MONTH + 1; period_i++)
            {
                string periodyear = "";
                string periodmonth = "";

                periodyear = "2018";
                periodmonth = (period_i < 10) ? "0"+period_i.ToString() : period_i.ToString();

                /*// сделал для "перехода" года
                if (period_i == 1)
                { 
                    periodyear = "2017";
                    periodmonth = "12";
                }
                else
                if (period_i == 2)
                {
                    periodyear = "2018";
                    periodmonth = "01";
                }*/

                for (int k = 0; k < columns_in_period_auto; k++)
                {
                    DataRow[] lsproprows = tableLSprop.Select("lspropertieid = '" + propidmas[k].ToString() + "'");
                    string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["captionlsprop"].ToString() : null;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(periodyear + " " + periodmonth + ", " + propvaluestr);

                    Color cellColor = Color.Black;
                    // "красим" названия столбцов
                    if (propidmas[k] == 28) cellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (propidmas[k] == 29) cellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (propidmas[k] == 30 || propidmas[k] == 53) cellColor = Color_IS_PO_SREDNEMES_YES_PU;
                    else if (propidmas[k] == 31 || propidmas[k] == 54) cellColor = Color_IS_PO_NORMATIV_YES_PU;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].Font.Color = cellColor;
                }

                tableLSprop.Dispose();

                //------------------------------------------------------------------------------

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(periodyear + " " + periodmonth + ", " + "Расчетный полезный отпуск");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(periodyear + " " + periodmonth + ", " + "Расход по показаниям ПУ");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(periodyear + " " + periodmonth + ", " + "Начисленный полезный отпуск от ИЭСБК");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(periodyear + " " + periodmonth + ", " + "Отклонение (недополученный ПО)");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;
            }

            //------------------------------------------------------------------------------
            //string periodmonthnext = (MAX_PERIOD_MONTH + 1 < 10) ? "0" + (MAX_PERIOD_MONTH + 1).ToString() : (MAX_PERIOD_MONTH + 1).ToString();            
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue("2018" + " " + periodmonthnext + ", " + "среднемесячное (прогноз ПО)");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue("Предыдущее показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue("Предыдущее показание, показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue("Текущее показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue("Текущее показание, показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue("Разница показаний");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue("Разница дней");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue("Среднесуточный расход");

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue("ИТОГО Расход по показаниям ПУ с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].Font.Color = Color.Green;

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue("Начальное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue("Начальное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11].SetValue("Начальное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 12].SetValue("Начальное показание, номер ПУ");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 13].SetValue("Конечное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 14].SetValue("Конечное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 15].SetValue("Конечное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 16].SetValue("Конечное показание, номер ПУ");

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].SetValue("ИТОГО Полезный отпуск от ИЭСБК с начала года (проверка Расхода по показаниям)");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].Font.Color = Color.Blue;

            for (int period_i = 1; period_i <= MAX_PERIOD_MONTH; period_i++)
            {             
                worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17 + period_i].
                    SetValue(worksheet[0, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString());
            }

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].SetValue("ИТОГО Отклонение (недополученный ПО) с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].Font.Color = Color.Red;
            //----------------------------------------------------------

            // главный цикл
            for (int i = 0; i < tableTOTALls10.Rows.Count; i++)
            //for (int i = 0; i < 150; i++)
            {                
                string codeIESBK = tableTOTALls10.Rows[i]["codeIESBK"].ToString();
                //string codeIESBK = "ККОО00019257";

                splashScreenManager1.SetWaitFormDescription(String.Concat(codeIESBK, " ", (i + 1).ToString(), " из ", tableTOTALls10.Rows.Count.ToString(), ")"));

                string otdelenieid = tableTOTALls10.Rows[i]["otdelenieid"].ToString();
                //string otdelenieid = "АО";
                string otdeleniecapt = DataSetIESBKLoad.tblIESBKotdelenie.FindByotdelenieid(otdelenieid)["captionotd"].ToString();
                
                worksheet[i + 1, 0].SetValue((i + 1).ToString());
                worksheet[i + 1, 1].SetValue(otdeleniecapt);
                worksheet[i + 1, 2].SetValue(codeIESBK);
                
                // "красим" строку в зависимости от признака isvalid
                DataRow findrow = DataSetIESBKLoad.tblIESBKls.FindBycodeIESBKlstypeidotdelenieid(codeIESBK, "ФЛ", otdelenieid);
                if (findrow != null)
                {
                    if (findrow["isvalid"].ToString() == "0")
                        worksheet.Rows[i + 1].FillColor = Color.Yellow;                        
                }

                /*// состояние ЛС за июнь 2016  ---------------------------------------
                string periodyear = "2016";
                string periodmonth = "06";                    
                worksheet[i + 1, 3].SetValue(MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 51, SQLconnection));*/

                // основные параметры ЛС за начальный период (01 2017)  ---------------------------------------
                string periodyear = "2018";
                string periodmonth = "01";

                DataRow[] lsproprows;
                /*// "статические" свойства лицевого счета                
                queryStringlsprop = "SELECT codeIESBK, lspropertieid, propvalue " +
                                     "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                     "WHERE codeIESBK='" + codeIESBK + "' AND periodyear = '" + periodyear + "' AND periodmonth = '" + periodmonth + "'";
                tableLSprop = new DataTable();
                MyFUNC_SelectDataFromSQLwoutConnection(tableLSprop, SQLconnection, queryStringlsprop);*/

                // сделано для увеличения производительности
                // фио
                /*string lspropvalue = "";
                                
                lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 5, SQLconnection);
                worksheet[i + 1, 3].SetValue(lspropvalue);

                /*DataRow[] lsproprows = tableLSprop.Select("lspropertieid='5'");
                if (lsproprows.Length > 0) worksheet[i + 1, 3].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // населенный пункт                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 12, SQLconnection);
                worksheet[i + 1, 4].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='12'");
                if (lsproprows.Length > 0) worksheet[i + 1, 4].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // улица                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 13, SQLconnection);
                worksheet[i + 1, 5].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='13'");
                if (lsproprows.Length > 0) worksheet[i + 1, 5].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // дом                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 14, SQLconnection);
                worksheet[i + 1, 6].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='14'");
                if (lsproprows.Length > 0) worksheet[i + 1, 6].SetValue(lsproprows[0]["propvalue"].ToString());                */

                // номер квартиры                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 15, SQLconnection);
                worksheet[i + 1, 7].SetValue(lspropvalue);                
                /*lsproprows = tableLSprop.Select("lspropertieid='15'");
                if (lsproprows.Length > 0) worksheet[i + 1, 7].SetValue(lsproprows[0]["propvalue"].ToString());                */

                //tableLSprop.Dispose();
                //---------------------------------------

                // переменные для подсчета ПО по разнице показаний + сам ПО
                double POKstart = -1;
                double POKend = -1;
                
                DateTime POKstart_date = Convert.ToDateTime("01.01.1900");
                DateTime POKend_date = Convert.ToDateTime("01.01.1900"); 
                string POKstart_kind = null;
                string POKend_kind = null;
                int periodPOKstart = -1; // нумерация с 1 (январь)
                int periodPOKend = -1; // нумерация с 1 (январь)

                string nomerPUstart = null;
                string nomerPUend = null;
                //------------------------------------------------

                // "бежим" по периодическим свойствам лицевого счета
                int START_PERIOD_MONTH = 1;
                for (int period_i = START_PERIOD_MONTH; period_i < START_PERIOD_MONTH + MAX_PERIOD_MONTH; period_i++)
                {
                    periodyear = "2018";
                    periodmonth = (period_i < 10) ? "0"+period_i.ToString() : period_i.ToString();

                    /*// сделал для "перехода" года
                    periodyear = "";
                    periodmonth = "";

                    if (period_i == 1)
                    {
                        periodyear = "2017";
                        periodmonth = "12";
                    }
                    else
                    if (period_i == 2)
                    {
                        periodyear = "2018";
                        periodmonth = "01";
                    }*/

                    //-----------------------------------
                    // "статические" свойства лицевого счета                                

                    string lspropvalue = "";

                    // фио
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 5, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 3].SetValue(lspropvalue);
                    
                    // населенный пункт                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 12, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 4].SetValue(lspropvalue);
                    
                    // улица                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 13, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 5].SetValue(lspropvalue);
                    
                    // дом                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 14, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 6].SetValue(lspropvalue);
                    
                    // номер квартиры                
                    lspropvalue = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 15, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 7].SetValue(lspropvalue);
                    //-----------------------------------

                    queryStringlsprop = 
                        String.Concat("SELECT codeIESBK, lspropertieid, propvalue ",
                                     "FROM [iesbk].[dbo].[tblIESBKlspropvalue] ",
                                     "WHERE codeIESBK='", codeIESBK, "' AND periodyear = '", periodyear, "' AND periodmonth = '", periodmonth, "'");
                    tableLSprop = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

                    /*// состояние ЛС (по последнему расчетному периоду)
                    if (period_i == MAX_PERIOD_MONTH)
                    {                           
                        lsproprows = tableLSprop.Select("lspropertieid='51'");                        
                        if (lsproprows.Length > 0) worksheet[i + 1, 8].SetValue(lsproprows[0]["propvalue"].ToString());
                    }*/

                    // флаги для раскраски ячейки полезного отпуска                    
                    bool IS_PO_NORMATIV_NOT_PU = false; // флаг "норматив - безприборник"
                    bool IS_PO_NORMATIV_YES_PU = false; // флаг "норматив - приборник"
                    bool IS_PO_SREDNEMES_YES_PU = false; // флаг "среднемесячное - приборник"
                    bool IS_PO_RASHOD_YES_PU = false; // флаг "расход по прибору"

                    // выводим "периодические" поля - сделал в цикле                    
                    for (int k = 0; k < columns_in_period_auto; k++)
                    {
                        /*string strTEST = "12345678";
                        string propvaluestr = strTEST;*/

                        lsproprows = tableLSprop.Select(String.Concat("lspropertieid = '", propidmas[k].ToString(), "'"));

                        string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;                        
                        double? propvalue = null;

                        if (System.Array.IndexOf(propidmas_doublevalue, propidmas[k]) >= 0)
                        {
                            if (!String.IsNullOrWhiteSpace(propvaluestr)) propvalue = Convert.ToDouble(propvaluestr);
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvalue);
                        }
                        else
                        {
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvaluestr);
                        }

                        // формируем флаги для раскраски ячейки полезного отпуска
                        // помним о массиве периодических свойств
                        //int[] propidmas = new int[] { 6, 7, 50, 24, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };                        
                        if (propvalue != null && propvalue != 0)
                        {
                            if (propidmas[k] == 28) IS_PO_NORMATIV_NOT_PU = true;
                            else if (propidmas[k] == 29) IS_PO_RASHOD_YES_PU = true;
                            else if (propidmas[k] == 30) IS_PO_SREDNEMES_YES_PU = true;
                            else if (propidmas[k] == 31) IS_PO_NORMATIV_YES_PU = true;
                        };
                        //-------------------------------------------------------
                    } // выводим "периодические" поля - сделал в цикле

                    // раскрашиваем колонку ПолОтп в зависимости от "слагаемых"
                    // помним о массиве периодических свойств
                    //int[] propidmas = new int[] { 6, 7, 50, 24, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };

                    Color PolOtpCellColor = Color.Black;
                    if (IS_PO_RASHOD_YES_PU) PolOtpCellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (IS_PO_NORMATIV_NOT_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (IS_PO_NORMATIV_YES_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_YES_PU;
                    else if (IS_PO_SREDNEMES_YES_PU) PolOtpCellColor = Color_IS_PO_SREDNEMES_YES_PU;                    
                    worksheet[i + 1, FIRST_COLUMNS + (period_i - 1 ) * columns_in_period + idprop_PO_in_propidmas].Font.Color = PolOtpCellColor; // "красим" ПолезныйОтпуск - propid = 27

                    //-----------------------------------------------------------------
                    // формируем "периодические" колонки "Расход по показаниям" и "Отклонение (недополученный ПО)", если не было замены ПУ

                    if (period_i > START_PERIOD_MONTH) // пропускаем первый месяц, т.к. в нем не найдем "предыдущих показаний"
                    {
                        string propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        double POKend_period = -1;
                        //double? POIESBK_period = null;
                        double POIESBK_period = 0;
                        double POIESBKend_period = 0; // ПО последнего периода

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";"))
                        {
                            POKend_period = Convert.ToDouble(propvaluestr);

                            double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                            POIESBK_period += POcellvalue;

                            POIESBKend_period = POcellvalue;
                        }

                        string nomerPUend_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        /*lsproprows = tableLSprop.Select("lspropertieid='22'"); // предыдущее показание ПУ
                        propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                        int period_i_step = 1;
                        double POKstart_period = -1;
                        string nomerPUstart_period = null;
                        propvaluestr = null;
                                                
                        do
                        {                            
                            propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();                         
                            nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();

                            // суммируем полезный отпуск, пропуская начальный интервал
                            if (String.IsNullOrWhiteSpace(propvaluestr))
                            {
                                double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.NumericValue;
                                POIESBK_period += POcellvalue;
                            }

                            period_i_step += 1;

                        } while (String.IsNullOrWhiteSpace(propvaluestr) && period_i_step < period_i);

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POKstart_period = Convert.ToDouble(propvaluestr);

                        //nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 2) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        // если имеются оба показания и не было замены ПУ, то считаем значения
                        if (POKend_period != -1 && POKstart_period != -1 && String.Equals(nomerPUstart_period, nomerPUend_period))
                        //if (POKend_period != -1 && POKstart_period != -1 && POKstart_period <= POKend_period)                         
                        {
                            /*propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                            double? POIESBK_period = null;
                            if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POIESBK_period = Convert.ToDouble(propvaluestr);*/

                            double POIESBKPU_period = POKend_period - POKstart_period;
                            double POIESBKDelta_period = POIESBKPU_period - POIESBK_period;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(POIESBKend_period + POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(POIESBKPU_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(POIESBK_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;                            
                        }

                    } // if (period_i > START_PERIOD_MONTH) 
                    //-----------------------------------------------------------------

                    // ищем стартовое и конечное показание для расчета ИТОГОВОГО ПО по показаниям (за все периоды)                  
                    lsproprows = tableLSprop.Select("lspropertieid='25'");
                    string pokstr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                    
                    if (!String.IsNullOrWhiteSpace(pokstr) && !pokstr.Contains(";"))
                    {
                        if (POKstart == -1)
                        {
                            POKstart = Convert.ToDouble(pokstr);
                            periodPOKstart = period_i;

                            lsproprows = tableLSprop.Select("lspropertieid='7'");
                            nomerPUstart = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;

                            lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKstart_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());

                            lsproprows = tableLSprop.Select("lspropertieid='26'");
                            POKstart_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                        }
                        else
                        {
                            // поменять местами нижнее условие и присовение значение конечному показанию!!!!!
                            POKend = Convert.ToDouble(pokstr);
                            periodPOKend = period_i;

                            lsproprows = tableLSprop.Select("lspropertieid='7'");
                            nomerPUend = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;

                            lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKend_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());

                            lsproprows = tableLSprop.Select("lspropertieid='26'");
                            POKend_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;

                            if (!String.Equals(nomerPUstart, nomerPUend)) // если номера ПУ не равны, то последнее делаем начальным
                            {
                                nomerPUstart = nomerPUend;
                                POKstart = POKend;
                                POKstart_date = POKend_date;
                                POKstart_kind = POKend_kind;

                                POKend = -1;

                                periodPOKstart = periodPOKend;
                                periodPOKend = -1;                        
                            };
                        }                        
                    }
                    
                    //-----------------------------------------------------------------

                    tableLSprop.Dispose();

                } // for (int period_i = 1; period_i < 7; period_i++)

                if (POKstart != -1 && POKend != -1) // если имеются оба показания ПУ (для расчета ПО по показаниям ПУ)
                {
                    /*// суммируем полезный отпуск от ИЭСБК со следующего расчетного периода, которому предшествовало показание ПУ
                    lsproprows = tableLSprop.Select("lspropertieid='27'");
                    string postr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                    if (POKstart != -1 && POKend == -1 && !String.IsNullOrWhiteSpace(postr)) POIESBKTotal += Convert.ToDouble(postr);*/

                    // суммируем полезный отпуск от ИЭСБК по ранее заполненным колонкам со следующего расчетного периода, которому предшествовало показание ПУ
                    double POIESBKTotal = 0;
                    for (int period_i = periodPOKstart+1; period_i <= periodPOKend; period_i++)
                    {
                        double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                        POIESBKTotal += POcellvalue;

                        // выводим расшифровку формирования ПО
                        worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17 + period_i].SetValue(POcellvalue);
                    }

                    double POIESBKPU = POKend - POKstart;
                    double POIESBKDelta = POIESBKPU - POIESBKTotal;

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue(POIESBKPU);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].Font.Color = Color.Green;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue(POKstart_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue(POKstart);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11].SetValue(POKstart_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 12].SetValue(nomerPUstart);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 13].SetValue(POKend_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 14].SetValue(POKend);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 15].SetValue(POKend_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 16].SetValue(nomerPUend);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].SetValue(POIESBKTotal);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 17].Font.Color = Color.Blue;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].SetValue(POIESBKDelta);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 18 + MAX_PERIOD_MONTH].Font.Color = Color.Red;

                    // beg расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)

                    string periodyeartek = "2018"; // было 2017
                    //string periodmonthtek = (MAX_PERIOD_MONTH < 10) ? "0" + MAX_PERIOD_MONTH.ToString() : MAX_PERIOD_MONTH.ToString();

                    // сделал для перехода года, потом убрать!!!
                    string periodmonthtek = (MAX_PERIOD_MONTH-1 < 10) ? "0" + (MAX_PERIOD_MONTH-1).ToString() : (MAX_PERIOD_MONTH-1).ToString();
                    string codels = codeIESBK;

                    // ищем ближайшее "правое" показание                
                    string value_right = null;
                    string dtvalue_right = null;

                    DateTime dt_right = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);
                    string year_right = null;
                    string month_right = null;

                    dt_right = dt_right.AddMonths(+1); // учитываем текущий месяц, т.е. +1-1 = 0

                    while (String.IsNullOrWhiteSpace(value_right) && dt_right >= dt_IESBK_MIN)
                    {
                        dt_right = dt_right.AddMonths(-1);
                        year_right = dt_right.Year.ToString();
                        month_right = null;
                        if (dt_right.Month < 10) month_right = "0" + dt_right.Month.ToString();
                        else month_right = dt_right.Month.ToString();

                        value_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 25, SQLconnection); // свойство "Текущее показание ПУ"                                        
                    };
                    //----------------------------------

                    // ищем "левое" показание, при условии, что нашли "правое" ------------------
                    string value_left = null;
                    string dtvalue_left = null;
                    string year_left = null;
                    string month_left = null;

                    if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    {
                        // получаем дату "правого" показания
                        dtvalue_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                        /*// выводим информацию о "правом" показании
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(dtvalue_right); // дата

                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(value_right); // показание

                        string rightpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_right, month_right, 26, SQLconnection);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(rightpok_type); // вид*/

                        DateTime dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-5); // было -6, отматываем -5-1 = -6 мес. = 180 дней от "правого" показания
                        
                        while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN)
                        {
                            dt_left = dt_left.AddMonths(-1);
                            year_left = dt_left.Year.ToString();
                            month_left = null;
                            if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                            else month_left = dt_left.Month.ToString();

                            value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                        };

                        // если нет данных да период не менее 6 мес., то ищем за в периоде [6 мес.;3 мес.]                        
                        if (String.IsNullOrWhiteSpace(value_left))
                        {
                            dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-7); // отматываем 7 мес., т.к. в теле цикла сразу +1, т.е. -7+1 = -6
                            
                            DateTime dt_IESBK_left_MAX = Convert.ToDateTime(dtvalue_right).AddMonths(-3);

                            while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN && dt_left < dt_IESBK_left_MAX)
                            {
                                dt_left = dt_left.AddMonths(+1);
                                year_left = dt_left.Year.ToString();
                                month_left = null;
                                if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                                else month_left = dt_left.Month.ToString();

                                value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                            };
                        } // if (String.IsNullOrWhiteSpace(value_left)) // если нет данных за период не менее 6 мес.

                        // получаем дату "левого" показания и выводим информацию о нем
                        if (!String.IsNullOrWhiteSpace(value_left))
                        {
                            dtvalue_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                            /*// выводим информацию о "правом" показании
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(dtvalue_left); // дата

                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(value_left); // показание

                            string leftpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_left, month_left, 26, SQLconnection);
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(leftpok_type); // вид*/
                        }

                        // если даты "левого" и "правого" показаний не пустые, то формируем расчет среднемесячного
                        if (!String.IsNullOrWhiteSpace(dtvalue_left) && !dtvalue_left.Contains(";") && !value_left.Contains(";") && !String.IsNullOrWhiteSpace(dtvalue_right))
                        {                            
                            double pokleft = Convert.ToDouble(value_left);
                            double pokright = Convert.ToDouble(value_right);

                            // если не нарушен нарастающий итог
                            if (pokleft <= pokright)
                            {
                                System.TimeSpan deltaday = Convert.ToDateTime(dtvalue_right) - Convert.ToDateTime(dtvalue_left);
                                double deltapok = pokright - pokleft;

                                double srednesut_calc = deltapok / deltaday.Days;
                                double srmes_calc = Math.Round(srednesut_calc * DateTime.DaysInMonth(Convert.ToInt32(periodyeartek), Convert.ToInt32(periodmonthtek)));

                                // формируем отчет -----------------------------------------------

                                /*if (!String.IsNullOrWhiteSpace(srmes_iesbk_str)) // еслм СрМес ПО ИЭСБК присутствует
                                {
                                    DateTime dt_period = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);

                                    if (srmes_iesbk >= 0 && Convert.ToDateTime(dtvalue_right).CompareTo(dt_period) < 0) // не выводим наши расчеты, если СрМес ИЭСБК < 0 и правое показание принадлежит текущему периоду анализа
                                    {
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(srmes_calc); // СрМес РАСЧ                                

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Red;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(srmes_calc - srmes_iesbk); // Недополученный ПО                                
                                    }
                                }*/

                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue(srmes_calc);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;

                                // выводим расшифровку расчета Прогноза (СрМес)
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue(dtvalue_left);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue(pokleft);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue(dtvalue_right);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue(pokright);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue(deltapok);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue(deltaday.Days);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue(srednesut_calc);

                            } // if (pokleft <= pokright)
                              //-----------------------------------------------------------------
                        } // if (!String.IsNullOrWhiteSpace(dtvalue_left) && !String.IsNullOrWhiteSpace(dtvalue_right))

                    } // if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    // end расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)
                }

                //splashScreenManager1.SetWaitFormDescription(String.Concat("Обработка данных (", (i + 1).ToString(), " из ", tableTOTALls10.Rows.Count.ToString(), ")"));
                //splashScreenManager1.SetWaitFormDescription(String.Concat(codeIESBK, " ", (i + 1).ToString(), " из ", tableTOTALls10.Rows.Count.ToString(), ")"));

            } // for (int i = 0; i < tableTOTALls10.Rows.Count; i++)

            for (int col = 0; col < MAXCOLinWRKSH; col++)
            {
                worksheet.Columns[col].Font.Name = "Arial";
                worksheet.Columns[col].Font.Size = 8;
            }

            // форматируем строку-заголовок
            worksheet.Rows[0].Font.Bold = true;
            worksheet.Rows[0].Alignment.WrapText = true;
            worksheet.Rows[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Rows[0].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            worksheet.Rows[0].AutoFit();

            worksheet.Columns.AutoFit(0, MAXCOLinWRKSH);

            worksheet.Columns.Group(3, 7, true); // группируем по колонкам "ФИО" - "Номер квартиры"            

            // группируем колонки расшифровки прогноза СрМес
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7,true);
            
            //int[] propidmas = new int[] { 6, 7, 50, 24, 25, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };
            //int[] propidmas = new int[] { 51, 24, 25, 6, 7, 50, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };

            // группируем "периодические" значения - в частности расшифровку ПО
            // КРИВО!!!! ПРИ ДОБАВЛЕНИИ СВОЙСТВА ВПЕРИОДИЧЕСКУЮ СЕКЦИЮ ЗДЕСЬ ПРИХОДИТСЯ ДОБАВЛЯТЬ!!!
            for (int period_i = 0; period_i < MAX_PERIOD_MONTH; period_i++)
            {
                //worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period, FIRST_COLUMNS + period_i * columns_in_period + 2, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 2 + 1, FIRST_COLUMNS + period_i * columns_in_period + 4 + 1 + 1, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 7 + 1 + 1, FIRST_COLUMNS + period_i * columns_in_period + 7 + 10 + 2 + 1, true); // после "ПолОтп"
            }

            // группируем последние колонки анализа ПО по показаниям ПУ
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1 + 1 + 7, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1 + 7, true);

            // группируем колонки слагаемых ПО от ИЭСБК
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 2 + 1 + 7, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1 + MAX_PERIOD_MONTH + 1 + 7, true);                                    

            worksheet.FreezeRows(0); // "фиксируем" верхнюю строку

            form1.spreadsheetControl1.EndUpdate();

            SQLconnection.Close();
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // отчет-"шахматка" по наличию л/с и полезного отпуска

        // ИЭСБК отчеты - "Кол-во л/с и ПО"
        private void barButtonItem35_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "ИЭСБК кол-во лс и полезный отпуск";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            string queryString =
                "SELECT otd.captionotd," +
                    " (SELECT COUNT(*)"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602"+
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '01' AND pv201602.lspropertieid = '3' AND"+
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls,"+
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602"+
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '01' AND pv201602.lspropertieid = '27' AND"+
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po,"+

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '02' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '02' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '03' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '03' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '04' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '04' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '05' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '05' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '06' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '06' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po," +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '07' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '07' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '08' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '08' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '09' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '09' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '10' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '10' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po, " +

                    " (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '11' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2018' AND pv201602.periodmonth = '11' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po " +

                    /*" (SELECT COUNT(*)" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2017' AND pv201602.periodmonth = '12' AND pv201602.lspropertieid = '3' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid) AS total201602ls," +
                    " (SELECT SUM(CAST(REPLACE(pv201602.propvalue, ',', '.') AS float))" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pv201602" +
                    " WHERE pv201602.lstypeid = 'ФЛ' AND pv201602.periodyear = '2017' AND pv201602.periodmonth = '12' AND pv201602.lspropertieid = '27' AND" +
                    " pv201602.otdelenieid = otd.otdelenieid AND pv201602.propvalue IS NOT NULL) AS total201602po" +*/

                " FROM[iesbk].[dbo].tblIESBKotdelenie otd" +
                " GROUP BY otd.otdelenieid,otd.captionotd" +
                " ORDER BY otd.otdelenieid";

            DataTable tableTOTALls10 = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls10, dbconnectionStringIESBK, queryString);
            //-----------------------------

            // задаем смещение табличной части отчета
            int strow = 2;
            int stcol = 0;

            worksheet[0, 0].SetValue("ФЛ без ОДПУ");

            worksheet[strow + 0, stcol + 0].SetValue("Отделение ИЭСБК");            
            worksheet[strow + 0, stcol + 1].SetValue("2018 01 лс");
            worksheet[strow + 0, stcol + 2].SetValue("2018 01 ПО");
            worksheet[strow + 0, stcol + 3].SetValue("2018 01 среднее");
            worksheet[strow + 0, stcol + 4].SetValue("2018 02 лс");
            worksheet[strow + 0, stcol + 5].SetValue("2018 02 ПО");
            worksheet[strow + 0, stcol + 6].SetValue("2018 02 среднее");
            worksheet[strow + 0, stcol + 7].SetValue("2018 03 лс");
            worksheet[strow + 0, stcol + 8].SetValue("2018 03 ПО");
            worksheet[strow + 0, stcol + 9].SetValue("2018 03 среднее");
            worksheet[strow + 0, stcol + 10].SetValue("2018 04 лс");
            worksheet[strow + 0, stcol + 11].SetValue("2018 04 ПО");
            worksheet[strow + 0, stcol + 12].SetValue("2018 04 среднее");
            worksheet[strow + 0, stcol + 13].SetValue("2018 05 лс");
            worksheet[strow + 0, stcol + 14].SetValue("2018 05 ПО");
            worksheet[strow + 0, stcol + 15].SetValue("2018 05 среднее");
            worksheet[strow + 0, stcol + 16].SetValue("2018 06 лс");
            worksheet[strow + 0, stcol + 17].SetValue("2018 06 ПО");
            worksheet[strow + 0, stcol + 18].SetValue("2018 06 среднее");
            worksheet[strow + 0, stcol + 19].SetValue("2018 07 лс");
            worksheet[strow + 0, stcol + 20].SetValue("2018 07 ПО");
            worksheet[strow + 0, stcol + 21].SetValue("2018 07 среднее");
            worksheet[strow + 0, stcol + 22].SetValue("2018 08 лс");
            worksheet[strow + 0, stcol + 23].SetValue("2018 08 ПО");
            worksheet[strow + 0, stcol + 24].SetValue("2018 08 среднее");
            worksheet[strow + 0, stcol + 25].SetValue("2018 09 лс");
            worksheet[strow + 0, stcol + 26].SetValue("2018 09 ПО");
            worksheet[strow + 0, stcol + 27].SetValue("2018 09 среднее");
            worksheet[strow + 0, stcol + 28].SetValue("2018 10 лс");
            worksheet[strow + 0, stcol + 29].SetValue("2018 10 ПО");
            worksheet[strow + 0, stcol + 30].SetValue("2018 10 среднее");
            worksheet[strow + 0, stcol + 31].SetValue("2018 11 лс");
            worksheet[strow + 0, stcol + 32].SetValue("2018 11 ПО");
            worksheet[strow + 0, stcol + 33].SetValue("2018 11 среднее");
            /*worksheet[strow + 0, stcol + 34].SetValue("2017 12 лс");
            worksheet[strow + 0, stcol + 35].SetValue("2017 12 ПО");
            worksheet[strow + 0, stcol + 36].SetValue("2017 12 среднее");*/

            for (int i = 0; i < tableTOTALls10.Rows.Count; i++)            
            {                   
                worksheet[strow + i + 1, stcol + 0].SetValue(tableTOTALls10.Rows[i][0].ToString());

                double? dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][2].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][1].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][2].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][1].ToString());
                    worksheet[strow + i + 1, stcol + 1].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][1].ToString()));
                    worksheet[strow + i + 1, stcol + 2].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][2].ToString()));
                    worksheet[strow + i + 1, stcol + 3].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 1].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 2].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 3].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][4].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][3].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][4].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][3].ToString());
                    worksheet[strow + i + 1, stcol + 4].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][3].ToString()));
                    worksheet[strow + i + 1, stcol + 5].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][4].ToString()));
                    worksheet[strow + i + 1, stcol + 6].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 4].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 5].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 6].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][6].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][5].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][6].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][5].ToString());
                    worksheet[strow + i + 1, stcol + 7].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][5].ToString()));
                    worksheet[strow + i + 1, stcol + 8].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][6].ToString()));
                    worksheet[strow + i + 1, stcol + 9].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 7].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 8].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 9].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][8].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][7].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][8].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][7].ToString());
                    worksheet[strow + i + 1, stcol + 10].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][7].ToString()));
                    worksheet[strow + i + 1, stcol + 11].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][8].ToString()));
                    worksheet[strow + i + 1, stcol + 12].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 10].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 11].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 12].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][10].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][9].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][10].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][9].ToString());
                    worksheet[strow + i + 1, stcol + 13].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][9].ToString()));
                    worksheet[strow + i + 1, stcol + 14].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][10].ToString()));
                    worksheet[strow + i + 1, stcol + 15].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 13].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 14].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 15].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][12].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][11].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][12].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][11].ToString());
                    worksheet[strow + i + 1, stcol + 16].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][11].ToString()));
                    worksheet[strow + i + 1, stcol + 17].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][12].ToString()));
                    worksheet[strow + i + 1, stcol + 18].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 16].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 17].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 18].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][14].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][13].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][14].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][13].ToString());
                    worksheet[strow + i + 1, stcol + 19].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][13].ToString()));
                    worksheet[strow + i + 1, stcol + 20].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][14].ToString()));
                    worksheet[strow + i + 1, stcol + 21].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 19].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 20].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 21].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][16].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][15].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][16].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][15].ToString());
                    worksheet[strow + i + 1, stcol + 22].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][15].ToString()));
                    worksheet[strow + i + 1, stcol + 23].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][16].ToString()));
                    worksheet[strow + i + 1, stcol + 24].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 22].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 23].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 24].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][18].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][17].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][18].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][17].ToString());
                    worksheet[strow + i + 1, stcol + 25].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][17].ToString()));
                    worksheet[strow + i + 1, stcol + 26].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][18].ToString()));
                    worksheet[strow + i + 1, stcol + 27].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 25].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 26].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 27].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][20].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][19].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][20].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][19].ToString());
                    worksheet[strow + i + 1, stcol + 28].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][19].ToString()));
                    worksheet[strow + i + 1, stcol + 29].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][20].ToString()));
                    worksheet[strow + i + 1, stcol + 30].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 28].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 29].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 30].NumberFormat = "#.##";

                //dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][22].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][21].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][22].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][21].ToString());
                    worksheet[strow + i + 1, stcol + 31].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][21].ToString()));
                    worksheet[strow + i + 1, stcol + 32].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][22].ToString()));
                    worksheet[strow + i + 1, stcol + 33].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 31].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 32].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 33].NumberFormat = "#.##";

                /*//dolyaPO = null;
                if (!String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][24].ToString()) && !String.IsNullOrWhiteSpace(tableTOTALls10.Rows[i][23].ToString()))
                {
                    dolyaPO = Convert.ToDouble(tableTOTALls10.Rows[i][24].ToString()) / Convert.ToInt32(tableTOTALls10.Rows[i][23].ToString());
                    worksheet[strow + i + 1, stcol + 34].SetValue(Convert.ToInt32(tableTOTALls10.Rows[i][23].ToString()));
                    worksheet[strow + i + 1, stcol + 35].SetValue(Convert.ToDouble(tableTOTALls10.Rows[i][24].ToString()));
                    worksheet[strow + i + 1, stcol + 36].SetValue(dolyaPO);
                }
                worksheet[strow + i + 1, stcol + 34].NumberFormat = "#####";
                worksheet[strow + i + 1, stcol + 35].NumberFormat = "#";
                worksheet[strow + i + 1, stcol + 36].NumberFormat = "#.##";*/

                splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + ")");
            } // for (int i = 0; i < tableTOTALls10.Rows.Count; i++)

            worksheet.Columns.AutoFit(0, 49); // переделать!!!
            worksheet.FreezeColumns(0);

            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // ИЭСБК отчеты - "Кол-во л/с и ПО"

        // ИЭСБК отчеты - "Нарушение нарастающего итога"
        private void barButtonItem36_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Нарушение нарастающего итога показаний";
            IWorkbook workbook = form1.spreadsheetControl1.Document;

            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            int MAXCOLinWRKSH = 17;
            Color infocolor_prev = Color.DimGray;

            //for (int i = 0; i < 1; i++)
            for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            {
                string otdelenieid = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();

                Worksheet worksheet = workbook.Worksheets[i];
                worksheet.Name = captionotd;
                workbook.Worksheets.Add();

                // задаем смещение табличной части отчета
                int strow = 6;
                int stcol = 0;

                string peryearprev = "2018";
                string permonthprev = "10";

                string peryear = "2018";
                string permonth = "11";

                for (int col = 0; col < MAXCOLinWRKSH; col++)
                {
                    worksheet.Columns[col].Font.Name = "Arial";
                    worksheet.Columns[col].Font.Size = 8;
                }

                worksheet[0, 0].SetValue("Нарушение нарастающего итога");
                worksheet[0, 0].Font.Bold = true;

                worksheet[1, 0].SetValue("3.3.2 ФЛ без ОДПУ");
                worksheet[2, 0].SetValue(captionotd);
                worksheet[3, 0].SetValue("Текущий период:" );
                worksheet[3, 1].SetValue(peryear + ", " + permonth);
                worksheet[4, 0].SetValue("Предыдущий период:");
                worksheet[4, 1].SetValue(peryearprev + ", " + permonthprev);

                worksheet[strow + 0, stcol + 0].SetValue("Код ИЭСБК");
                worksheet[strow + 0, stcol + 1].SetValue("ФИО");

                worksheet[strow + 0, stcol + 2].SetValue("Район");
                worksheet[strow + 0, stcol + 3].SetValue("Населенный пункт");
                worksheet[strow + 0, stcol + 4].SetValue("Улица");
                worksheet[strow + 0, stcol + 5].SetValue("Дом");
                worksheet[strow + 0, stcol + 6].SetValue("Номер квартиры");

                string periodstr = peryear + " " + permonth + ", ";
                string periodstrprev = peryearprev + " " + permonthprev + ", ";
                worksheet[strow + 0, stcol + 7].SetValue(periodstrprev + "предыдущее показание, дата");
                worksheet[strow + 0, stcol + 7].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 8].SetValue(periodstrprev + "предыдущее показание, значение");
                worksheet[strow + 0, stcol + 8].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 9].SetValue(periodstrprev + "предыдущее показание, номер ПУ");
                worksheet[strow + 0, stcol + 9].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 10].SetValue(periodstrprev + "предыдущее показание, источник");
                worksheet[strow + 0, stcol + 10].Font.Color = infocolor_prev;
                worksheet[strow + 0, stcol + 11].SetValue(periodstr + "начальное показание, дата");
                worksheet[strow + 0, stcol + 12].SetValue(periodstr + "начальное показание, значение");
                worksheet[strow + 0, stcol + 13].SetValue(periodstr + "начальное показание, номер ПУ");
                worksheet[strow + 0, stcol + 14].SetValue(periodstr + "начальное показание, источник");

                worksheet[strow + 0, stcol + 15].SetValue("НАРУШЕНИЕ значения при одинаковом номере ПУ");
                worksheet[strow + 0, stcol + 15].Font.Color = Color.Red;
                worksheet[strow + 0, stcol + 16].SetValue("НАРУШЕНИЕ даты при одинаковом номере ПУ");
                worksheet[strow + 0, stcol + 16].Font.Color = Color.Blue;
                //-----------------------

                string queryString =
                    "SELECT otdelenie.captionotd,"+
	                "pvStartPok2.codeIESBK AS codels,"+

                    "pvStartPokFIO.propvalue AS pvStartPokFIOvalue, " +
                    "pvStartPokAddr1.propvalue AS pvStartPokAddr1value, pvStartPokAddr2.propvalue AS pvStartPokAddr2value," +
                    "pvStartPokAddr3.propvalue AS pvStartPokAddr3value, pvStartPokAddr4.propvalue AS pvStartPokAddr4value," +
                    "pvStartPokAddr5.propvalue AS pvStartPokAddr5value, " +

                    "pvStartPokPUNomer.propvalue AS StartPokPUNomervalue, " +
                    "pvEndPokPUNomer.propvalue AS prevPokPUNomervalue, " +

                    "pvEndPok1.propvalue AS prevPOKdatePrevPer,pvEndPok2.propvalue AS prevPOKvaluePrevPer, pvEndPok3.propvalue AS prevPOKtypePrevPer," +
	                "pvStartPok1.propvalue AS startPOKdateTekPer,pvStartPok2.propvalue AS startPOKvalueTekPer,pvStartPok3.propvalue AS startPOKtypeTekPer"+
                    " FROM"+
                    " ([iesbk].[dbo].[tblIESBKlspropvalue] pvStartPok2"+
                    " LEFT JOIN[iesbk].[dbo].[tblIESBKotdelenie] otdelenie"+
                    " ON otdelenie.otdelenieid = pvStartPok2.otdelenieid)"+
	                " LEFT JOIN"+
                    " (SELECT pvStartPok1.propvalue, pvStartPok1.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPok1"+
                    " WHERE pvStartPok1.lstypeid = 'ФЛ' AND pvStartPok1.periodyear = '"+peryear+ "' AND pvStartPok1.periodmonth = '" + permonth + "' AND pvStartPok1.lspropertieid = '21' AND pvStartPok1.otdelenieid = '" + otdelenieid + "') pvStartPok1" +
                    " ON pvStartPok2.codeIESBK = pvStartPok1.codeIESBK"+
                    " LEFT JOIN"+
                    " (SELECT pvStartPok3.propvalue, pvStartPok3.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPok3"+
                    " WHERE pvStartPok3.lstypeid = 'ФЛ' AND pvStartPok3.periodyear = '" + peryear + "' AND pvStartPok3.periodmonth = '" + permonth + "' AND pvStartPok3.lspropertieid = '23' AND pvStartPok3.otdelenieid = '" + otdelenieid + "') pvStartPok3" +
                    " ON pvStartPok2.codeIESBK = pvStartPok3.codeIESBK"+

                    // добавляем поля ФИО и адреса
                    " LEFT JOIN" +
                    " (SELECT pvStartPokFIO.propvalue, pvStartPokFIO.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokFIO" +
                    " WHERE pvStartPokFIO.lstypeid = 'ФЛ' AND pvStartPokFIO.periodyear = '" + peryear + "' AND pvStartPokFIO.periodmonth = '" + permonth + "' AND pvStartPokFIO.lspropertieid = '5' AND pvStartPokFIO.otdelenieid = '" + otdelenieid + "') pvStartPokFIO" +
                    " ON pvStartPok2.codeIESBK = pvStartPokFIO.codeIESBK" +

                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr1.propvalue, pvStartPokAddr1.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr1" +
                    " WHERE pvStartPokAddr1.lstypeid = 'ФЛ' AND pvStartPokAddr1.periodyear = '" + peryear + "' AND pvStartPokAddr1.periodmonth = '" + permonth + "' AND pvStartPokAddr1.lspropertieid = '11' AND pvStartPokAddr1.otdelenieid = '" + otdelenieid + "') pvStartPokAddr1" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr1.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr2.propvalue, pvStartPokAddr2.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr2" +
                    " WHERE pvStartPokAddr2.lstypeid = 'ФЛ' AND pvStartPokAddr2.periodyear = '" + peryear + "' AND pvStartPokAddr2.periodmonth = '" + permonth + "' AND pvStartPokAddr2.lspropertieid = '12' AND pvStartPokAddr2.otdelenieid = '" + otdelenieid + "') pvStartPokAddr2" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr2.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr3.propvalue, pvStartPokAddr3.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr3" +
                    " WHERE pvStartPokAddr3.lstypeid = 'ФЛ' AND pvStartPokAddr3.periodyear = '" + peryear + "' AND pvStartPokAddr3.periodmonth = '" + permonth + "' AND pvStartPokAddr3.lspropertieid = '13' AND pvStartPokAddr3.otdelenieid = '" + otdelenieid + "') pvStartPokAddr3" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr3.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr4.propvalue, pvStartPokAddr4.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr4" +
                    " WHERE pvStartPokAddr4.lstypeid = 'ФЛ' AND pvStartPokAddr4.periodyear = '" + peryear + "' AND pvStartPokAddr4.periodmonth = '" + permonth + "' AND pvStartPokAddr4.lspropertieid = '14' AND pvStartPokAddr4.otdelenieid = '" + otdelenieid + "') pvStartPokAddr4" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr4.codeIESBK" +
                    " LEFT JOIN" +
                    " (SELECT pvStartPokAddr5.propvalue, pvStartPokAddr5.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokAddr5" +
                    " WHERE pvStartPokAddr5.lstypeid = 'ФЛ' AND pvStartPokAddr5.periodyear = '" + peryear + "' AND pvStartPokAddr5.periodmonth = '" + permonth + "' AND pvStartPokAddr5.lspropertieid = '15' AND pvStartPokAddr5.otdelenieid = '" + otdelenieid + "') pvStartPokAddr5" +
                    " ON pvStartPok2.codeIESBK = pvStartPokAddr5.codeIESBK" +
                    //----------------------

                    // номер текущего ПУ
                    " LEFT JOIN" +
                    " (SELECT pvStartPokPUNomer.propvalue, pvStartPokPUNomer.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvStartPokPUNomer" +
                    " WHERE pvStartPokPUNomer.lstypeid = 'ФЛ' AND pvStartPokPUNomer.periodyear = '" + peryear + "' AND pvStartPokPUNomer.periodmonth = '" + permonth + "' AND pvStartPokPUNomer.lspropertieid = '7' AND pvStartPokPUNomer.otdelenieid = '" + otdelenieid + "') pvStartPokPUNomer" +
                    " ON pvStartPok2.codeIESBK = pvStartPokPUNomer.codeIESBK" +
                    //----------------------

                    " LEFT JOIN" +
                    " (SELECT pvEndPok2.propvalue, pvEndPok2.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPok2"+
                    " WHERE pvEndPok2.lstypeid = 'ФЛ' AND pvEndPok2.periodyear = '" + peryearprev + "' AND pvEndPok2.periodmonth = '" + permonthprev + "' AND pvEndPok2.lspropertieid = '25' AND pvEndPok2.otdelenieid = '" + otdelenieid + "') pvEndPok2" +
                    " ON pvStartPok2.codeIESBK = pvEndPok2.codeIESBK"+
                    " LEFT JOIN" +
                    " (SELECT pvEndPok1.propvalue, pvEndPok1.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPok1"+
                    " WHERE pvEndPok1.lstypeid = 'ФЛ' AND pvEndPok1.periodyear = '" + peryearprev + "' AND pvEndPok1.periodmonth = '" + permonthprev + "' AND pvEndPok1.lspropertieid = '24' AND pvEndPok1.otdelenieid = '" + otdelenieid + "') pvEndPok1" +
                    " ON pvStartPok2.codeIESBK = pvEndPok1.codeIESBK"+
                    " LEFT JOIN" +
                    " (SELECT pvEndPok3.propvalue, pvEndPok3.codeIESBK"+
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPok3"+
                    " WHERE pvEndPok3.lstypeid = 'ФЛ' AND pvEndPok3.periodyear = '" + peryearprev + "' AND pvEndPok3.periodmonth = '" + permonthprev + "' AND pvEndPok3.lspropertieid = '26' AND pvEndPok3.otdelenieid = '" + otdelenieid + "') pvEndPok3" +
                    " ON pvStartPok2.codeIESBK = pvEndPok3.codeIESBK"+

                    // номер предыдущего ПУ
                    " LEFT JOIN" +
                    " (SELECT pvEndPokPUNomer.propvalue, pvEndPokPUNomer.codeIESBK" +
                    " FROM[iesbk].[dbo].[tblIESBKlspropvalue] pvEndPokPUNomer" +
                    " WHERE pvEndPokPUNomer.lstypeid = 'ФЛ' AND pvEndPokPUNomer.periodyear = '" + peryearprev + "' AND pvEndPokPUNomer.periodmonth = '" + permonthprev + "' AND pvEndPokPUNomer.lspropertieid = '7' AND pvEndPokPUNomer.otdelenieid = '" + otdelenieid + "') pvEndPokPUNomer" +
                    " ON pvStartPok2.codeIESBK = pvEndPokPUNomer.codeIESBK" +
                    //----------------------


                    " WHERE pvStartPok2.lstypeid = 'ФЛ' AND pvStartPok2.periodyear = '" + peryear + "' AND pvStartPok2.periodmonth = '" + permonth + "' AND pvStartPok2.lspropertieid = '22' AND pvStartPok2.otdelenieid = '" + otdelenieid + "'" +
                    " AND pvStartPok2.propvalue <> pvEndPok2.propvalue";

                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                //-----------------------------------------------------------               
                
                for (int j = 0; j < tableTOTALls.Rows.Count; j++)
                {
                    worksheet[strow + j + 1, stcol + 0].SetValue(tableTOTALls.Rows[j]["codels"].ToString());

                    worksheet[strow + j + 1, stcol + 1].SetValue(tableTOTALls.Rows[j]["pvStartPokFIOvalue"]);

                    worksheet[strow + j + 1, stcol + 2].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr1value"]);
                    worksheet[strow + j + 1, stcol + 3].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr2value"]);
                    worksheet[strow + j + 1, stcol + 4].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr3value"]);
                    worksheet[strow + j + 1, stcol + 5].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr4value"]);
                    worksheet[strow + j + 1, stcol + 6].SetValue(tableTOTALls.Rows[j]["pvStartPokAddr5value"]);

                    // информация о предыдущем периоде
                    string date_prev_str = tableTOTALls.Rows[j]["prevPOKdatePrevPer"].ToString();                    
                    worksheet[strow + j + 1, stcol + 7].SetValue(date_prev_str);
                    worksheet[strow + j + 1, stcol + 7].Font.Color = infocolor_prev;

                    string value_prev_str = tableTOTALls.Rows[j]["prevPOKvaluePrevPer"].ToString();
                    worksheet[strow + j + 1, stcol + 8].SetValue(value_prev_str);
                    worksheet[strow + j + 1, stcol + 8].Font.Color = infocolor_prev;

                    string nomerPU_prev_str = tableTOTALls.Rows[j]["prevPokPUNomervalue"].ToString();
                    worksheet[strow + j + 1, stcol + 9].SetValue(nomerPU_prev_str);
                    worksheet[strow + j + 1, stcol + 9].Font.Color = infocolor_prev;

                    worksheet[strow + j + 1, stcol + 10].SetValue(tableTOTALls.Rows[j]["prevPOKtypePrevPer"]);
                    worksheet[strow + j + 1, stcol + 10].Font.Color = infocolor_prev;
                    //-------------------------------

                    // информация о текущем периоде
                    string date_tek_str = tableTOTALls.Rows[j]["startPOKdateTekPer"].ToString();
                    worksheet[strow + j + 1, stcol + 11].SetValue(date_tek_str);

                    string value_tek_str = tableTOTALls.Rows[j]["startPOKvalueTekPer"].ToString();
                    worksheet[strow + j + 1, stcol + 12].SetValue(value_tek_str);

                    string nomerPU_tek_str = tableTOTALls.Rows[j]["startPokPUNomervalue"].ToString();
                    worksheet[strow + j + 1, stcol + 13].SetValue(nomerPU_tek_str);

                    worksheet[strow + j + 1, stcol + 14].SetValue(tableTOTALls.Rows[j]["startPOKtypeTekPer"]);
                    //-------------------------------

                    worksheet[strow + j + 1, stcol + 8].NumberFormat = "#####";
                    worksheet[strow + j + 1, stcol + 12].NumberFormat = "#####";

                    // формируем и выводим статусы                 
                    bool flag_value_error = false; // нарушение нарастающего значения
                    bool flag_date_error = false; // нарушение нарастающей даты снятия показания

                    // если номер ПУ одинаковый
                    if (nomerPU_tek_str == nomerPU_prev_str)
                    {
                        if (date_tek_str != date_prev_str) flag_date_error = true;
                        if (value_tek_str != value_prev_str) flag_value_error = true;
                    }

                    if (flag_value_error) worksheet[strow + j + 1, stcol + 15].SetValue("да");
                    worksheet[strow + j + 1, stcol + 15].Font.Color = Color.Red;

                    if (flag_date_error) worksheet[strow + j + 1, stcol + 16].SetValue("да");
                    worksheet[strow + j + 1, stcol + 16].Font.Color = Color.Blue;
                    //----------------------------
                }

                // форматируем строку-заголовок
                worksheet.Rows[strow].Font.Bold = true;
                worksheet.Rows[strow].Alignment.WrapText = true;
                worksheet.Rows[strow].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[strow].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[strow].AutoFit();

                worksheet.Columns.AutoFit(0, MAXCOLinWRKSH);
                worksheet.Columns.Group(2, 6, true);
                worksheet.FreezeRows(strow);                
            } // for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)

            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
            splashScreenManager1.CloseWaitForm();

            form1.spreadsheetControl1.EndUpdate();

            form1.Show();
        } // ИЭСБК отчеты - "Нарушение нарастающего итога"

        public static string MyFUNC_GetPropValueFromIESBKOLAP2(int IESBKlsid, string periodyear, string periodmonth, int lspropidglobal,
            SqlConnection connection/*, out string pvstr, out double? pvdouble, out DateTime? pvdate*/)
        {
            string result = null;

            // получаем основной тип свойства
            /*string queryString =
                    "SELECT valuetypeid " +
                    "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                    "WHERE lspropidglobal = " + lspropidglobal.ToString();
            DataTable tablelsPropGLOBAL = new DataTable();
            MyFUNC_SelectDataFromSQLwoutConnection(tablelsPropGLOBAL, connection, queryString);            */

            DataRow[] lsproprows = tablelsPropGLOBAL.Select("lspropidglobal = " + lspropidglobal.ToString());
            int valuetypeid = Convert.ToInt32(lsproprows[0]["valuetypeid"].ToString());
            //int valuetypeid = Convert.ToInt32(tablelsPropGLOBAL.Rows[0]["valuetypeid"].ToString());

            //tablelsPropGLOBAL.Dispose();
            //-------------------------------

            DateTime period = Convert.ToDateTime("01." + periodmonth + "." + periodyear);

            // определяем таблицу назначения
            string propvaluetabledest = null;
            if (valuetypeid == 1)
            {
                propvaluetabledest = "tblIESBKlspropvaluestr"; // текстовый
                //result = "text";
            }
            else if (valuetypeid == 2)
            {
                propvaluetabledest = "tblIESBKlspropvaluenum"; // числовой
                //result = "123";
            }
            else if (valuetypeid == 3)
            {
                propvaluetabledest = "tblIESBKlspropvaluedate"; // дата
                //result = "01.12.1983";
            }
            //----------------------------

            string queryStringPropValue = "SELECT propvalue " +
                                          "FROM iesbk2.dbo." + propvaluetabledest +
                                          " WHERE lspropidglobal = " + lspropidglobal.ToString() + " AND period = '" + period.ToShortDateString() + "' AND IESBKlsid = " + IESBKlsid.ToString();
            //" WHERE IESBKlsid = " + IESBKlsid.ToString() + " AND lspropidglobal = " + lspropidglobal.ToString() + " AND period = '" + period.ToShortDateString() +"'";
            DataTable tablePropValue = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(connection, tablePropValue, queryStringPropValue);

            //pvstr = null; pvdate = null; pvdouble = null; // первоначальное обнуление возвращаемых значений

            if (tablePropValue.Rows.Count != 0)
            {
                result = tablePropValue.Rows[0]["propvalue"].ToString(); // формируем текстовое представление возвращаемого значения

                /*if (valuetypeid == 1) // текстовый
                {
                    pvstr = result; 
                }
                else if (valuetypeid == 2) // числовой
                {
                    pvdouble = Convert.ToDouble(tablePropValue.Rows[0]["propvalue"]); 
                }
                else if (valuetypeid == 3) // дата
                {
                    pvdate = Convert.ToDateTime(tablePropValue.Rows[0]["propvalue"]);
                } */
            }
            else if (valuetypeid != 1) // если значение свойства не найдено в основной не текстовой таблице, то смотрим в текстовой
            {
                queryStringPropValue = "SELECT propvalue " +
                                       "FROM iesbk2.dbo.tblIESBKlspropvaluestr" +
                                       " WHERE lspropidglobal = " + lspropidglobal.ToString() + " AND period = '" + period.ToShortDateString() + "' AND IESBKlsid = " + IESBKlsid.ToString();
                DataTable tablePropValue2 = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(connection, tablePropValue2, queryStringPropValue);

                if (tablePropValue2.Rows.Count != 0) result = tablePropValue2.Rows[0]["propvalue"].ToString();

                tablePropValue2.Dispose();
            }

            tablePropValue.Dispose();

            return result;
        }

        /*// пока БЕСПОЛЕЗНАЯ функция
        // функция возврата текстовых значений года и месяца в смещении от заданных
        // параметр offset считается "в месяцах" (+/-)
        public static DateTime MyFUNC_GetPeriodValuesOffset(string year_tek, string month_tek, int offset)
        {
            DateTime dt_tek = Convert.ToDateTime("01." + month_tek + "." + year_tek);
            DateTime dt_new = dt_tek.AddMonths(offset);
                        
            return dt_new;
        }*/

        // тест "месяца"
        private void barButtonItem39_ItemClick(object sender, ItemClickEventArgs e)
        {
            //textBox1.Text = DateTime.Now.Month.ToString();

            /*SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            textBox1.Text = MyFUNC_GetPropValueFromIESBKOLAP("КООО0016410", "2016", "06", 30, SQLconnection);

            SQLconnection.Close();

            DateTime dt = DateTime.Now;            
            dt = dt.AddMonths(-1);
            textBox1.Text = dt.ToString();*/

            DateTime dt = Convert.ToDateTime("10.08.2015");
            DateTime dt2 = Convert.ToDateTime("01.04.2016");

            textBox1.Text = dt.CompareTo(dt2).ToString();
        }

        // "шахматка" по среднемесячному
        private void barButtonItem40_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();
            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Проверка среднемесячного начисления";
            IWorkbook workbook = form1.spreadsheetControl1.Document;

            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            // выборка ЛС за текущий период 2016 по текущему отделению и по тем л/с, где поле Среднемесячное заполнено ---------------------------------------            
            DateTime dt_IESBK_MIN = Convert.ToDateTime("01.01.2015"); // левая граница имеющихся данных в OLAP-кубе

            //int MAX_PERIOD_MONTH = 12;
            int MAX_PERIOD_MONTH = 10;

            string yeartek = "2018";
            string monthtek = "10";
            string queryStringsrmes =
                                 "SELECT otdelenieid,codeIESBK,propvalue " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                 "WHERE" +
                                 //" otdelenieid = 'ВО' AND" + // ТЕСТ !!!!
                                 //" codeIESBK = '35000000041' AND" +
                                 " periodyear = '" + yeartek + "' AND periodmonth = '" + monthtek + "' AND lspropertieid='30'" + " AND propvalue IS NOT NULL";
            DataTable tablePOsrmestekPERIOD = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tablePOsrmestekPERIOD, queryStringsrmes);
            //------------------------------------------------------------------------------------

            Worksheet worksheet = workbook.Worksheets[0];
            //worksheet.Name = periodyeartek + ", " + periodmonthtek;

            // задаем смещение от левого угла листа книги            
            int strow = 0;
            int stcol = 0;
                        
            worksheet[strow + 0, stcol + 0].SetValue("Отделение ИЭСБК");
            worksheet[strow + 0, stcol + 1].SetValue("Код л/с ИЭСБК");

            worksheet[strow + 0, stcol + 2].SetValue("ФИО");
            worksheet[strow + 0, stcol + 3].SetValue("Населенный пункт");
            worksheet[strow + 0, stcol + 4].SetValue("Улица");
            worksheet[strow + 0, stcol + 5].SetValue("Дом");
            worksheet[strow + 0, stcol + 6].SetValue("Номер квартиры");

            // "пробегаем" по всем лицевым счетам выборки текущего периода            
            int rd = 1;
            int columns_in_period = 10+1;

            for (int i = 0; i < tablePOsrmestekPERIOD.Rows.Count; i++)
            //for (int i = 0; i < 500; i++)
            {
                string otdelenieid = tablePOsrmestekPERIOD.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.FindByotdelenieid(otdelenieid)["captionotd"].ToString();
                string codels = tablePOsrmestekPERIOD.Rows[i]["codeIESBK"].ToString();

                worksheet[strow + rd + 0, stcol + 0].SetValue(captionotd); // отделение ИЭСБК
                worksheet[strow + rd + 0, stcol + 1].SetValue(codels); // код ИЭСБК л/с
                
                // "бежим" по периодам с 01 по 09 2016 по выборке текущего периода -------------------                
                for (int period_i = 1; period_i <= MAX_PERIOD_MONTH; period_i++) //!!!
                {
                    //string periodyeartek = "2016";
                    string periodyeartek = "2018";
                    string periodmonthtek = (period_i < 10) ? "0" + period_i.ToString() : period_i.ToString();
                    
                    // выводим справочную информации на основе данных начального периода интервала
                    if (period_i == MAX_PERIOD_MONTH)
                    {
                        // выбираем все свойства л/с периода из базы
                        string queryStringlsprop = "SELECT lspropertieid, propvalue " +
                                         "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                         "WHERE codeIESBK='" + codels + "' AND periodyear = '" + periodyeartek + "' AND periodmonth = '" + periodmonthtek + "'";
                        DataTable tableLSprop = new DataTable();
                        MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

                        // ФИО
                        DataRow[] lsproprows = tableLSprop.Select("lspropertieid = '5'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 2].SetValue(lsproprows[0]["propvalue"].ToString());

                        // населенный пункт
                        lsproprows = tableLSprop.Select("lspropertieid = '12'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 3].SetValue(lsproprows[0]["propvalue"].ToString());

                        // улица
                        lsproprows = tableLSprop.Select("lspropertieid = '13'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 4].SetValue(lsproprows[0]["propvalue"].ToString());

                        // дом
                        lsproprows = tableLSprop.Select("lspropertieid = '14'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 5].SetValue(lsproprows[0]["propvalue"].ToString());

                        // номер квартиры
                        lsproprows = tableLSprop.Select("lspropertieid = '15'");
                        if (lsproprows.Length > 0) worksheet[strow + rd + 0, stcol + 6].SetValue(lsproprows[0]["propvalue"].ToString());

                        tableLSprop.Dispose();
                    }

                    // выводим заголовки столбцов инфо о периоде
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(periodmonthtek + " Дата ПредПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(periodmonthtek + " ПредПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(periodmonthtek + " Вид ПредПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(periodmonthtek + " Дата ПослПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(periodmonthtek + " ПослПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(periodmonthtek + " Вид ПослПок");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(periodmonthtek + " СрМес РАСЧ");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 14].SetValue(periodmonthtek + " СрМес РАСЧ Дельта дней");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 14].Font.Color = Color.Green;

                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 15].SetValue(periodmonthtek + " ПолОтп ИЭСБК");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 15].Font.Color = Color.Blue;

                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(periodmonthtek + " СрМес ИЭСБК");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Blue;

                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 17].SetValue(periodmonthtek + " Недоп ПО");
                    worksheet[strow + 0, stcol + (period_i - 1) * columns_in_period + 17].Font.Color = Color.Red;

                    // СрМес ПО ИЭСБК текущего периода
                    string po_iesbk_str = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, periodyeartek, periodmonthtek, 27, SQLconnection);
                    string srmes_iesbk_str = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, periodyeartek, periodmonthtek, 30, SQLconnection);
                    double? srmes_iesbk = null;
                    double? po_iesbk = null;

                    if (!String.IsNullOrWhiteSpace(po_iesbk_str) && !po_iesbk_str.Contains(";"))
                    {
                        po_iesbk = Convert.ToDouble(po_iesbk_str);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 15].Font.Color = Color.Blue;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 15].SetValue(po_iesbk); // Полезный отпуск ИЭСБК
                    }

                    if (!String.IsNullOrWhiteSpace(srmes_iesbk_str) && !srmes_iesbk_str.Contains(";"))
                    {
                        srmes_iesbk = Convert.ToDouble(srmes_iesbk_str);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Blue;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(srmes_iesbk); // СрМес ИЭСБК
                    }

                    // ищем ближайшее "правое" показание                
                    string value_right = null;
                    string dtvalue_right = null;

                    DateTime dt_right = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);
                    string year_right = null;
                    string month_right = null;

                    dt_right = dt_right.AddMonths(+1); // учитываем текущий месяц, т.е. +1-1 = 0

                    while (String.IsNullOrWhiteSpace(value_right) && dt_right >= dt_IESBK_MIN)
                    {
                        dt_right = dt_right.AddMonths(-1);
                        year_right = dt_right.Year.ToString();
                        month_right = null;
                        if (dt_right.Month < 10) month_right = "0" + dt_right.Month.ToString();
                        else month_right = dt_right.Month.ToString();

                        value_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 25, SQLconnection); // свойство "Текущее показание ПУ"                                        
                    };
                    //----------------------------------

                    // ищем "левое" показание, при условии, что нашли "правое" ------------------
                    string value_left = null;
                    string dtvalue_left = null;
                    string year_left = null;
                    string month_left = null;

                    if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    {
                        // получаем дату "правого" показания
                        dtvalue_right = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                        // выводим информацию о "правом" показании
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(dtvalue_right); // дата

                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(value_right); // показание

                        string rightpok_type = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_right, month_right, 26, SQLconnection); 
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(rightpok_type); // вид

                        DateTime dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-5); // БЫЛО -6, отматываем 6 мес. (-5-1 = -6) = 180 дней от "правого" показания

                        while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN)
                        {
                            dt_left = dt_left.AddMonths(-1);
                            year_left = dt_left.Year.ToString();
                            month_left = null;
                            if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                            else month_left = dt_left.Month.ToString();

                            value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                        };

                        // если нет данных за период не менее 6 мес., то ищем за в периоде [6 мес.;3 мес.]
                        if (String.IsNullOrWhiteSpace(value_left))
                        {
                            dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-7); // отматываем 7 мес., т.к. в теле цикла сразу +1, т.е. -7+1 = -6

                            DateTime dt_IESBK_left_MAX = Convert.ToDateTime(dtvalue_right).AddMonths(-3);

                            while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN && dt_left < dt_IESBK_left_MAX)
                            {
                                dt_left = dt_left.AddMonths(+1);
                                year_left = dt_left.Year.ToString();
                                month_left = null;
                                if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                                else month_left = dt_left.Month.ToString();

                                value_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 25, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                            };
                        } // if (String.IsNullOrWhiteSpace(value_left)) // если нет данных за период не менее 6 мес.

                        // получаем дату "левого" показания и выводим информацию о нем
                        if (!String.IsNullOrWhiteSpace(value_left))
                        {
                            dtvalue_left = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 24, SQLconnection); // свойство "Дата последнего показания ПУ"

                            // выводим информацию о "правом" показании
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(dtvalue_left); // дата

                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(value_left); // показание

                            string leftpok_type = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codels, year_left, month_left, 26, SQLconnection);
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(leftpok_type); // вид
                        }

                        // если даты "левого" и "правого" показаний не пустые, то формируем расчет среднемесячного
                        if (!String.IsNullOrWhiteSpace(dtvalue_left) && !dtvalue_left.Contains(";") && !String.IsNullOrWhiteSpace(dtvalue_right))
                        {
                            double pokleft = Convert.ToDouble(value_left);
                            double pokright = Convert.ToDouble(value_right);

                            // если не нарушен нарастающий итог
                            if (pokleft <= pokright)
                            {
                                System.TimeSpan deltaday = Convert.ToDateTime(dtvalue_right) - Convert.ToDateTime(dtvalue_left);
                                double deltapok = pokright - pokleft;

                                double srednesut_calc = deltapok / deltaday.Days;
                                double srmes_calc = Math.Round(srednesut_calc * DateTime.DaysInMonth(Convert.ToInt32(periodyeartek), Convert.ToInt32(periodmonthtek)));

                                // формируем отчет -----------------------------------------------
                                                                
                                if (!String.IsNullOrWhiteSpace(srmes_iesbk_str)) // еслм СрМес ПО ИЭСБК присутствует
                                {
                                    DateTime dt_period = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);

                                    if (srmes_iesbk >= 0 && Convert.ToDateTime(dtvalue_right).CompareTo(dt_period) < 0) // не выводим наши расчеты, если СрМес ИЭСБК < 0 и правое показание принадлежит текущему периоду анализа
                                    {
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(srmes_calc); // СрМес РАСЧ                                

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 14].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 14].SetValue(deltaday.Days); // СрМес РАСЧ Дельта дней

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 17].Font.Color = Color.Red;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 17].SetValue(srmes_calc - srmes_iesbk); // Недополученный ПО                                
                                    }
                                }
                                                                
                            } // if (pokleft <= pokright)
                              //-----------------------------------------------------------------
                        } // if (!String.IsNullOrWhiteSpace(dtvalue_left) && !String.IsNullOrWhiteSpace(dtvalue_right))


                    } // if (!String.IsNullOrWhiteSpace(value_right))
                      //----------------------------------
                      
                    splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " из " + tablePOsrmestekPERIOD.Rows.Count.ToString() + ")" + "-период " + period_i.ToString());

                }  // for (int period_i = 1; period_i < 7; period_i++)
                
                rd++;

            } // for (int i = 0; i < tablePOsrmestekPERIOD.Rows.Count; i++)

            //workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];

            worksheet.Columns.AutoFit(0, 7+10 * MAX_PERIOD_MONTH); // ПЕРЕДЕЛАТЬ для универсальности
            form1.spreadsheetControl1.EndUpdate();

            SQLconnection.Close();
            tablePOsrmestekPERIOD.Dispose();
            splashScreenManager1.CloseWaitForm();

            form1.Show();
        } // "шахматка" по среднемесячному

        // отчет на "разрыв" показаний (модернизированный "Нарушение нарастающего итога"), когда последние показания предыдущего расчетного периода не равны начальным текущего
        private void barButtonItem41_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Нарушение нарастающего итога";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            
            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            //for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            for (int i = 0; i < 1; i++)
            {
                string otdelenieid = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();

                Worksheet worksheet = workbook.Worksheets[i];
                worksheet.Name = captionotd;
                workbook.Worksheets.Add();

                // задаем смещение табличной части отчета
                int strow = 5;
                int stcol = 0;

                string peryearprev = "2016";
                string permonthprev = "06";

                string peryeartek = "2016";
                string permonthtek = "07";

                worksheet[0, 0].SetValue("ФЛ без ОДПУ");
                worksheet[1, 0].SetValue(captionotd);
                worksheet[2, 0].SetValue("Текущий период:");
                worksheet[2, 1].SetValue(peryeartek + ", " + permonthtek);
                worksheet[3, 0].SetValue("Предыдущий период:");
                worksheet[3, 1].SetValue(peryearprev + ", " + permonthprev);

                worksheet[strow + 0, stcol + 0].SetValue("Код ИЭСБК");
                worksheet[strow + 0, stcol + 1].SetValue("Пред период посл пок дата");
                worksheet[strow + 0, stcol + 2].SetValue("Пред период посл пок значение");
                worksheet[strow + 0, stcol + 3].SetValue("Пред период посл пок источник");
                worksheet[strow + 0, stcol + 4].SetValue("Тек период нач пок дата");
                worksheet[strow + 0, stcol + 5].SetValue("Тек период нач пок значение");
                worksheet[strow + 0, stcol + 6].SetValue("Тек период нач пок источник");

                worksheet[strow + 0, stcol + 7].SetValue("Нарушение по дате");
                worksheet[strow + 0, stcol + 8].SetValue("Разница показаний (НЕ ДОЛЖНО БЫТЬ)");

                string queryString =
                                 "SELECT otdelenieid,codeIESBK,propvalue " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                 "WHERE " +
                                 "periodyear = '" + peryeartek + "' AND periodmonth = '" + permonthtek + "'" + " AND otdelenieid = '" + otdelenieid + "' AND lspropertieid = '3'" + " AND propvalue IS NOT NULL";
                
                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                //-----------------------------------------------------------               
                int row_rep = 0;
                for (int j = 0; j < tableTOTALls.Rows.Count; j++)                
                {
                    string codeIESBK = tableTOTALls.Rows[j]["codeIESBK"].ToString();
                    
                    // информация о предыдущем периоде
                    string tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryearprev, permonthprev, 24, SQLconnection); // дата последнего показания периода
                    DateTime? POKprev_date_last = null;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKprev_date_last = Convert.ToDateTime(tempstr);
                    
                    tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryearprev, permonthprev, 25, SQLconnection); // значение последнего показания периода
                    double POKprev_value_last = -1;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKprev_value_last = Convert.ToDouble(tempstr);
                    
                    string POKprev_type_last = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryearprev, permonthprev, 26, SQLconnection); // вид последнего показания периода                    
                    

                    // информация о текущем периоде
                    tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryeartek, permonthtek, 21, SQLconnection); // дата предыдущего показания периода
                    DateTime? POKtek_date_first = null;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKtek_date_first = Convert.ToDateTime(tempstr);
                    
                    tempstr = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryeartek, permonthtek, 22, SQLconnection); // значение предыдущего показания периода
                    double POKtek_value_first = -1;
                    if (!String.IsNullOrWhiteSpace(tempstr) && !tempstr.Contains(";")) POKtek_value_first = Convert.ToDouble(tempstr);
                    
                    string POKtek_type_first = MC_SQLDataProvider.GetPropValueFromIESBKOLAP(codeIESBK, peryeartek, permonthtek, 23, SQLconnection); // вид предыдущего показания периода                    

                    // нарушение может еще быть и по дате!!!!! ИСПРАВИТЬ
                    // + замена ПУ (тоже учесть, на всякий случай (кроме ";")

                    if (POKprev_value_last != POKtek_value_first && POKprev_value_last != -1 && POKtek_value_first != -1)
                    {
                        worksheet[strow + row_rep + 1, stcol + 0].SetValue(codeIESBK);

                        worksheet[strow + row_rep + 1, stcol + 1].SetValue(POKprev_date_last);
                        worksheet[strow + row_rep + 1, stcol + 2].SetValue(POKprev_value_last);
                        worksheet[strow + row_rep + 1, stcol + 3].SetValue(POKprev_type_last);

                        worksheet[strow + row_rep + 1, stcol + 4].SetValue(POKtek_date_first);
                        worksheet[strow + row_rep + 1, stcol + 5].SetValue(POKtek_value_first);
                        worksheet[strow + row_rep + 1, stcol + 6].SetValue(POKtek_type_first);

                        worksheet[strow + row_rep + 1, stcol + 7].SetValue(POKtek_date_first < POKprev_date_last ? "да" : null);
                        double deltaPOK = POKtek_value_first - POKprev_value_last;
                        worksheet[strow + row_rep + 1, stcol + 8].SetValue(deltaPOK);

                        row_rep++;
                    }

                    splashScreenManager1.SetWaitFormDescription("Обработка (" + (i+1).ToString() + " - " + (j + 1).ToString() + " из " + tableTOTALls.Rows.Count.ToString() + ")");
                }

                worksheet.Columns.AutoFit(0, 10+2);
            } // for (int j = 0; j < tableTOTALls.Rows.Count; j++)

            SQLconnection.Close();

            form1.spreadsheetControl1.EndUpdate();

            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets[0];
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        }

        // перенос данных "Свойства"
        private void barButtonItem42_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK);
            SQLconnection.Open();
            
            // читаем "старые" свойства
            string queryStringlsprop = "SELECT lspropertieid, templateid, numcolumninfile, captionlsprop " +
                                        "FROM [iesbk].[dbo].[tblIESBKlsprop]";
            DataTable tableLSprop = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

            for (int i = 0; i < tableLSprop.Rows.Count; i++)
            {
                int templateid = 1;
                int lspropidtmpl = Convert.ToInt32(tableLSprop.Rows[i]["lspropertieid"].ToString());
                int lspropidglobal = templateid * 1000 + lspropidtmpl;

                string numcolumninfilestr = tableLSprop.Rows[i]["numcolumninfile"].ToString();
                int? numcolumninfile = null;
                if (!String.IsNullOrWhiteSpace(numcolumninfilestr)) numcolumninfile = Convert.ToInt32(numcolumninfilestr);
                                
                string captionlsprop = tableLSprop.Rows[i]["captionlsprop"].ToString();                

                string queryStringlsprop2 = "INSERT INTO iesbk2.dbo.tblIESBKlsprop(lspropidglobal, templateid, lspropidtmpl, numcolumninfile, captionlsprop, comment, valuetypeid) " +
                    "VALUES (" +
                    lspropidglobal.ToString() + "," +
                    templateid.ToString() + "," +
                    lspropidtmpl.ToString() + "," +                    
                    "NULL" + "," +
                    "'" + captionlsprop + "'" + "," +
                    "NULL" + "," + 
                    "-100" + ")";                    
                MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);

                /*queryStringlsprop2 = "UPDATE[iesbk2].[dbo].[tblIESBKlsprop]";
                MyFUNC_RunSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);*/

                splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " из " + tableLSprop.Rows.Count.ToString() + ")");
            }

            tableLSprop.Dispose();

            SQLconnection.Close();
            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Свойства"

        // перенос данных "Лицевые счета"
        private void barButtonItem43_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            // загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);

            // переносим по отделениям
            for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)
            {
                string otdelenieidstr = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["otdelenieid"].ToString();
                string captionotd = DataSetIESBKLoad.tblIESBKotdelenie.Rows[i]["captionotd"].ToString();

                string queryString =
                    "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded " +
                    "FROM [iesbk].[dbo].[tblIESBKls] " +
                    "WHERE otdelenieid = '" + otdelenieidstr.ToString() + "'";
                DataTable tableTOTALls = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

                //-----------------------------------------------------------               

                string queryString2 =
                    "SELECT otdelenieid, captionotd " +
                    "FROM [iesbk2].[dbo].[tblIESBKotdelenie] " +
                    "WHERE captionotd = '" + captionotd + "'";
                DataTable table2 = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(table2, dbconnectionStringIESBK2, queryString2);
                int otdelenieid = Convert.ToInt32(table2.Rows[0]["otdelenieid"].ToString());
                table2.Dispose();

                int lstypeid = 1; // фл без одпу
                int j = 0;
                for (j = 0; j < tableTOTALls.Rows.Count; j++)
                {
                    string codeIESBK = tableTOTALls.Rows[j]["codeIESBK"].ToString();
                    
                    DateTime dateloaded = Convert.ToDateTime(tableTOTALls.Rows[j]["dateloaded"].ToString());
                    int IESBKlsid = lstypeid * 100000000 + otdelenieid * 1000000 + j + 1;
                    //------------------------

                    string queryStringlsprop2 = "INSERT INTO iesbk2.dbo.tblIESBKls(codeIESBK, lstypeid, otdelenieid, dateloaded, IESBKlsid) " +
                        "VALUES (" +
                        "'" + codeIESBK + "'" + "," +
                        lstypeid.ToString() + "," +
                        otdelenieid.ToString() + "," +
                        "'" + dateloaded.ToString() + "'" + "," +
                        IESBKlsid.ToString() + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);

                    splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " - " + (j + 1).ToString() + " из " + tableTOTALls.Rows.Count.ToString() + ")");
                }

                // записываем последний lastlsnumber в таблицу [tblIESBKlastlsid]
                int lastlsid = lstypeid * 100000000 + otdelenieid * 1000000 + j;
                queryString2 = "UPDATE iesbk2.dbo.tblIESBKlastlsid " +
                                "SET lastlsid = " + lastlsid.ToString() +
                                "WHERE otdelenieid = " + otdelenieid.ToString();
                MC_SQLDataProvider.UpdateSQLQuery(dbconnectionStringIESBK2, queryString2);
                //-------------------------------------------------------------------

                tableTOTALls.Dispose();

            } // for (int i = 0; i < DataSetIESBKLoad.tblIESBKotdelenie.Rows.Count; i++)

            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Лицевые счета"

        //-----------------------------------------------

        // перенос данных "Значения свойств"
        private void barButtonItem44_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            // параметры фильтра для "многозадачности" (запуск нескольких приложений)
            string peryear = textBox1.Text;
            string permonth = textBox3.Text;

            // формируем выборку лицевых счетов
            string queryString =
                    "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded " +
                    "FROM [iesbk].[dbo].[tblIESBKls]";                    
            DataTable tableTOTALls = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls, dbconnectionStringIESBK, queryString);

            // "бежим" по выборке
            for (int i = 0; i < tableTOTALls.Rows.Count; i++)
            {
                string codeIESBK = tableTOTALls.Rows[i]["codeIESBK"].ToString();
                string otdelenieid_old = tableTOTALls.Rows[i]["otdelenieid"].ToString();

                // собираем все свойства "старого" лицевого счета через фильтр Код+Отделение+Год+Месяц (формируем фильтр-выборку)
                queryString =
                    "SELECT propvalue, periodyear, periodmonth, codeIESBK, lstypeid, lspropertieid, templateid, otdelenieid " +
                    "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                    "WHERE codeIESBK = '" + codeIESBK + "' AND otdelenieid = '" + otdelenieid_old + "'"
                    + " AND periodyear = '" + peryear + "'" + " AND periodmonth = '" + permonth + "'";
                DataTable tablelsprop = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsprop, dbconnectionStringIESBK, queryString);

                // получаем "новый" id отделения
                int otdelenieid_new = -1;
                if (otdelenieid_old == "АО") otdelenieid_new = 10;
                else if (otdelenieid_old == "ВО") otdelenieid_new = 11;
                else if (otdelenieid_old == "КО") otdelenieid_new = 12;
                else if (otdelenieid_old == "МЧО") otdelenieid_new = 13;
                else if (otdelenieid_old == "СлО") otdelenieid_new = 14;
                else if (otdelenieid_old == "СОЗ") otdelenieid_new = 15;
                else if (otdelenieid_old == "СОС") otdelenieid_new = 16;
                else if (otdelenieid_old == "ТОН") otdelenieid_new = 17;
                else if (otdelenieid_old == "ТОТ") otdelenieid_new = 18;
                else if (otdelenieid_old == "ТшО") otdelenieid_new = 19;
                else if (otdelenieid_old == "УКО") otdelenieid_new = 20;
                else if (otdelenieid_old == "УсО") otdelenieid_new = 21;
                else if (otdelenieid_old == "ЦО") otdelenieid_new = 22;
                else if (otdelenieid_old == "ЧО") otdelenieid_new = 23;
                //-------------------------------------------------

                // получаем id "нового" лицевого счета (учитывая отделение)
                queryString =
                    "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded, IESBKlsid "+
                    "FROM [iesbk2].[dbo].[tblIESBKls] " +
                    "WHERE codeIESBK = '" + codeIESBK + "' AND otdelenieid = '" + otdelenieid_new.ToString() + "'";
                DataTable tablelsnew = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                int IESBKlsid = Convert.ToInt32(tablelsnew.Rows[0]["IESBKlsid"].ToString());
                tablelsnew.Dispose();
                //-------------------------------------------------                

                // переносим свойства лицевого счета из фильтр-выборки
                for (int j = 0; j < tablelsprop.Rows.Count; j++)
                {                    
                    int templateid = 1; // 3.3.2
                    DateTime period = Convert.ToDateTime("01." + tablelsprop.Rows[j]["periodmonth"].ToString() + "." + tablelsprop.Rows[j]["periodyear"].ToString());

                    int lspropertieid = Convert.ToInt32(tablelsprop.Rows[j]["lspropertieid"].ToString());
                    string propvaluestr_src = tablelsprop.Rows[j]["propvalue"].ToString();

                    // получаем "новый" тип свойства
                    queryString =
                        "SELECT lspropidglobal, valuetypeid " +
                        "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                        "WHERE templateid = " + templateid.ToString() +
                        " AND lspropidtmpl = " + lspropertieid.ToString();
                    tablelsnew = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                    int valuetypeid = Convert.ToInt32(tablelsnew.Rows[0]["valuetypeid"].ToString());
                    int lspropidglobal = Convert.ToInt32(tablelsnew.Rows[0]["lspropidglobal"].ToString());
                    tablelsnew.Dispose();
                    //------------------------------

                    // если в значении свойства присутствует символ ";" (замена ПУ), то 
                    // загружаем в текстовый куб
                    if (propvaluestr_src.Contains(";")) valuetypeid = 1;

                    // очистка старых косяков
                    if (propvaluestr_src.Contains("REF!")) propvaluestr_src = null;
                    //-----------------------------------------------------------------

                    string propvaluetabledest = null;
                    if (valuetypeid == 1) propvaluetabledest = "tblIESBKlspropvaluestr"; // текстовый
                    else if (valuetypeid == 2) propvaluetabledest = "tblIESBKlspropvaluenum"; // числовой
                    else if (valuetypeid == 3) propvaluetabledest = "tblIESBKlspropvaluedate"; // дата
                    //----------------------                    

                    string propvaluestr_dest_for_insert = "NULL";
                    if (!String.IsNullOrWhiteSpace(propvaluestr_src))
                    {
                        if (valuetypeid == 1) propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                        else if (valuetypeid == 2)
                        {
                            //propvaluestr_dest_for_insert = Convert.ToDouble(propvaluestr_src).ToString().Replace(",","."); //? зачем конвертить

                            /*double doubleValue;                            
                            if (Double.TryParse(propvaluestr_src, out doubleValue)) propvaluestr_dest_for_insert = propvaluestr_src.Replace(",", ".");
                            else propvaluestr_dest_for_insert = "NULL";*/ // пока парсинг убрал - для увеличения производительности, при загрузке нужен обязательно!!!

                            propvaluestr_dest_for_insert = propvaluestr_src.Replace(",", ".");
                        }
                        else if (valuetypeid == 3)
                        {
                            //propvaluestr_dest_for_insert = "'" + Convert.ToDateTime(propvaluestr_src).ToShortDateString() + "'"; //? зачем конвертить

                            /*DateTime dateValue;
                            if (DateTime.TryParse(propvaluestr_src, out dateValue)) propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                            else propvaluestr_dest_for_insert = "NULL";*/

                            propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                        }
                    }

                    string queryStringlsprop2 = "INSERT INTO iesbk2.dbo." + propvaluetabledest + "(propvalue, lspropidglobal, period, IESBKlsid) " +
                        "VALUES (" +
                        propvaluestr_dest_for_insert + "," +
                        lspropidglobal.ToString() + "," +                        
                        "'" + period.ToShortDateString() + "'" + "," +
                        IESBKlsid.ToString() + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlsprop2);

                    //splashScreenManager1.SetWaitFormDescription("Обработка (" + (i + 1).ToString() + " - " + (j + 1).ToString() + " из " + tablelsprop.Rows.Count.ToString() + ")");
                    splashScreenManager1.SetWaitFormDescription(peryear + "," + permonth + " (" + (i + 1).ToString() +  " из " + tableTOTALls.Rows.Count.ToString() + ")");
                }
                                
                tablelsprop.Dispose();

            } // for (int i = 0; i < tableTOTALls.Rows.Count; i++)

            tableTOTALls.Dispose();

            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Значения свойств"

        //----------------------------------------------

        // перенос данных "Значения свойств2"
        // - грузим полную выборку свойств - вывалилась с ошибкой переполнения
        // пробуем по периодам - год, месяц - долго (лучше, чем полная задница)
        // пробуем по отборам Код л/с / Код свойства - полная задница!!!
        // оставим по коду свойства
        private void barButtonItem45_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            //string peryear = "2015"; // ручной ввод года
            string peryear = textBox1.Text; // ручной ввод года

            int startlspropidtmpl = Convert.ToInt32(textBox2.Text);

            // считываем значение lastpropvalueid
            string queryString = "SELECT lastpropvalueid FROM [iesbk2].[dbo].[tblIESBKsystem]";
            DataTable table_system = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(table_system, dbconnectionStringIESBK, queryString);

            Int64 propvalueid = Convert.ToInt64(table_system.Rows[0]["lastpropvalueid"].ToString());

            table_system.Dispose();
            //-----------------------------------------------

            for (int lspropidtmpl = startlspropidtmpl; lspropidtmpl <= 54; lspropidtmpl++)
            {
                //--------------------------------------
                /*string permonth = periodmonth.ToString();
                if (periodmonth < 10) permonth = "0" + permonth;*/
                                    
                string lspropidtmplstr = lspropidtmpl.ToString();
                //--------------------------------------

                // берем полную выборку (по lspropertieid)
                queryString =
                        "SELECT propvalue, periodyear, periodmonth, codeIESBK, lstypeid, lspropertieid, templateid, otdelenieid " +
                        "FROM [iesbk].[dbo].[tblIESBKlspropvalue] "
                        //+ "WHERE periodyear = '" + peryear + "' AND periodmonth = '" + permonth + "'";
                        + "WHERE lspropertieid = '" + lspropidtmplstr + "'" + " AND periodyear = '" + peryear + "'";
                DataTable tablelsprop = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsprop, dbconnectionStringIESBK, queryString);

                // получаем тип свойства
                int templateid = 1; // 3.3.2
                                    //int lspropertieid = Convert.ToInt32(tablelsprop.Rows[j]["lspropertieid"].ToString());

                queryString =
                    "SELECT lspropidglobal, valuetypeid " +
                    "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                    "WHERE templateid = " + templateid.ToString() +
                    " AND lspropidtmpl = " + lspropidtmpl;
                DataTable tablelsnew = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                int valuetypeid = Convert.ToInt32(tablelsnew.Rows[0]["valuetypeid"].ToString());
                int lspropidglobal = Convert.ToInt32(tablelsnew.Rows[0]["lspropidglobal"].ToString());
                tablelsnew.Dispose();
                //------------------------------

                for (int j = 0; j < tablelsprop.Rows.Count; j++)
                {
                    string codeIESBK = tablelsprop.Rows[j]["codeIESBK"].ToString();
                    string otdelenieid_old = tablelsprop.Rows[j]["otdelenieid"].ToString();

                    int otdelenieid_new = -1;
                    if (otdelenieid_old == "АО") otdelenieid_new = 10;
                    else if (otdelenieid_old == "ВО") otdelenieid_new = 11;
                    else if (otdelenieid_old == "КО") otdelenieid_new = 12;
                    else if (otdelenieid_old == "МЧО") otdelenieid_new = 13;
                    else if (otdelenieid_old == "СлО") otdelenieid_new = 14;
                    else if (otdelenieid_old == "СОЗ") otdelenieid_new = 15;
                    else if (otdelenieid_old == "СОС") otdelenieid_new = 16;
                    else if (otdelenieid_old == "ТОН") otdelenieid_new = 17;
                    else if (otdelenieid_old == "ТОТ") otdelenieid_new = 18;
                    else if (otdelenieid_old == "ТшО") otdelenieid_new = 19;
                    else if (otdelenieid_old == "УКО") otdelenieid_new = 20;
                    else if (otdelenieid_old == "УсО") otdelenieid_new = 21;
                    else if (otdelenieid_old == "ЦО") otdelenieid_new = 22;
                    else if (otdelenieid_old == "ЧО") otdelenieid_new = 23;

                    // получаем id "нового" лицевого счета
                    queryString =
                        "SELECT codeIESBK, lstypeid, otdelenieid, dateloaded, IESBKlsid " +
                        "FROM [iesbk2].[dbo].[tblIESBKls] " +
                        "WHERE codeIESBK = '" + codeIESBK + "' AND otdelenieid = '" + otdelenieid_new.ToString() + "'";
                    tablelsnew = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                    int IESBKlsid = Convert.ToInt32(tablelsnew.Rows[0]["IESBKlsid"].ToString());
                    tablelsnew.Dispose();
                    //-----------------------------------------------------------               

                    /*for (int j = 0; j < tablelsprop.Rows.Count; j++)
                    {*/
                    DateTime period = Convert.ToDateTime("01." + tablelsprop.Rows[j]["periodmonth"].ToString() + "." + tablelsprop.Rows[j]["periodyear"].ToString());
                    string propvaluestr_src = tablelsprop.Rows[j]["propvalue"].ToString();

                    /*// получаем тип свойства
                    int templateid = 1; // 3.3.2
                    int lspropertieid = Convert.ToInt32(tablelsprop.Rows[j]["lspropertieid"].ToString());

                    queryString =
                        "SELECT valuetypeid " +
                        "FROM [iesbk2].[dbo].[tblIESBKlsprop] " +
                        "WHERE templateid = " + templateid.ToString() +
                        " AND lspropertieid = " + lspropertieid;
                    tablelsnew = new DataTable();
                    SelectDataFromSQL(tablelsnew, dbconnectionStringIESBK2, queryString);
                    int valuetypeid = Convert.ToInt32(tablelsnew.Rows[0]["valuetypeid"].ToString());
                    tablelsnew.Dispose();
                    //------------------------------*/

                    // если в значении свойства присутствует символ ";" (замена ПУ), то 
                    // загружаем в текстовый куб
                    if (propvaluestr_src.Contains(";")) valuetypeid = 1;

                    // очистка старых косяков
                    if (propvaluestr_src.Contains("REF!")) propvaluestr_src = null;
                    //-----------------------------------------------------------------
                    
                    string propvaluestr_dest_for_insert = "NULL";
                    if (!String.IsNullOrWhiteSpace(propvaluestr_src))
                    {
                        if (valuetypeid == 1) propvaluestr_dest_for_insert = "'" + propvaluestr_src + "'";
                        else 
                        if (valuetypeid == 2) // число
                        {
                            propvaluestr_dest_for_insert = Convert.ToDouble(propvaluestr_src).ToString().Replace(",", ".");

                            /*double newDouble;
                            try
                            {
                                newDouble = Convert.ToDouble(propvaluestr_src);
                                propvaluestr_dest_for_insert = newDouble.ToString().Replace(",", ".");
                            }                            
                            catch (System.FormatException) // если нарушен формат, то считаем текстом
                            {
                                valuetypeid = 1;
                                propvaluestr_dest_for_insert = propvaluestr_src;
                            }*/
                        }
                        else 
                        if (valuetypeid == 3) // дата
                        {
                            propvaluestr_dest_for_insert = "'" + Convert.ToDateTime(propvaluestr_src).ToShortDateString() + "'";

                            /*DateTime newDate;
                            try
                            {
                                newDate = Convert.ToDateTime(propvaluestr_src);
                                propvaluestr_dest_for_insert = "'" + newDate.ToShortDateString() + "'";
                            }
                            catch (System.FormatException) // если нарушен формат, то считаем текстом
                            {
                                valuetypeid = 1;
                                propvaluestr_dest_for_insert = propvaluestr_src;
                            }*/
                        }
                    }
                    
                    // определяем таблицу назначения
                    string propvaluetabledest = null;
                    if (valuetypeid == 1)
                    {
                        propvaluetabledest = "tblIESBKlspropvaluestr"; // текстовый
                    }
                    else if (valuetypeid == 2)
                    {
                        propvaluetabledest = "tblIESBKlspropvaluenum"; // числовой
                    }
                    else if (valuetypeid == 3)
                    {
                        propvaluetabledest = "tblIESBKlspropvaluedate"; // дата время
                    }

                    // вставляем запись в куб "ссылок" на значения свойств
                    propvalueid++; // нумерация с 1
                    string queryStringlspropvalue = "INSERT INTO iesbk2.dbo.tblIESBKlspropvalue(propvalueid, lspropidglobal, period, IESBKlsid, valuetypeid) " +
                        "VALUES (" +
                        propvalueid.ToString() + "," +
                        lspropidglobal.ToString() + "," +
                        "'" + period.ToShortDateString() + "'" + "," +
                        IESBKlsid.ToString() + "," +
                        valuetypeid.ToString() + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlspropvalue);

                    // вставляем значения свойств по ссылке из куба
                    string queryStringlspropvalue2 = "INSERT INTO iesbk2.dbo." + propvaluetabledest + "(propvalueid, propvalue) " +
                        "VALUES (" +
                        propvalueid.ToString() + "," +
                        propvaluestr_dest_for_insert + ")";
                    MC_SQLDataProvider.InsertSQLQuery(dbconnectionStringIESBK2, queryStringlspropvalue2);

                    splashScreenManager1.SetWaitFormDescription("Обработка (" + lspropidtmplstr + " - " + (j + 1).ToString() + " из " + tablelsprop.Rows.Count.ToString() + ")");
                    //} // for (int j = 0; j < tablelsprop.Rows.Count; j++)

                    //tablelsprop.Dispose();

                } // for (int i = 0; i < tableTOTALls.Rows.Count; i++)

                tablelsprop.Dispose();

            } // for (int periodmonth = 1; periodmonth <= 12; periodmonth++)                

            // записываем последний lastpropvalueid в таблицу [tblIESBKsystem]            
            string queryString2 = "UPDATE iesbk2.dbo.tblIESBKsystem " +
                            "SET lastlsid = " + propvalueid.ToString();
            MC_SQLDataProvider.UpdateSQLQuery(dbconnectionStringIESBK2, queryString2);

            splashScreenManager1.CloseWaitForm();

        } // перенос данных "Значения свойств"

        // отчет-"шахматка" по наличию л/с и полезного отпуска - 2
        private void barButtonItem46_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            SqlConnection SQLconnection = new SqlConnection(dbconnectionStringIESBK2);
            SQLconnection.Open();

            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.MdiParent = this;
            form1.Text = "Шахматка с января 2016 по сентябрь 2016";
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            Worksheet worksheet = workbook.Worksheets[0];

            workbook.History.IsEnabled = false;
            form1.spreadsheetControl1.BeginUpdate();

            // константы
            int MAX_PERIOD_MONTH = 9;
            int columns_in_period_auto = 20; // колонок в периоде для автоматического вывода
            int columns_in_period_manual = 4; // колонок в периоде для ручного вывода
            int columns_in_period = columns_in_period_auto + columns_in_period_manual; // общее кол-во колонок в периоде
            int FIRST_COLUMNS = 8;
            int END_COLUMNS = 1 + 3 + 8 + MAX_PERIOD_MONTH;

            //int MAXPROPMAS = 9;
            int MAXCOLinWRKSH = FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + END_COLUMNS;

            for (int col = 0; col < MAXCOLinWRKSH; col++)
            {
                worksheet.Columns[col].Font.Name = "Arial";
                worksheet.Columns[col].Font.Size = 8;
            }

            DateTime dt_IESBK_MIN = Convert.ToDateTime("01.01.2015"); // левая граница имеющихся данных в OLAP-кубе

            // загружаем параметры свойств в глобальную таблицу (тестируем производительность)
            string queryString22 =
                    "SELECT lspropidglobal, valuetypeid " +
                    "FROM [iesbk2].[dbo].[tblIESBKlsprop]";
            tablelsPropGLOBAL = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tablelsPropGLOBAL, queryString22);            
            //-------------------------------

            /*// загружаем отделения ИЭСБК
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();
            tblIESBKotdelenieTableAdapter.Fill(DataSetIESBKLoad.tblIESBKotdelenie);*/
            string queryString = "SELECT otdelenieid, captionotd " +
                                 "FROM [iesbk2].[dbo].[tblIESBKotdelenie]";                                
            DataTable tableIESBKotdelenie = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableIESBKotdelenie, queryString);

            //-----------------------------------------

            // продумать выборку!!! пробуем JOIN

            // св-во 36 - "Расход ОДН по нормативу" (убрал)
            /*string queryString = "SELECT DISTINCT codeIESBK,otdelenieid " +
                                 "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +                                
                                "WHERE periodyear = '2016' AND (periodmonth = '01' OR periodmonth = '07') AND lspropertieid='36' AND propvalue IS NULL";*/

            /*queryString = "SELECT DISTINCT IESBKlsid " +
                          "FROM [iesbk2].[dbo].[tblIESBKlspropvaluestr] " +
                          "WHERE lspropidglobal = 1003 AND DATEPART(year, period) = 2016";*/

            queryString = "SELECT tblIESBKPVstr.IESBKlsid, tblIESBKls.codeIESBK, tblIESBKotd.captionotd" +
                          " FROM" + 
                          " (SELECT DISTINCT IESBKlsid, lspropidglobal" +
                          " FROM iesbk2.dbo.tblIESBKlspropvaluestr" +
                          " WHERE lspropidglobal = 1003 AND DATEPART(year, period) = 2016) tblIESBKPVstr" + 
                          " LEFT JOIN iesbk2.dbo.tblIESBKls tblIESBKls ON tblIESBKPVstr.IESBKlsid = tblIESBKls.IESBKlsid" +
                          " LEFT JOIN iesbk2.dbo.tblIESBKotdelenie tblIESBKotd ON tblIESBKls.otdelenieid = tblIESBKotd.otdelenieid";

            //+ " AND (otdelenieid = 'СОС' OR otdelenieid = 'СОЗ')";// AND lspropertieid='36' AND propvalue IS NULL";
            DataTable tableTOTALls10 = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableTOTALls10, dbconnectionStringIESBK2, queryString);
            //-----------------------------

            // выводим заголовки столбцов -----------------------------------------------------

            // статические
            worksheet[0, 0].SetValue("№ п/п");
            worksheet[0, 1].SetValue("Отделение ИЭСБК");
            worksheet[0, 2].SetValue("Код л/с ИЭСБК");
            worksheet[0, 3].SetValue("ФИО");
            worksheet[0, 4].SetValue("Населенный пункт");
            worksheet[0, 5].SetValue("Улица");
            worksheet[0, 6].SetValue("Дом");
            worksheet[0, 7].SetValue("Номер квартиры");
            //worksheet[0, 8].SetValue("Состояние ЛС (на 2016 09)");

            // периодические
            // id "периодических" свойств
            //int[] propidmas = new int[] { 6, 7, 50, 24, 25, 26, 27, 28, 29, 30, 53, 31, 54, 32, 33, 34 };
            int[] propidmas = new int[] { 1051, 1024, 1025, 1006, 1007, 1050, 1026, 1027, 1028, 1055, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034, 1035, 1036 };
            int idprop_PO_in_propidmas = 7; // индекс идентификатора поля ПО от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_lastPOK_in_propidmas = 2; // индекс идентификатора поля ПослПоказаниеПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int idprop_nomerPU_in_propidmas = 4; // индекс идентификатора поля ЗаводскойНомерПУ от ИЭСБК в массиве периодических свойств (нумерация с 0)
            int[] propidmas_doublevalue = new int[] { 1027, 1028, 1055, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034, 1035, 1036 }; // id числовых полей

            string queryStringlsprop = "SELECT lspropidglobal, captionlsprop " +
                                        "FROM [iesbk2].[dbo].[tblIESBKlsprop]";
            DataTable tableLSprop = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLSprop, queryStringlsprop);

            Color Color_IS_PO_NORMATIV_NOT_PU = Color.Orange; // цвет "норматив - безприборник"
            Color Color_IS_PO_NORMATIV_YES_PU = Color.Red; // цвет "норматив - приборник"
            Color Color_IS_PO_SREDNEMES_YES_PU = Color.Blue; // цвет "среднемесячное - приборник"
            Color Color_IS_PO_RASHOD_YES_PU = Color.Green; // цвет "расход по прибору"

            for (int period_i = 1; period_i < MAX_PERIOD_MONTH + 1; period_i++)
            {
                string periodyear = "2016";
                string periodmonth = (period_i < 10) ? "0" + period_i.ToString() : period_i.ToString();

                for (int k = 0; k < columns_in_period_auto; k++)
                {
                    DataRow[] lsproprows = tableLSprop.Select("lspropidglobal = " + propidmas[k].ToString());
                    string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["captionlsprop"].ToString() : null;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(periodyear + " " + periodmonth + ", " + propvaluestr);

                    Color cellColor = Color.Black;
                    // "красим" названия столбцов
                    if (propidmas[k] == 1028) cellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (propidmas[k] == 1029) cellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (propidmas[k] == 1030 || propidmas[k] == 53) cellColor = Color_IS_PO_SREDNEMES_YES_PU;
                    else if (propidmas[k] == 1031 || propidmas[k] == 54) cellColor = Color_IS_PO_NORMATIV_YES_PU;
                    worksheet[0, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].Font.Color = cellColor;
                }

                tableLSprop.Dispose();

                //------------------------------------------------------------------------------

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(periodyear + " " + periodmonth + ", " + "Расчетный полезный отпуск");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(periodyear + " " + periodmonth + ", " + "Расход по показаниям ПУ");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(periodyear + " " + periodmonth + ", " + "Начисленный полезный отпуск от ИЭСБК");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(periodyear + " " + periodmonth + ", " + "Отклонение (недополученный ПО)");
                worksheet[0, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;
            }

            //------------------------------------------------------------------------------
            string periodmonthnext = (MAX_PERIOD_MONTH + 1 < 10) ? "0" + (MAX_PERIOD_MONTH + 1).ToString() : (MAX_PERIOD_MONTH + 1).ToString();
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue("2016" + " " + periodmonthnext + ", " + "среднемесячное (прогноз ПО)");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue("ИТОГО Расход по показаниям ПУ с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].Font.Color = Color.Green;

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue("Начальное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue("Начальное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue("Начальное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue("Начальное показание, номер ПУ");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue("Конечное показание, дата");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue("Конечное показание");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue("Конечное показание, вид");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue("Конечное показание, номер ПУ");

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue("ИТОГО Полезный отпуск от ИЭСБК с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].Font.Color = Color.Blue;

            for (int period_i = 1; period_i <= MAX_PERIOD_MONTH; period_i++)
            {
                worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10 + period_i].
                    SetValue(worksheet[0, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString());
            }

            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].SetValue("ИТОГО Отклонение (недополученный ПО) с начала года");
            worksheet[0, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].Font.Color = Color.Red;
            //----------------------------------------------------------

            // главный цикл
            //for (int i = 0; i < tableTOTALls10.Rows.Count; i++)
            for (int i = 0; i < 400; i++)
            {
                //DataRow[] lsproprows;

                // получаем внутренний идентификатор л/с, отделение и код ИЭСБК
                /*queryString = "SELECT IESBKlsid, codeIESBK, otdelenieid " +
                              "FROM [iesbk2].[dbo].[tblIESBKls] " +
                              "WHERE IESBKlsid = " + tableTOTALls10.Rows[i]["IESBKlsid"].ToString();
                DataTable tableIESBKls = new DataTable();
                MyFUNC_SelectDataFromSQLwoutConnection(tableIESBKls, SQLconnection, queryString);
                
                int IESBKlsid = Convert.ToInt32(tableTOTALls10.Rows[i]["IESBKlsid"]);
                string codeIESBK = tableIESBKls.Rows[0]["codeIESBK"].ToString();
                //string codeIESBK = "ККОО00019257";
                string otdelenieid = tableIESBKls.Rows[0]["otdelenieid"].ToString();
                
                tableIESBKls.Dispose();

                string otdeleniecapt = tableIESBKotdelenie.Select("otdelenieid = " + otdelenieid)[0]["captionotd"].ToString();*/
                //string otdeleniecapt = "АО";
                                
                int IESBKlsid = Convert.ToInt32(tableTOTALls10.Rows[i]["IESBKlsid"]);
                string codeIESBK = tableTOTALls10.Rows[i]["codeIESBK"].ToString();                
                string otdeleniecapt = tableTOTALls10.Rows[i]["captionotd"].ToString();
                //--------------------------------

                worksheet[i + 1, 0].SetValue((i + 1).ToString());
                worksheet[i + 1, 1].SetValue(otdeleniecapt);
                worksheet[i + 1, 2].SetValue(codeIESBK);

                /*// состояние ЛС за июнь 2016  ---------------------------------------
                string periodyear = "2016";
                string periodmonth = "06";                    
                worksheet[i + 1, 3].SetValue(MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 51, SQLconnection));*/

                // основные параметры ЛС за начальный период (01 2016)  ---------------------------------------
                string periodyear = "2016";
                string periodmonth = "01";

                // сделано для увеличения производительности
                /*// "статические" свойства лицевого счета                
                queryStringlsprop = "SELECT codeIESBK, lspropertieid, propvalue " +
                                     "FROM [iesbk].[dbo].[tblIESBKlspropvalue] " +
                                     "WHERE codeIESBK='" + codeIESBK + "' AND periodyear = '" + periodyear + "' AND periodmonth = '" + periodmonth + "'";
                tableLSprop = new DataTable();
                MyFUNC_SelectDataFromSQLwoutConnection(tableLSprop, SQLconnection, queryStringlsprop);*/
                                                
                // фио
                /*string lspropvalue = "";
                                
                lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 5, SQLconnection);
                worksheet[i + 1, 3].SetValue(lspropvalue);

                /*DataRow[] lsproprows = tableLSprop.Select("lspropertieid='5'");
                if (lsproprows.Length > 0) worksheet[i + 1, 3].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // населенный пункт                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 12, SQLconnection);
                worksheet[i + 1, 4].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='12'");
                if (lsproprows.Length > 0) worksheet[i + 1, 4].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // улица                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 13, SQLconnection);
                worksheet[i + 1, 5].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='13'");
                if (lsproprows.Length > 0) worksheet[i + 1, 5].SetValue(lsproprows[0]["propvalue"].ToString());*/

                // дом                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 14, SQLconnection);
                worksheet[i + 1, 6].SetValue(lspropvalue);
                /*lsproprows = tableLSprop.Select("lspropertieid='14'");
                if (lsproprows.Length > 0) worksheet[i + 1, 6].SetValue(lsproprows[0]["propvalue"].ToString());                */

                // номер квартиры                
                /*lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP(codeIESBK, periodyear, periodmonth, 15, SQLconnection);
                worksheet[i + 1, 7].SetValue(lspropvalue);                
                /*lsproprows = tableLSprop.Select("lspropertieid='15'");
                if (lsproprows.Length > 0) worksheet[i + 1, 7].SetValue(lsproprows[0]["propvalue"].ToString());                */

                //tableLSprop.Dispose();
                //---------------------------------------

                // переменные для подсчета ПО по разнице показаний + сам ПО
                double POKstart = -1;
                double POKend = -1;

                DateTime POKstart_date = Convert.ToDateTime("01.01.1900");
                DateTime POKend_date = Convert.ToDateTime("01.01.1900");
                string POKstart_kind = null;
                string POKend_kind = null;
                int periodPOKstart = -1; // нумерация с 1 (январь)
                int periodPOKend = -1; // нумерация с 1 (январь)

                string nomerPUstart = null;
                string nomerPUend = null;
                //------------------------------------------------

                // "бежим" по периодическим свойствам лицевого счета
                int START_PERIOD_MONTH = 1;
                for (int period_i = START_PERIOD_MONTH; period_i < START_PERIOD_MONTH + MAX_PERIOD_MONTH; period_i++)
                {
                    periodyear = "2016";
                    periodmonth = (period_i < 10) ? "0" + period_i.ToString() : period_i.ToString();

                    DateTime periodAsDateTime = Convert.ToDateTime("01." + periodmonth + "." + periodyear);

                    /*queryStringlsprop = "SELECT lspropidglobal, propvalue " +
                                     "FROM [iesbk2].[dbo].[tblIESBKlspropvaluestr] " +
                                     "WHERE period ='" + periodAsDateTime + "' AND IESBKlsid = " + IESBKlsid.ToString();
                    tableLSprop = new DataTable();
                    MyFUNC_SelectDataFromSQLwoutConnection(tableLSprop, SQLconnection, queryStringlsprop);                    */
                    //-----------------------------------
                    // "статические" свойства лицевого счета                                
                                        
                    string lspropvalue = "";

                    // фио                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1005, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 3].SetValue(lspropvalue);

                    // населенный пункт                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1012, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 4].SetValue(lspropvalue);

                    // улица                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1013, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 5].SetValue(lspropvalue);

                    // дом                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1014, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 6].SetValue(lspropvalue);

                    // номер квартиры                                    
                    lspropvalue = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1015, SQLconnection);
                    if (!String.IsNullOrWhiteSpace(lspropvalue)) worksheet[i + 1, 7].SetValue(lspropvalue);
                    //-----------------------------------
                                        
                    /*// состояние ЛС (по последнему расчетному периоду)
                    if (period_i == MAX_PERIOD_MONTH)
                    {                        
                        worksheet[i + 1, 8].SetValue(MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1051, SQLconnection));
                    }*/

                    // флаги для раскраски ячейки полезного отпуска                    
                    bool IS_PO_NORMATIV_NOT_PU = false; // флаг "норматив - безприборник"
                    bool IS_PO_NORMATIV_YES_PU = false; // флаг "норматив - приборник"
                    bool IS_PO_SREDNEMES_YES_PU = false; // флаг "среднемесячное - приборник"
                    bool IS_PO_RASHOD_YES_PU = false; // флаг "расход по прибору"

                    // выводим "периодические" поля - сделал в цикле                    
                    for (int k = 0; k < columns_in_period_auto; k++)
                    {
                        /*string strTEST = "12345678";
                        string propvaluestr = strTEST;*/

                        /*lsproprows = tableLSprop.Select("lspropertieid = '" + propidmas[k].ToString() + "'");
                        string propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                        string propvaluestr = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, propidmas[k], SQLconnection);

                        double? propvalue = null;

                        if (System.Array.IndexOf(propidmas_doublevalue, propidmas[k]) >= 0)
                        {
                            if (!String.IsNullOrWhiteSpace(propvaluestr)) propvalue = Convert.ToDouble(propvaluestr);
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvalue);
                        }
                        else
                        {
                            worksheet[i + 1, FIRST_COLUMNS + k + (period_i - 1) * columns_in_period].SetValue(propvaluestr);
                        }

                        // формируем флаги для раскраски ячейки полезного отпуска
                        // помним о массиве периодических свойств
                        //int[] propidmas = new int[] { 1006, 1007, 1050, 1024, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };                        
                        if (propvalue != null && propvalue != 0)
                        {
                            if (propidmas[k] == 1028) IS_PO_NORMATIV_NOT_PU = true;
                            else if (propidmas[k] == 1029) IS_PO_RASHOD_YES_PU = true;
                            else if (propidmas[k] == 1030) IS_PO_SREDNEMES_YES_PU = true;
                            else if (propidmas[k] == 1031) IS_PO_NORMATIV_YES_PU = true;
                        };
                        //-------------------------------------------------------
                    } // выводим "периодические" поля - сделал в цикле

                    // раскрашиваем колонку ПолОтп в зависимости от "слагаемых"
                    // помним о массиве периодических свойств
                    //int[] propidmas = new int[] { 1006, 1007, 1050, 1024, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };

                    Color PolOtpCellColor = Color.Black;
                    if (IS_PO_RASHOD_YES_PU) PolOtpCellColor = Color_IS_PO_RASHOD_YES_PU;
                    else if (IS_PO_NORMATIV_NOT_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_NOT_PU;
                    else if (IS_PO_NORMATIV_YES_PU) PolOtpCellColor = Color_IS_PO_NORMATIV_YES_PU;
                    else if (IS_PO_SREDNEMES_YES_PU) PolOtpCellColor = Color_IS_PO_SREDNEMES_YES_PU;
                    worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + idprop_PO_in_propidmas].Font.Color = PolOtpCellColor; // "красим" ПолезныйОтпуск - propid = 27

                    //-----------------------------------------------------------------
                    // формируем "периодические" колонки "Расход по показаниям" и "Отклонение (недополученный ПО)", если не было замены ПУ

                    if (period_i > START_PERIOD_MONTH) // пропускаем первый месяц, т.к. в нем не найдем "предыдущих показаний"
                    {
                        string propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        double POKend_period = -1;
                        //double? POIESBK_period = null;
                        double POIESBK_period = 0;
                        double POIESBKend_period = 0; // ПО последнего периода

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";"))
                        {
                            POKend_period = Convert.ToDouble(propvaluestr);

                            double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                            POIESBK_period += POcellvalue;

                            POIESBKend_period = POcellvalue;
                        }

                        string nomerPUend_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        /*lsproprows = tableLSprop.Select("lspropertieid='22'"); // предыдущее показание ПУ
                        propvaluestr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                        int period_i_step = 1;
                        double POKstart_period = -1;
                        string nomerPUstart_period = null;
                        propvaluestr = null;

                        do
                        {
                            propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_lastPOK_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();
                            nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.ToString();

                            // суммируем полезный отпуск, пропуская начальный интервал
                            if (String.IsNullOrWhiteSpace(propvaluestr))
                            {
                                double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1 - period_i_step) * columns_in_period].Value.NumericValue;
                                POIESBK_period += POcellvalue;
                            }

                            period_i_step += 1;

                        } while (String.IsNullOrWhiteSpace(propvaluestr) && period_i_step < period_i);

                        if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POKstart_period = Convert.ToDouble(propvaluestr);

                        //nomerPUstart_period = worksheet[i + 1, FIRST_COLUMNS + idprop_nomerPU_in_propidmas + (period_i - 2) * columns_in_period].Value.ToString();
                        //--------------------------------------------------------------

                        // если имеются оба показания и не было замены ПУ, то считаем значения
                        if (POKend_period != -1 && POKstart_period != -1 && String.Equals(nomerPUstart_period, nomerPUend_period))
                        //if (POKend_period != -1 && POKstart_period != -1 && POKstart_period <= POKend_period)                         
                        {
                            /*propvaluestr = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.ToString();
                            double? POIESBK_period = null;
                            if (!String.IsNullOrWhiteSpace(propvaluestr) && !propvaluestr.Contains(";")) POIESBK_period = Convert.ToDouble(propvaluestr);*/

                            double POIESBKPU_period = POKend_period - POKstart_period;
                            double POIESBKDelta_period = POIESBKPU_period - POIESBK_period;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].SetValue(POIESBKend_period + POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 0].Font.Color = Color.DimGray;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].SetValue(POIESBKPU_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 1].Font.Color = Color.Green;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].SetValue(POIESBK_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 2].Font.Color = Color.Blue;

                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].SetValue(POIESBKDelta_period);
                            worksheet[i + 1, FIRST_COLUMNS + (period_i - 1) * columns_in_period + columns_in_period_auto + 3].Font.Color = Color.Red;
                        }

                    } // if (period_i > START_PERIOD_MONTH) 
                    //-----------------------------------------------------------------

                    // ищем стартовое и конечное показание для расчета ИТОГОВОГО ПО по показаниям (за все периоды)                  
                    /*lsproprows = tableLSprop.Select("lspropertieid='25'");
                    string pokstr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                    string pokstr = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1025, SQLconnection);

                    if (!String.IsNullOrWhiteSpace(pokstr) && !pokstr.Contains(";"))
                    {
                        if (POKstart == -1)
                        {
                            POKstart = Convert.ToDouble(pokstr);
                            periodPOKstart = period_i;

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1007");
                            nomerPUstart = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            nomerPUstart = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1007, SQLconnection);

                            /*lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKstart_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());*/
                            POKstart_date = Convert.ToDateTime(MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1024, SQLconnection));

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1026");
                            POKstart_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            POKstart_kind = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1026, SQLconnection);
                        }
                        else
                        {
                            // поменять местами нижнее условие и присовение значение конечному показанию!!!!!
                            POKend = Convert.ToDouble(pokstr);
                            periodPOKend = period_i;

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1007");
                            nomerPUend = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            nomerPUend = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1007, SQLconnection);

                            /*lsproprows = tableLSprop.Select("lspropertieid='24'");
                            POKend_date = Convert.ToDateTime(lsproprows[0]["propvalue"].ToString());*/
                            POKend_date = Convert.ToDateTime(MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1024, SQLconnection));

                            /*lsproprows = tableLSprop.Select("lspropidglobal = 1026");
                            POKend_kind = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;*/
                            POKend_kind = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, periodyear, periodmonth, 1026, SQLconnection);

                            if (!String.Equals(nomerPUstart, nomerPUend)) // если номера ПУ не равны, то последнее делаем начальным
                            {
                                nomerPUstart = nomerPUend;
                                POKstart = POKend;
                                POKstart_date = POKend_date;
                                POKstart_kind = POKend_kind;

                                POKend = -1;

                                periodPOKstart = periodPOKend;
                                periodPOKend = -1;
                            };
                        }
                    }

                    //-----------------------------------------------------------------

                    tableLSprop.Dispose();

                } // for (int period_i = 1; period_i < 7; period_i++)

                if (POKstart != -1 && POKend != -1) // если имеются оба показания ПУ (для расчета ПО по показаниям ПУ)
                {
                    /*// суммируем полезный отпуск от ИЭСБК со следующего расчетного периода, которому предшествовало показание ПУ
                    lsproprows = tableLSprop.Select("lspropertieid='27'");
                    string postr = (lsproprows.Length > 0) ? lsproprows[0]["propvalue"].ToString() : null;
                    if (POKstart != -1 && POKend == -1 && !String.IsNullOrWhiteSpace(postr)) POIESBKTotal += Convert.ToDouble(postr);*/

                    // суммируем полезный отпуск от ИЭСБК по ранее заполненным колонкам со следующего расчетного периода, которому предшествовало показание ПУ
                    double POIESBKTotal = 0;
                    for (int period_i = periodPOKstart + 1; period_i <= periodPOKend; period_i++)
                    {
                        double POcellvalue = worksheet[i + 1, FIRST_COLUMNS + idprop_PO_in_propidmas + (period_i - 1) * columns_in_period].Value.NumericValue;
                        POIESBKTotal += POcellvalue;

                        // выводим расшифровку формирования ПО
                        worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10 + period_i].SetValue(POcellvalue);
                    }

                    double POIESBKPU = POKend - POKstart;
                    double POIESBKDelta = POIESBKPU - POIESBKTotal;

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].SetValue(POIESBKPU);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1].Font.Color = Color.Green;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 2].SetValue(POKstart_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 3].SetValue(POKstart);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 4].SetValue(POKstart_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 5].SetValue(nomerPUstart);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 6].SetValue(POKend_date);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 7].SetValue(POKend);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8].SetValue(POKend_kind);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 9].SetValue(nomerPUend);

                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].SetValue(POIESBKTotal);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 10].Font.Color = Color.Blue;
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].SetValue(POIESBKDelta);
                    worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 11 + MAX_PERIOD_MONTH].Font.Color = Color.Red;

                    // beg расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)

                    string periodyeartek = "2016";
                    string periodmonthtek = (MAX_PERIOD_MONTH < 10) ? "0" + MAX_PERIOD_MONTH.ToString() : MAX_PERIOD_MONTH.ToString();
                    //string codels = codeIESBK;

                    // ищем ближайшее "правое" показание                
                    string value_right = null;
                    string dtvalue_right = null;

                    DateTime dt_right = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);
                    string year_right = null;
                    string month_right = null;

                    dt_right = dt_right.AddMonths(+1); // учитываем текущий месяц, т.е. +1-1 = 0

                    while (String.IsNullOrWhiteSpace(value_right) && dt_right >= dt_IESBK_MIN)
                    {
                        dt_right = dt_right.AddMonths(-1);
                        year_right = dt_right.Year.ToString();
                        month_right = null;
                        if (dt_right.Month < 10) month_right = "0" + dt_right.Month.ToString();
                        else month_right = dt_right.Month.ToString();

                        value_right = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_right, month_right, 1025, SQLconnection); // свойство "Текущее показание ПУ"                                        
                    };
                    //----------------------------------

                    // ищем "левое" показание, при условии, что нашли "правое" ------------------
                    string value_left = null;
                    string dtvalue_left = null;
                    string year_left = null;
                    string month_left = null;

                    if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    {
                        // получаем дату "правого" показания
                        dtvalue_right = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_right, month_right, 1024, SQLconnection); // свойство "Дата последнего показания ПУ"

                        /*// выводим информацию о "правом" показании
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 10].SetValue(dtvalue_right); // дата

                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 11].SetValue(value_right); // показание

                        string rightpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_right, month_right, 26, SQLconnection);
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].Font.Color = Color.Green;
                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 12].SetValue(rightpok_type); // вид*/

                        DateTime dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-5); // отматываем -5-1 = -6 мес. (было 6) = 180 дней от "правого" показания

                        while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN)
                        {
                            dt_left = dt_left.AddMonths(-1);
                            year_left = dt_left.Year.ToString();
                            month_left = null;
                            if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                            else month_left = dt_left.Month.ToString();

                            value_left = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_left, month_left, 1025, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                        };

                        // если нет данных да период не менее 6 мес., то ищем за в периоде [6 мес.;3 мес.]
                        if (String.IsNullOrWhiteSpace(value_left))
                        {
                            dt_left = Convert.ToDateTime(dtvalue_right).AddMonths(-7); // отматываем 7 мес., т.к. в теле цикла сразу +1, т.е. -7+1 = -6

                            DateTime dt_IESBK_left_MAX = Convert.ToDateTime(dtvalue_right).AddMonths(-3);

                            while (String.IsNullOrWhiteSpace(value_left) && dt_left >= dt_IESBK_MIN && dt_left < dt_IESBK_left_MAX)
                            {
                                dt_left = dt_left.AddMonths(+1);
                                year_left = dt_left.Year.ToString();
                                month_left = null;
                                if (dt_left.Month < 10) month_left = "0" + dt_left.Month.ToString();
                                else month_left = dt_left.Month.ToString();

                                value_left = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_left, month_left, 1025, SQLconnection); // свойство "Текущее показание ПУ"                                                                    
                            };
                        } // if (String.IsNullOrWhiteSpace(value_left)) // если нет данных за период не менее 6 мес.

                        // получаем дату "левого" показания и выводим информацию о нем
                        if (!String.IsNullOrWhiteSpace(value_left))
                        {
                            dtvalue_left = MyFUNC_GetPropValueFromIESBKOLAP2(IESBKlsid, year_left, month_left, 1024, SQLconnection); // свойство "Дата последнего показания ПУ"

                            /*// выводим информацию о "правом" показании
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 7].SetValue(dtvalue_left); // дата

                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 8].SetValue(value_left); // показание

                            string leftpok_type = MyFUNC_GetPropValueFromIESBKOLAP(codels, year_left, month_left, 26, SQLconnection);
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].Font.Color = Color.Green;
                            worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 9].SetValue(leftpok_type); // вид*/
                        }

                        // если даты "левого" и "правого" показаний не пустые, то формируем расчет среднемесячного
                        if (!String.IsNullOrWhiteSpace(dtvalue_left) && !dtvalue_left.Contains(";") && !String.IsNullOrWhiteSpace(dtvalue_right))
                        {
                            double pokleft = Convert.ToDouble(value_left);
                            double pokright = Convert.ToDouble(value_right);

                            // если не нарушен нарастающий итог
                            if (pokleft <= pokright)
                            {
                                System.TimeSpan deltaday = Convert.ToDateTime(dtvalue_right) - Convert.ToDateTime(dtvalue_left);
                                double deltapok = pokright - pokleft;

                                double srednesut_calc = deltapok / deltaday.Days;
                                double srmes_calc = Math.Round(srednesut_calc * DateTime.DaysInMonth(Convert.ToInt32(periodyeartek), Convert.ToInt32(periodmonthtek)));

                                // формируем отчет -----------------------------------------------

                                /*if (!String.IsNullOrWhiteSpace(srmes_iesbk_str)) // еслм СрМес ПО ИЭСБК присутствует
                                {
                                    DateTime dt_period = Convert.ToDateTime("01." + periodmonthtek + "." + periodyeartek);

                                    if (srmes_iesbk >= 0 && Convert.ToDateTime(dtvalue_right).CompareTo(dt_period) < 0) // не выводим наши расчеты, если СрМес ИЭСБК < 0 и правое показание принадлежит текущему периоду анализа
                                    {
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].Font.Color = Color.Green;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 13].SetValue(srmes_calc); // СрМес РАСЧ                                

                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].Font.Color = Color.Red;
                                        worksheet[strow + rd + 0, stcol + (period_i - 1) * columns_in_period + 16].SetValue(srmes_calc - srmes_iesbk); // Недополученный ПО                                
                                    }
                                }*/

                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].SetValue(srmes_calc);
                                worksheet[i + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 0].Font.Color = Color.BlueViolet;

                            } // if (pokleft <= pokright)
                              //-----------------------------------------------------------------
                        } // if (!String.IsNullOrWhiteSpace(dtvalue_left) && !String.IsNullOrWhiteSpace(dtvalue_right))

                    } // if (!String.IsNullOrWhiteSpace(value_right) && !value_right.Contains(";"))
                    // end расчет среднемесячного начисления в следующем расчетном периоде (прогноза ПО)
                }

                splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + tableTOTALls10.Rows.Count.ToString() + ")");

            } // for (int i = 0; i < tableTOTALls10.Rows.Count; i++)

            // форматируем строку-заголовок
            worksheet.Rows[0].Font.Bold = true;
            worksheet.Rows[0].Alignment.WrapText = true;
            worksheet.Rows[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            worksheet.Rows[0].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            worksheet.Rows[0].AutoFit();

            worksheet.Columns.AutoFit(0, MAXCOLinWRKSH);

            worksheet.Columns.Group(3, 7, true); // группируем по колонкам "ФИО" - "Номер квартиры"            

            //int[] propidmas = new int[] { 1006, 1007, 1050, 1024, 1025, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };
            //int[] propidmas = new int[] { 1051, 1024, 1025, 1006, 1007, 1050, 1026, 1027, 1028, 1029, 1030, 1053, 1031, 1054, 1032, 1033, 1034 };

            // группируем "периодические" значения - в частности расшифровку ПО
            for (int period_i = 0; period_i < MAX_PERIOD_MONTH; period_i++)
            {
                //worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period, FIRST_COLUMNS + period_i * columns_in_period + 2, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 2 + 1, FIRST_COLUMNS + period_i * columns_in_period + 4 + 1, true); // до "ПолОтп"
                worksheet.Columns.Group(FIRST_COLUMNS + period_i * columns_in_period + 7 + 1, FIRST_COLUMNS + period_i * columns_in_period + 7 + 10 + 1, true); // после "ПолОтп"
            }

            // группируем последние колонки анализа ПО по показаниям ПУ
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 1 + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1, true);

            // группируем колонки слагаемых ПО от ИЭСБК
            worksheet.Columns.Group(FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 2 + 1, FIRST_COLUMNS + MAX_PERIOD_MONTH * columns_in_period + 8 + 1 + MAX_PERIOD_MONTH + 1, true);

            worksheet.FreezeRows(0); // "фиксируем" верхнюю строку

            form1.spreadsheetControl1.EndUpdate();

            tableIESBKotdelenie.Dispose();
            tablelsPropGLOBAL.Dispose();

            SQLconnection.Close();
            splashScreenManager1.CloseWaitForm();
            form1.Show();
        } // отчет-"шахматка" по наличию л/с и полезного отпуска - 2

        
        private void barButtonItem47_ItemClick(object sender, ItemClickEventArgs e)
        {

        }

        // РТП-3 - Отчеты из РТП-3 - Реестр потребителей
        private void barButtonItemRTP3_ConsumersInfo_ItemClick(object sender, ItemClickEventArgs e)
        {
            OpenFileDialog openFileDialogRTP = new OpenFileDialog();
            openFileDialogRTP.Title = "Выберите файл базы данных РТП-3";
            openFileDialogRTP.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
            openFileDialogRTP.FileName = "";

            if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
            {
                splashScreenManager1.ShowWaitForm();

                splashScreenManager1.SetWaitFormDescription("подключение к базе данных...");

                // создаем соединение и загружаем данные
                FbConnection connection = new FbConnection();
                connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", openFileDialogRTP.FileName, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");
                                
                FbCommand command = new FbCommand();
                command.Connection = connection;

                // читаем запрос из файла (для теста, потом убрать!!!)
                command.CommandText = "";
                string line;
                StreamReader sr = new StreamReader(@"sql_RTP3.txt");
                while ((line = sr.ReadLine()) != null)
                {
                    command.CommandText = String.Concat(command.CommandText, line, Environment.NewLine);
                }
                //command.CommandText = "select \"Customer\" from \"LVConsumersInfo\"";
                /*command.CommandText = String.Concat(
                    SELECT
                    t_LVConsumersInfo."GUID", t_LVConsumersInfo."Contract", t_LVConsumersInfo."Customer", t_LVConsumersInfo."PointName", t_LVConsumersInfo."Address",
                    v_LVConsumersA_LVNodes_GUID, v_LVConsumersA_LVNodes_OwnerGUID, v_LVConsumersA_LVNodes_Name,
                    v_LVFiders_GUID, v_LVFiders_OwnerGUID, v_LVFiders_Name,
                    v_Nodes_GUID, v_Nodes_OwnerGUID, v_Nodes_Name, v_Nodes_onBalance, v_Nodes_Transf,
                    v_Transforms2_Ident, v_Transforms2_Unom, v_Transforms2_TypeTR, v_Transforms2_Snom,
                    v_Fiders_GUID, v_Fiders_OwnerGUID, v_Fiders_Name,
                    v_Sections_GUID, v_Sections_OwnerGUID, v_Sections_Name, v_Sections_Unom,
                    v_Centers_GUID, v_Centers_OwnerGUID, v_Centers_Name

                    FROM "LVConsumersInfo" AS t_LVConsumersInfo

                    LEFT JOIN

                    (SELECT
                    t_LVConsumersA_LVNodes."GUID" AS v_LVConsumersA_LVNodes_GUID, t_LVConsumersA_LVNodes."OwnerGUID" AS v_LVConsumersA_LVNodes_OwnerGUID, t_LVConsumersA_LVNodes."Name" AS v_LVConsumersA_LVNodes_Name,
                    v_LVFiders_GUID, v_LVFiders_OwnerGUID, (CASE WHEN v_LVFidersName IS NULL THEN 'подключен к ТП' ELSE v_LVFidersName END) AS v_LVFiders_Name
                    FROM
                    (SELECT "GUID", "OwnerGUID", "Name" FROM "LVConsumersA"
                    UNION
                    SELECT "GUID", "OwnerGUID", "Name" FROM "LVNodes") AS t_LVConsumersA_LVNodes
                    LEFT JOIN
                    (SELECT "GUID" AS v_LVFiders_GUID, "OwnerGUID" AS v_LVFiders_OwnerGUID, "Name" AS v_LVFidersName FROM "LVFiders") AS t_LVFiders
                    ON t_LVConsumersA_LVNodes."OwnerGUID" = t_LVFiders.v_LVFiders_GUID
                    )
                    AS t_LVConsumersA_LVNodes_Fiders

                    LEFT JOIN
                    (SELECT "GUID" AS v_Nodes_GUID, "OwnerGUID" AS v_Nodes_OwnerGUID, "Name" AS v_Nodes_Name,
                    (CASE
                    WHEN "Nodes"."onBalance" = 0 THEN 'на балансе'
                    WHEN "Nodes"."onBalance" = 1 THEN 'потребительская'
                    WHEN "Nodes"."onBalance" = 2 THEN 'ССО'
                    WHEN "Nodes"."onBalance" = 3 THEN 'ССП'
                    END) AS v_Nodes_onBalance,
                    "Transf" AS v_Nodes_Transf
                    FROM "Nodes") AS t_Nodes
                    ON(t_LVConsumersA_LVNodes_Fiders.v_LVConsumersA_LVNodes_OwnerGUID = t_Nodes.v_Nodes_GUID) OR(t_LVConsumersA_LVNodes_Fiders.v_LVFiders_OwnerGUID = t_Nodes.v_Nodes_GUID)

                    LEFT JOIN
                    (SELECT "Ident" AS v_Transforms2_Ident, "Unom" AS v_Transforms2_Unom, "TypeTR" AS v_Transforms2_TypeTR, "Snom" AS v_Transforms2_Snom
                    FROM "Transforms2") AS t_Transforms2
                    ON t_Nodes.v_Nodes_Transf = t_Transforms2.v_Transforms2_Ident

                    LEFT JOIN
                    (SELECT "GUID" AS v_Fiders_GUID, "OwnerGUID" AS v_Fiders_OwnerGUID, "Name" AS v_Fiders_Name
                    FROM "Fiders") AS t_Fiders
                    ON t_Nodes.v_Nodes_OwnerGUID = t_Fiders.v_Fiders_GUID

                    LEFT JOIN
                    (SELECT "GUID" AS v_Sections_GUID, "OwnerGUID" AS v_Sections_OwnerGUID, "Name" AS v_Sections_Name, "Unom" AS v_Sections_Unom
                    FROM "Sections") AS t_Sections
                    ON t_Fiders.v_Fiders_OwnerGUID = t_Sections.v_Sections_GUID

                    LEFT JOIN
                    (SELECT "GUID" AS v_Centers_GUID, "OwnerGUID" AS v_Centers_OwnerGUID, "Name" AS v_Centers_Name
                    FROM "Centers") AS t_Centers
                    ON t_Sections.v_Sections_OwnerGUID = t_Centers.v_Centers_GUID

                    ON t_LVConsumersInfo.guid = v_LVConsumersA_LVNodes_GUID
                    );*/

                FbDataAdapter dataAdapter = new FbDataAdapter();
                dataAdapter.SelectCommand = command;
                DataTable dt = new DataTable();
                connection.Open();

                splashScreenManager1.SetWaitFormDescription("загрузка данных...");
                dataAdapter.Fill(dt);

                connection.Close();

                // формируем отчет
                splashScreenManager1.SetWaitFormDescription("формирование отчета...");

                FormReportGrid formReport = null;
                formReport = new FormReportGrid();
                formReport.MdiParent = this;
                formReport.Text = String.Concat("Реестр потребителей РТП-3 (", openFileDialogRTP.FileName, ")");
                IWorkbook workbook = formReport.spreadsheetControl1.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                workbook.History.IsEnabled = false;
                formReport.spreadsheetControl1.BeginUpdate();

                int row_columns_group_names = 0;
                int row_start_data = 3;
                int row_columns_names = 1;
                int row_columns_index = 2;

                // выводим и форматируем названия групп столбцов
                worksheet.Rows[row_columns_group_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_group_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[row_columns_group_names].Font.Name = "Arial";
                worksheet.Rows[row_columns_group_names].Font.Size = 8;
                
                worksheet[row_columns_group_names, 0].SetValue("№ п/п"); worksheet.MergeCells(worksheet.Range["A1:A2"]);
                worksheet[row_columns_group_names, 1].SetValue("Данные потребителя"); worksheet.MergeCells(worksheet.Range["B1:F1"]);
                worksheet[row_columns_group_names, 6].SetValue("Наименование на схеме"); worksheet.MergeCells(worksheet.Range["G1:I1"]);
                worksheet[row_columns_group_names, 9].SetValue("Наименование фидера 0,4 кВ"); worksheet.MergeCells(worksheet.Range["J1:L1"]);
                worksheet[row_columns_group_names, 12].SetValue("ТП 6(10)/0,4 кВ"); worksheet.MergeCells(worksheet.Range["M1:Q1"]);
                worksheet[row_columns_group_names, 17].SetValue("Свойства ТП 6(10)/0,4 кВ"); worksheet.MergeCells(worksheet.Range["R1:U1"]);
                worksheet[row_columns_group_names, 21].SetValue("Наименование фидера ВН"); worksheet.MergeCells(worksheet.Range["V1:X1"]);
                worksheet[row_columns_group_names, 24].SetValue("Секция шин"); worksheet.MergeCells(worksheet.Range["Y1:AB1"]);
                worksheet[row_columns_group_names, 28].SetValue("Наименование ПС"); worksheet.MergeCells(worksheet.Range["AC1:AE1"]);

                // выводим названия столбцов
                string[] columns_names = {
                    "", "GUID", "Код точки поставки", "Наименование", "Точка учета", "Адрес",
                    "GUID", "OwnerGUID", "Наименование на схеме",
                    "GUID", "OwnerGUID", "Наименование фидера 0,4 кВ",
                    "GUID", "OwnerGUID", "Наименование ТП", "Балансовая принадлежность", "Номер типа трансформатора",
                    "Номер типа трансформатора", "Напряжение", "Тип", "Мощность",
                    "GUID", "OwnerGUID", "Наименование фидера ВН",
                    "GUID", "OwnerGUID", "Номер секции шин", "Номинальное напряжение",
                    "GUID", "OwnerGUID", "Наименование ПС"
                    };

                for (int col = 0; col < 31; col++)
                {
                    worksheet[row_columns_names, col].SetValue(columns_names[col]);                    
                    worksheet[row_columns_index, col].SetValue((col + 1).ToString());
                }

                // форматируем заголовок таблицы                                
                worksheet.Rows[row_columns_names].Alignment.WrapText = true;                
                worksheet.Rows[row_columns_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Rows[row_columns_names].Font.Name = "Arial";
                worksheet.Rows[row_columns_names].Font.Size = 8;

                worksheet.Rows[row_columns_index].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Rows[row_columns_index].Font.Name = "Arial";
                worksheet.Rows[row_columns_index].Font.Size = 8;

                worksheet.Range["A1:AE3"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                // выводим данные
                for (int row = 0; row < dt.Rows.Count; row++)
                {
                    worksheet.Rows[row + row_start_data].Font.Name = "Arial";
                    worksheet.Rows[row + row_start_data].Font.Size = 8;

                    int row_index = row + 1;
                    worksheet[row + row_start_data, 0].SetValue(row_index.ToString());

                    for (int col = 0; col < dt.Columns.Count; col++)
                    {
                        worksheet[row + row_start_data, col + 1].SetValue(dt.Rows[row][col].ToString());
                    }
                }

                // выравнивание столбцов
                worksheet.Columns[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                splashScreenManager1.SetWaitFormDescription("выравнивание ширины столбцов...");
                worksheet.Columns.AutoFit(0, dt.Columns.Count);

                formReport.spreadsheetControl1.EndUpdate();

                splashScreenManager1.CloseWaitForm();
                                
                formReport.Show();

                this.ribbonControl1.SelectedPage = this.ribbonControl1.MergedPages[0];
            }
                        
        } // private void barButtonItemRTP3_ConsumersInfo_ItemClick(object sender, ItemClickEventArgs e)

        // РТП-3 - Администрирование - Загрузить БД
        private void barButtonItemRTP3_Admin_LoadRTP3DB_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormRTP3dbLoadParams formRTP3dbLoadParams = new FormRTP3dbLoadParams();        

            if (formRTP3dbLoadParams.ShowDialog() == DialogResult.OK)
            {
                splashScreenManager1.ShowWaitForm();

                splashScreenManager1.SetWaitFormDescription("подключение к базе данных...");

                // создаем соединение
                FbConnection connection = new FbConnection();
                connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", formRTP3dbLoadParams.textEditRTP3DBFile.Text, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");                
                FbCommand command = new FbCommand();
                command.Connection = connection;

                FbDataAdapter dataAdapter = new FbDataAdapter();
                dataAdapter.SelectCommand = command;
                DataTable dt = new DataTable();
                connection.Open();

                RTP3ESHEntities rtp3eshEntities = new RTP3ESHEntities();
                //--------------
                                                
                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["GUID"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица GUID...");
                    command.CommandText = "SELECT \"GUID\", \"RELATION_ID\" FROM \"GUID\"";
                    dataAdapter.Fill(dt);                                
                    List<RTP3_GUID> rtp3_GUID_List = new List<RTP3_GUID>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_GUID rtp3_GUID = new RTP3_GUID();
                        rtp3_GUID.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_GUID.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["RELATION_ID"].ToString())) rtp3_GUID.RELATION_ID = Convert.ToInt32(dataRow["RELATION_ID"].ToString());                        
                        else rtp3_GUID.RELATION_ID = null;
                        rtp3_GUID_List.Add(rtp3_GUID);
                    }
                    rtp3eshEntities.RTP3_GUID.AddRange(rtp3_GUID_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["ReferencesTypes"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица ReferencesTypes...");
                    command.CommandText = "SELECT \"Ident\", \"Name\" FROM \"ReferencesTypes\"";
                    dataAdapter.Fill(dt);                
                    List<RTP3_ReferencesTypes> rtp3_ReferencesTypes_List = new List<RTP3_ReferencesTypes>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_ReferencesTypes rtp3_ReferencesTypes = new RTP3_ReferencesTypes();
                        rtp3_ReferencesTypes.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_ReferencesTypes.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_ReferencesTypes.Name = dataRow["Name"].ToString();
                        rtp3_ReferencesTypes_List.Add(rtp3_ReferencesTypes);
                    }
                    rtp3eshEntities.RTP3_ReferencesTypes.AddRange(rtp3_ReferencesTypes_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["References"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица References...");
                    command.CommandText = "SELECT \"Ident\", \"ReferenceType\" FROM \"References\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_References> rtp3_References_List = new List<RTP3_References>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_References rtp3_References = new RTP3_References();
                        rtp3_References.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_References.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_References.ReferenceType = int.Parse(dataRow["ReferenceType"].ToString());
                        rtp3_References_List.Add(rtp3_References);
                    }
                    rtp3eshEntities.RTP3_References.AddRange(rtp3_References_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Unom"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Unom...");
                    command.CommandText = "SELECT \"Unom\" FROM \"Unom\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Unom> rtp3_Unom_List = new List<RTP3_Unom>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_Unom rtp3_Unom = new RTP3_Unom();
                        rtp3_Unom.Unom = Decimal.Parse(dataRow["Unom"].ToString());
                        if (!rtp3eshEntities.RTP3_Unom.Where(p => p.Unom == rtp3_Unom.Unom).Select(p => p).Any()) rtp3_Unom_List.Add(rtp3_Unom);
                    }
                    rtp3eshEntities.RTP3_Unom.AddRange(rtp3_Unom_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Owners"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Owners...");
                    command.CommandText = "SELECT \"Ident\", \"Name\" FROM \"Owners\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Owners> rtp3_Owners_List = new List<RTP3_Owners>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_Owners rtp3_Owners = new RTP3_Owners();
                        rtp3_Owners.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Owners.Ident = short.Parse(dataRow["Ident"].ToString());
                        rtp3_Owners.Name = dataRow["Name"].ToString();
                        rtp3_Owners_List.Add(rtp3_Owners);
                    }
                    rtp3eshEntities.RTP3_Owners.AddRange(rtp3_Owners_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Structures"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Structures...");
                    command.CommandText = "SELECT \"Ident\", \"Name\", \"ShortName\" FROM \"Structures\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Structures> rtp3_Structures_List = new List<RTP3_Structures>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_Structures rtp3_Structures = new RTP3_Structures();
                        rtp3_Structures.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Structures.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_Structures.Name = dataRow["Name"].ToString();
                        rtp3_Structures.ShortName = dataRow["ShortName"].ToString();
                        rtp3_Structures_List.Add(rtp3_Structures);
                    }
                    rtp3eshEntities.RTP3_Structures.AddRange(rtp3_Structures_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Coors"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Coors...");
                    command.CommandText = "SELECT \"GUID\", \"X\", \"Y\", \"deltaX\", \"deltaY\", \"Rule\", \"Additional\" FROM \"Coors\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Coors> rtp3_Coors_List = new List<RTP3_Coors>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_Coors rtp3_Coors = new RTP3_Coors();
                        rtp3_Coors.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Coors.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["X"].ToString())) rtp3_Coors.X = Convert.ToInt32(dataRow["X"].ToString()); else rtp3_Coors.X = null;
                        if (!String.IsNullOrEmpty(dataRow["Y"].ToString())) rtp3_Coors.Y = Convert.ToInt32(dataRow["Y"].ToString()); else rtp3_Coors.Y = null;
                        if (!String.IsNullOrEmpty(dataRow["deltaX"].ToString())) rtp3_Coors.deltaX = Convert.ToInt32(dataRow["deltaX"].ToString()); else rtp3_Coors.deltaX = null;
                        if (!String.IsNullOrEmpty(dataRow["deltaY"].ToString())) rtp3_Coors.deltaY = Convert.ToInt32(dataRow["deltaY"].ToString()); else rtp3_Coors.deltaY = null;
                        if (!String.IsNullOrEmpty(dataRow["Rule"].ToString())) rtp3_Coors.Rule_ = Convert.ToInt32(dataRow["Rule"].ToString()); else rtp3_Coors.Rule_ = null;
                        if (!String.IsNullOrEmpty(dataRow["Additional"].ToString())) rtp3_Coors.Additional = Convert.ToInt32(dataRow["Additional"].ToString()); else rtp3_Coors.Additional = null;
                        rtp3_Coors_List.Add(rtp3_Coors);
                    }
                    rtp3eshEntities.RTP3_Coors.AddRange(rtp3_Coors_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Regions"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Regions...");
                    command.CommandText = "SELECT \"Ident\" FROM \"Regions\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Regions> rtp3_Regions_List = new List<RTP3_Regions>();
                    foreach (DataRow dataRow in dt.Rows)
                    {                        
                        RTP3_Regions rtp3_Regions = new RTP3_Regions();
                        rtp3_Regions.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Regions.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_Regions_List.Add(rtp3_Regions);
                    }
                    rtp3eshEntities.RTP3_Regions.AddRange(rtp3_Regions_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["AOEnergo"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица AOEnergo...");
                    command.CommandText = "SELECT \"GUID\", \"Region\", \"Name\" FROM \"AOEnergo\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_AOEnergo> rtp3_AOEnergo_List = new List<RTP3_AOEnergo>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_AOEnergo rtp3_AOEnergo = new RTP3_AOEnergo();
                        rtp3_AOEnergo.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_AOEnergo.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_AOEnergo.Region = int.Parse(dataRow["Region"].ToString());
                        rtp3_AOEnergo.Name = dataRow["Name"].ToString();
                        rtp3_AOEnergo_List.Add(rtp3_AOEnergo);
                    }
                    rtp3eshEntities.RTP3_AOEnergo.AddRange(rtp3_AOEnergo_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["PES"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица PES...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Name\" FROM \"PES\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_PES> rtp3_PES_List = new List<RTP3_PES>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_PES rtp3_PES = new RTP3_PES();
                        rtp3_PES.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_PES.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_PES.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_PES.Name = dataRow["Name"].ToString();
                        rtp3_PES_List.Add(rtp3_PES);
                    }
                    rtp3eshEntities.RTP3_PES.AddRange(rtp3_PES_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["RES"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица RES...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Name\" FROM \"RES\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_RES> rtp3_RES_List = new List<RTP3_RES>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_RES rtp3_RES = new RTP3_RES();
                        rtp3_RES.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_RES.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_RES.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_RES.Name = dataRow["Name"].ToString();
                        rtp3_RES_List.Add(rtp3_RES);
                    }
                    rtp3eshEntities.RTP3_RES.AddRange(rtp3_RES_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Centers"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Centers...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Name\", \"DBB\" FROM \"Centers\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Centers> rtp3_Centers_List = new List<RTP3_Centers>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Centers rtp3_Centers = new RTP3_Centers();
                        rtp3_Centers.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Centers.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_Centers.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_Centers.Name = dataRow["Name"].ToString();
                        if (!String.IsNullOrEmpty(dataRow["DBB"].ToString())) rtp3_Centers.DBB = Convert.ToDateTime(dataRow["DBB"].ToString()); else rtp3_Centers.DBB = null;
                        rtp3_Centers_List.Add(rtp3_Centers);
                    }
                    rtp3eshEntities.RTP3_Centers.AddRange(rtp3_Centers_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Sections"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Sections...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Name\", \"Unom\", \"Rc\", \"Xc\" FROM \"Sections\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Sections> rtp3_Sections_List = new List<RTP3_Sections>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Sections rtp3_Sections = new RTP3_Sections();
                        rtp3_Sections.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Sections.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_Sections.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_Sections.Name = dataRow["Name"].ToString();
                        rtp3_Sections.Unom = Decimal.Parse(dataRow["Unom"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["Rc"].ToString())) rtp3_Sections.Rc = Convert.ToDecimal(dataRow["Rc"].ToString()); else rtp3_Sections.Rc = null;
                        if (!String.IsNullOrEmpty(dataRow["Xc"].ToString())) rtp3_Sections.Xc = Convert.ToDecimal(dataRow["Xc"].ToString()); else rtp3_Sections.Xc = null;
                        rtp3_Sections_List.Add(rtp3_Sections);
                    }
                    rtp3eshEntities.RTP3_Sections.AddRange(rtp3_Sections_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Fiders"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Fiders...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Name\", \"fCircuit\", \"DBB\", \"Dkor\" FROM \"Fiders\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Fiders> rtp3_Fiders_List = new List<RTP3_Fiders>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Fiders rtp3_Fiders = new RTP3_Fiders();
                        rtp3_Fiders.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Fiders.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_Fiders.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_Fiders.Name = dataRow["Name"].ToString();
                        if (!String.IsNullOrEmpty(dataRow["fCircuit"].ToString())) rtp3_Fiders.fCircuit = Convert.ToInt16(dataRow["fCircuit"].ToString()); else rtp3_Fiders.fCircuit = null;
                        if (!String.IsNullOrEmpty(dataRow["DBB"].ToString())) rtp3_Fiders.DBB = Convert.ToDateTime(dataRow["DBB"].ToString()); else rtp3_Fiders.DBB = null;
                        if (!String.IsNullOrEmpty(dataRow["Dkor"].ToString())) rtp3_Fiders.Dkor = Convert.ToDateTime(dataRow["Dkor"].ToString()); else rtp3_Fiders.Dkor = null;
                        rtp3_Fiders_List.Add(rtp3_Fiders);
                        //rtp3eshEntities.RTP3_Fiders.Add(rtp3_Fiders);
                        //rtp3eshEntities.SaveChanges();
                    }
                    rtp3eshEntities.RTP3_Fiders.AddRange(rtp3_Fiders_List);                    
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Lines"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Lines...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Node1\", \"Node2\", \"Info\", \"Provod\", \"State\", \"LineLength\", \"Kpl\", \"onBalance\", \"Idop1\", \"Idop2\", \"Structures\", \"DBB\", \"TabOrder\" FROM \"Lines\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Lines> rtp3_Lines_List = new List<RTP3_Lines>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Lines rtp3_Lines = new RTP3_Lines();
                        rtp3_Lines.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Lines.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_Lines.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_Lines.Node1 = int.Parse(dataRow["Node1"].ToString());
                        rtp3_Lines.Node2 = int.Parse(dataRow["Node2"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["Info"].ToString())) rtp3_Lines.Info = dataRow["Info"].ToString(); else rtp3_Lines.Info = null;
                        rtp3_Lines.Provod = int.Parse(dataRow["Provod"].ToString());
                        rtp3_Lines.State_ = short.Parse(dataRow["State"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["LineLength"].ToString())) rtp3_Lines.LineLength = Convert.ToDecimal(dataRow["LineLength"].ToString()); else rtp3_Lines.LineLength = null;
                        if (!String.IsNullOrEmpty(dataRow["Kpl"].ToString())) rtp3_Lines.Kpl = Convert.ToInt16(dataRow["Kpl"].ToString()); else rtp3_Lines.Kpl = null;
                        if (!String.IsNullOrEmpty(dataRow["onBalance"].ToString())) rtp3_Lines.onBalance = Convert.ToInt16(dataRow["onBalance"].ToString()); else rtp3_Lines.onBalance = null;
                        if (!String.IsNullOrEmpty(dataRow["Idop1"].ToString())) rtp3_Lines.Idop1 = Convert.ToDecimal(dataRow["Idop1"].ToString()); else rtp3_Lines.Idop1 = null;
                        if (!String.IsNullOrEmpty(dataRow["Idop2"].ToString())) rtp3_Lines.Idop2 = Convert.ToDecimal(dataRow["Idop2"].ToString()); else rtp3_Lines.Idop2 = null;
                        if (!String.IsNullOrEmpty(dataRow["Structures"].ToString())) rtp3_Lines.Structures = Convert.ToInt32(dataRow["Structures"].ToString()); else rtp3_Lines.Structures = null;
                        if (!String.IsNullOrEmpty(dataRow["DBB"].ToString())) rtp3_Lines.DBB = Convert.ToDateTime(dataRow["DBB"].ToString()); else rtp3_Lines.DBB = null;
                        if (!String.IsNullOrEmpty(dataRow["TabOrder"].ToString())) rtp3_Lines.TabOrder = Convert.ToInt16(dataRow["TabOrder"].ToString()); else rtp3_Lines.TabOrder = null;
                        rtp3_Lines_List.Add(rtp3_Lines);
                    }
                    rtp3eshEntities.RTP3_Lines.AddRange(rtp3_Lines_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["TPType"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица TPType...");
                    command.CommandText = "SELECT \"Ident\", \"Name\", \"ShortName\", \"PictureIdent\" FROM \"TPType\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_TPType> rtp3_TPType_List = new List<RTP3_TPType>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_TPType rtp3_TPType = new RTP3_TPType();
                        rtp3_TPType.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_TPType.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_TPType.Name = dataRow["Name"].ToString();                        
                        if (!String.IsNullOrEmpty(dataRow["ShortName"].ToString())) rtp3_TPType.ShortName = dataRow["ShortName"].ToString(); else rtp3_TPType.ShortName = null;
                        if (!String.IsNullOrEmpty(dataRow["PictureIdent"].ToString())) rtp3_TPType.PictureIdent = Convert.ToInt32(dataRow["PictureIdent"].ToString()); else rtp3_TPType.PictureIdent = null;
                        rtp3_TPType_List.Add(rtp3_TPType);
                    }
                    rtp3eshEntities.RTP3_TPType.AddRange(rtp3_TPType_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["TypicalLoadCurves"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица TypicalLoadCurves...");
                    command.CommandText = "SELECT \"Ident\", \"Name\", \"Unom\", \"LoadType\", \"kzap1\", \"sqrKfg1\", \"kzap2\", \"sqrKfg2\", \"Comment\" FROM \"TypicalLoadCurves\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_TypicalLoadCurves> rtp3_TypicalLoadCurves_List = new List<RTP3_TypicalLoadCurves>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_TypicalLoadCurves rtp3_TypicalLoadCurves = new RTP3_TypicalLoadCurves();
                        rtp3_TypicalLoadCurves.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_TypicalLoadCurves.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_TypicalLoadCurves.Name = dataRow["Name"].ToString();
                        rtp3_TypicalLoadCurves.Unom = Decimal.Parse(dataRow["Unom"].ToString());
                        rtp3_TypicalLoadCurves.LoadType = int.Parse(dataRow["LoadType"].ToString());
                        rtp3_TypicalLoadCurves.kzap1 = double.Parse(dataRow["kzap1"].ToString());
                        rtp3_TypicalLoadCurves.sqrKfg1 = double.Parse(dataRow["sqrKfg1"].ToString());
                        rtp3_TypicalLoadCurves.kzap2 = double.Parse(dataRow["kzap2"].ToString());
                        rtp3_TypicalLoadCurves.sqrKfg2 = double.Parse(dataRow["sqrKfg2"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["Comment"].ToString())) rtp3_TypicalLoadCurves.Comment = dataRow["Comment"].ToString(); else rtp3_TypicalLoadCurves.Comment = null;                        
                        rtp3_TypicalLoadCurves_List.Add(rtp3_TypicalLoadCurves);
                    }
                    rtp3eshEntities.RTP3_TypicalLoadCurves.AddRange(rtp3_TypicalLoadCurves_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Nodes"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Nodes...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Info\", \"Transf\", \"PositionOtp\", \"TPType\", \"onBalance\", \"DBB\", \"TabOrder\", \"TypicalLoadCurve\", \"Name\" FROM \"Nodes\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Nodes> rtp3_Nodes_List = new List<RTP3_Nodes>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Nodes rtp3_Nodes = new RTP3_Nodes();
                        rtp3_Nodes.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Nodes.GUID = Convert.ToInt32(dataRow["GUID"].ToString());
                        rtp3_Nodes.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_Nodes.Info = dataRow["Info"].ToString();
                        rtp3_Nodes.Transf = int.Parse(dataRow["Transf"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["PositionOtp"].ToString())) rtp3_Nodes.PositionOtp = Convert.ToInt32(dataRow["PositionOtp"].ToString()); else rtp3_Nodes.PositionOtp = null;
                        if (!String.IsNullOrEmpty(dataRow["TPType"].ToString())) rtp3_Nodes.TPType = Convert.ToInt32(dataRow["TPType"].ToString()); else rtp3_Nodes.TPType = null;
                        if (!String.IsNullOrEmpty(dataRow["onBalance"].ToString())) rtp3_Nodes.onBalance = Convert.ToInt16(dataRow["onBalance"].ToString()); else rtp3_Nodes.onBalance = null;
                        if (!String.IsNullOrEmpty(dataRow["DBB"].ToString())) rtp3_Nodes.DBB = Convert.ToDateTime(dataRow["DBB"].ToString()); else rtp3_Nodes.DBB = null;
                        if (!String.IsNullOrEmpty(dataRow["TabOrder"].ToString())) rtp3_Nodes.TabOrder = Convert.ToInt16(dataRow["TabOrder"].ToString()); rtp3_Nodes.TabOrder = null;
                        if (!String.IsNullOrEmpty(dataRow["TypicalLoadCurve"].ToString())) rtp3_Nodes.TypicalLoadCurve = Convert.ToInt32(dataRow["TypicalLoadCurve"].ToString()); else rtp3_Nodes.TypicalLoadCurve = null;
                        if (!String.IsNullOrEmpty(dataRow["Name"].ToString())) rtp3_Nodes.Name = dataRow["Name"].ToString(); else rtp3_Nodes.Name = null;
                        rtp3_Nodes_List.Add(rtp3_Nodes);
                    }
                    rtp3eshEntities.RTP3_Nodes.AddRange(rtp3_Nodes_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["ExLines"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица ExLines...");
                    command.CommandText = "SELECT \"GUID1\", \"GUID2\", \"State\" FROM \"ExLines\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_ExLines> rtp3_ExLines_List = new List<RTP3_ExLines>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_ExLines rtp3_ExLines = new RTP3_ExLines();
                        rtp3_ExLines.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_ExLines.GUID1 = Convert.ToInt32(dataRow["GUID1"].ToString());
                        rtp3_ExLines.GUID2 = Convert.ToInt32(dataRow["GUID2"].ToString());
                        rtp3_ExLines.State_ = Convert.ToInt32(dataRow["State"].ToString());
                        rtp3_ExLines_List.Add(rtp3_ExLines);
                    }
                    rtp3eshEntities.RTP3_ExLines.AddRange(rtp3_ExLines_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["TypeLoads"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица TypeLoads...");
                    command.CommandText = "SELECT \"Ident\", \"CosFi\", \"Name\" FROM \"TypeLoads\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_TypeLoads> rtp3_TypeLoads_List = new List<RTP3_TypeLoads>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_TypeLoads rtp3_TypeLoads = new RTP3_TypeLoads();
                        rtp3_TypeLoads.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_TypeLoads.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_TypeLoads.CosFi = Convert.ToDecimal(dataRow["CosFi"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["Name"].ToString())) rtp3_TypeLoads.Name = dataRow["Name"].ToString(); else rtp3_TypeLoads.Name = null;                        
                        rtp3_TypeLoads_List.Add(rtp3_TypeLoads);
                    }
                    rtp3eshEntities.RTP3_TypeLoads.AddRange(rtp3_TypeLoads_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["ProvodTypes"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица ProvodTypes...");
                    command.CommandText = "SELECT \"Ident\", \"Unom\", \"Name\", \"CableFlags\", \"PictureIdent\" FROM \"ProvodTypes\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_ProvodTypes> rtp3_ProvodTypes_List = new List<RTP3_ProvodTypes>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_ProvodTypes rtp3_ProvodTypes = new RTP3_ProvodTypes();
                        rtp3_ProvodTypes.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_ProvodTypes.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_ProvodTypes.Unom = Convert.ToDecimal(dataRow["Unom"].ToString());
                        rtp3_ProvodTypes.Name = dataRow["Name"].ToString();
                        rtp3_ProvodTypes.CableFlags = int.Parse(dataRow["CableFlags"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["PictureIdent"].ToString())) rtp3_ProvodTypes.PictureIdent = Convert.ToInt32(dataRow["PictureIdent"].ToString()); else rtp3_ProvodTypes.PictureIdent = null;
                        rtp3_ProvodTypes_List.Add(rtp3_ProvodTypes);
                    }
                    rtp3eshEntities.RTP3_ProvodTypes.AddRange(rtp3_ProvodTypes_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Manufacturers"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Manufacturers...");
                    command.CommandText = "SELECT \"Ident\", \"ReferenceType\", \"Manufacturer\", \"Address\", \"PhoneNumber\", \"FaxNumber\", \"Email\", \"WebPage\" FROM \"Manufacturers\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Manufacturers> rtp3_Manufacturers_List = new List<RTP3_Manufacturers>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Manufacturers rtp3_Manufacturers = new RTP3_Manufacturers();
                        rtp3_Manufacturers.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Manufacturers.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_Manufacturers.ReferenceType = int.Parse(dataRow["ReferenceType"].ToString());
                        rtp3_Manufacturers.Manufacturer = dataRow["Manufacturer"].ToString();
                        if (!String.IsNullOrEmpty(dataRow["Address"].ToString())) rtp3_Manufacturers.Address = dataRow["Address"].ToString(); else rtp3_Manufacturers.Address = null;
                        if (!String.IsNullOrEmpty(dataRow["PhoneNumber"].ToString())) rtp3_Manufacturers.PhoneNumber = dataRow["PhoneNumber"].ToString(); else rtp3_Manufacturers.PhoneNumber = null;
                        if (!String.IsNullOrEmpty(dataRow["FaxNumber"].ToString())) rtp3_Manufacturers.FaxNumber = dataRow["FaxNumber"].ToString(); else rtp3_Manufacturers.FaxNumber = null;
                        if (!String.IsNullOrEmpty(dataRow["Email"].ToString())) rtp3_Manufacturers.Email = dataRow["Email"].ToString(); else rtp3_Manufacturers.Email = null;
                        if (!String.IsNullOrEmpty(dataRow["WebPage"].ToString())) rtp3_Manufacturers.WebPage = dataRow["WebPage"].ToString(); else rtp3_Manufacturers.WebPage = null;
                        rtp3_Manufacturers_List.Add(rtp3_Manufacturers);
                    }
                    rtp3eshEntities.RTP3_Manufacturers.AddRange(rtp3_Manufacturers_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Provods"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Provods...");
                    command.CommandText = "SELECT \"Ident\", \"ProvodType\", \"Marka\", \"F\", \"AddStr\", \"R0\", \"X0\", \"B0\", \"Idop\", \"G0\", \"TgDelta\", \"Manufacturer\", \"Hide\" FROM \"Provods\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Provods> rtp3_Provods_List = new List<RTP3_Provods>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Provods rtp3_Provods = new RTP3_Provods();
                        rtp3_Provods.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Provods.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_Provods.ProvodType = int.Parse(dataRow["ProvodType"].ToString());
                        rtp3_Provods.Marka = dataRow["Marka"].ToString();
                        rtp3_Provods.F = Convert.ToDecimal(dataRow["F"].ToString());
                        rtp3_Provods.AddStr = dataRow["AddStr"].ToString();
                        rtp3_Provods.R0 = Convert.ToDecimal(dataRow["R0"].ToString());
                        rtp3_Provods.X0 = Convert.ToDecimal(dataRow["X0"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["B0"].ToString())) rtp3_Provods.B0 = Convert.ToDecimal(dataRow["B0"].ToString()); else rtp3_Provods.B0 = null;
                        rtp3_Provods.Idop = int.Parse(dataRow["Idop"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["G0"].ToString())) rtp3_Provods.G0 = Convert.ToDecimal(dataRow["G0"].ToString()); else rtp3_Provods.G0 = null;
                        if (!String.IsNullOrEmpty(dataRow["TgDelta"].ToString())) rtp3_Provods.TgDelta = Convert.ToDecimal(dataRow["TgDelta"].ToString()); else rtp3_Provods.TgDelta = null;
                        if (!String.IsNullOrEmpty(dataRow["Manufacturer"].ToString())) rtp3_Provods.Manufacturer = Convert.ToInt32(dataRow["Manufacturer"].ToString()); else rtp3_Provods.Manufacturer = null;
                        if (!String.IsNullOrEmpty(dataRow["Hide"].ToString())) rtp3_Provods.Hide = Convert.ToInt16(dataRow["Hide"].ToString()); else rtp3_Provods.Hide = null;
                        rtp3_Provods_List.Add(rtp3_Provods);
                    }
                    rtp3eshEntities.RTP3_Provods.AddRange(rtp3_Provods_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["DoubleLines"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица DoubleLines...");
                    command.CommandText = "SELECT \"GUID\", \"Provod\", \"LineLength\" FROM \"DoubleLines\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_DoubleLines> rtp3_DoubleLines_List = new List<RTP3_DoubleLines>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_DoubleLines rtp3_DoubleLines = new RTP3_DoubleLines();
                        rtp3_DoubleLines.GUID_PK = Guid.NewGuid();
                        rtp3_DoubleLines.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_DoubleLines.GUID = int.Parse(dataRow["GUID"].ToString());
                        rtp3_DoubleLines.Provod = int.Parse(dataRow["Provod"].ToString());
                        rtp3_DoubleLines.LineLength = Convert.ToDecimal(dataRow["LineLength"].ToString());                        
                        rtp3_DoubleLines_List.Add(rtp3_DoubleLines);
                    }
                    rtp3eshEntities.RTP3_DoubleLines.AddRange(rtp3_DoubleLines_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["LVFiders"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица LVFiders...");
                    command.CommandText = "SELECT \"GUID\", \"Name\", \"Unom\", \"OwnerGUID\", \"TypeLoad\", \"DBB\", \"Dkor\" FROM \"LVFiders\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_LVFiders> rtp3_LVFiders_List = new List<RTP3_LVFiders>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_LVFiders rtp3_LVFiders = new RTP3_LVFiders();                        
                        rtp3_LVFiders.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_LVFiders.GUID = int.Parse(dataRow["GUID"].ToString());
                        rtp3_LVFiders.Name = dataRow["Name"].ToString();
                        rtp3_LVFiders.Unom = Convert.ToDecimal(dataRow["Unom"].ToString());
                        rtp3_LVFiders.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["TypeLoad"].ToString())) rtp3_LVFiders.TypeLoad = Convert.ToInt32(dataRow["TypeLoad"].ToString()); else rtp3_LVFiders.TypeLoad = null;
                        if (!String.IsNullOrEmpty(dataRow["DBB"].ToString())) rtp3_LVFiders.DBB = Convert.ToDateTime(dataRow["DBB"].ToString()); else rtp3_LVFiders.DBB = null;
                        if (!String.IsNullOrEmpty(dataRow["Dkor"].ToString())) rtp3_LVFiders.Dkor = Convert.ToDateTime(dataRow["Dkor"].ToString()); else rtp3_LVFiders.Dkor = null;
                        rtp3_LVFiders_List.Add(rtp3_LVFiders);
                    }
                    rtp3eshEntities.RTP3_LVFiders.AddRange(rtp3_LVFiders_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["LVNodes"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица LVNodes...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Info\", \"TabOrder\", \"Name\" FROM \"LVNodes\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_LVNodes> rtp3_LVNodes_List = new List<RTP3_LVNodes>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_LVNodes rtp3_LVNodes = new RTP3_LVNodes();
                        rtp3_LVNodes.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_LVNodes.GUID = int.Parse(dataRow["GUID"].ToString());                        
                        rtp3_LVNodes.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());                                                                      
                        if (!String.IsNullOrEmpty(dataRow["Info"].ToString())) rtp3_LVNodes.Info = dataRow["Info"].ToString(); else rtp3_LVNodes.Info = null;
                        if (!String.IsNullOrEmpty(dataRow["TabOrder"].ToString())) rtp3_LVNodes.TabOrder = Convert.ToInt16(dataRow["TabOrder"].ToString()); else rtp3_LVNodes.TabOrder = null;
                        if (!String.IsNullOrEmpty(dataRow["Name"].ToString())) rtp3_LVNodes.Name = dataRow["Name"].ToString(); else rtp3_LVNodes.Name = null;

                        rtp3_LVNodes_List.Add(rtp3_LVNodes);
                    }
                    rtp3eshEntities.RTP3_LVNodes.AddRange(rtp3_LVNodes_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["LVLines"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица LVLines...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Node1\", \"Node2\", \"Info\", \"Phase\", \"ProvodPhase\", \"ProvodNull\", \"State\", \"LineLength\", \"Kpl\", \"onBalance\", \"Idop1\", \"Idop2\", \"Structures\", \"DBB\", \"TabOrder\" FROM \"LVLines\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_LVLines> rtp3_LVLines_List = new List<RTP3_LVLines>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_LVLines rtp3_LVLines = new RTP3_LVLines();
                        rtp3_LVLines.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_LVLines.GUID = int.Parse(dataRow["GUID"].ToString());
                        rtp3_LVLines.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        rtp3_LVLines.Node1 = int.Parse(dataRow["Node1"].ToString());
                        rtp3_LVLines.Node2 = int.Parse(dataRow["Node2"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["Info"].ToString())) rtp3_LVLines.Info = dataRow["Info"].ToString(); else rtp3_LVLines.Info = null;
                        if (!String.IsNullOrEmpty(dataRow["Phase"].ToString())) rtp3_LVLines.Phase = Convert.ToInt16(dataRow["Phase"].ToString()); else rtp3_LVLines.Phase = null;
                        if (!String.IsNullOrEmpty(dataRow["ProvodPhase"].ToString())) rtp3_LVLines.ProvodPhase = Convert.ToInt32(dataRow["ProvodPhase"].ToString()); else rtp3_LVLines.ProvodPhase = null;
                        if (!String.IsNullOrEmpty(dataRow["ProvodNull"].ToString())) rtp3_LVLines.ProvodNull = Convert.ToInt32(dataRow["ProvodNull"].ToString()); else rtp3_LVLines.ProvodNull = null;
                        rtp3_LVLines.State_ = int.Parse(dataRow["State"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["LineLength"].ToString())) rtp3_LVLines.LineLength = Convert.ToDecimal(dataRow["LineLength"].ToString()); else rtp3_LVLines.LineLength = null;
                        if (!String.IsNullOrEmpty(dataRow["Kpl"].ToString())) rtp3_LVLines.Kpl = Convert.ToInt16(dataRow["Kpl"].ToString()); else rtp3_LVLines.Kpl = null;
                        if (!String.IsNullOrEmpty(dataRow["onBalance"].ToString())) rtp3_LVLines.onBalance = Convert.ToInt16(dataRow["onBalance"].ToString()); else rtp3_LVLines.onBalance = null;
                        if (!String.IsNullOrEmpty(dataRow["Idop1"].ToString())) rtp3_LVLines.Idop1 = Convert.ToDecimal(dataRow["Idop1"].ToString()); else rtp3_LVLines.Idop1 = null;
                        if (!String.IsNullOrEmpty(dataRow["Idop2"].ToString())) rtp3_LVLines.Idop2 = Convert.ToDecimal(dataRow["Idop2"].ToString()); else rtp3_LVLines.Idop2 = null;
                        if (!String.IsNullOrEmpty(dataRow["Structures"].ToString())) rtp3_LVLines.Structures = Convert.ToInt32(dataRow["Structures"].ToString()); else rtp3_LVLines.Structures = null;                        
                        if (!String.IsNullOrEmpty(dataRow["DBB"].ToString())) rtp3_LVLines.DBB = Convert.ToDateTime(dataRow["DBB"].ToString()); else rtp3_LVLines.DBB = null;
                        if (!String.IsNullOrEmpty(dataRow["TabOrder"].ToString())) rtp3_LVLines.TabOrder = Convert.ToInt16(dataRow["TabOrder"].ToString()); else rtp3_LVLines.TabOrder = null;
                        rtp3_LVLines_List.Add(rtp3_LVLines);
                    }
                    rtp3eshEntities.RTP3_LVLines.AddRange(rtp3_LVLines_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Nodes3"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Nodes3...");
                    command.CommandText = "SELECT \"GUID_HH\", \"GUID_CH\", \"TypicalLoadCurveCH\" FROM \"Nodes3\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Nodes3> rtp3_Nodes3_List = new List<RTP3_Nodes3>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Nodes3 rtp3_Nodes3 = new RTP3_Nodes3();
                        rtp3_Nodes3.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Nodes3.GUID_HH = int.Parse(dataRow["GUID_HH"].ToString());
                        rtp3_Nodes3.GUID_CH = int.Parse(dataRow["GUID_CH"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["TypicalLoadCurveCH"].ToString())) rtp3_Nodes3.TypicalLoadCurveCH = Convert.ToInt32(dataRow["TypicalLoadCurveCH"].ToString()); else rtp3_Nodes3.TypicalLoadCurveCH = null;
                        rtp3_Nodes3_List.Add(rtp3_Nodes3);
                    }
                    rtp3eshEntities.RTP3_Nodes3.AddRange(rtp3_Nodes3_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Transforms2"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Transforms2...");
                    command.CommandText = "SELECT \"Ident\", \"Unom\", \"TypeTR\", \"Snom\", \"AddStr\", \"Uk\", \"Pxx\", \"Pkz\", \"Ixx\", \"UnomB\", \"UnomH\", \"NumberOtp\", \"DeltaUotp\", \"Manufacturer\", \"Hide\" FROM \"Transforms2\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_Transforms2> rtp3_Transforms2_List = new List<RTP3_Transforms2>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Transforms2 rtp3_Transforms2 = new RTP3_Transforms2();
                        rtp3_Transforms2.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Transforms2.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_Transforms2.Unom = Convert.ToDecimal(dataRow["Unom"].ToString());
                        rtp3_Transforms2.TypeTR = dataRow["TypeTR"].ToString();
                        rtp3_Transforms2.Snom = Convert.ToDecimal(dataRow["Snom"].ToString());
                        rtp3_Transforms2.AddStr = dataRow["AddStr"].ToString();
                        rtp3_Transforms2.Uk = Convert.ToDecimal(dataRow["Uk"].ToString());
                        rtp3_Transforms2.Pxx = Convert.ToDecimal(dataRow["Pxx"].ToString());
                        rtp3_Transforms2.Pkz = Convert.ToDecimal(dataRow["Pkz"].ToString());
                        rtp3_Transforms2.Ixx = Convert.ToDecimal(dataRow["Ixx"].ToString());
                        rtp3_Transforms2.UnomB = Convert.ToDecimal(dataRow["UnomB"].ToString());
                        rtp3_Transforms2.UnomH = Convert.ToDecimal(dataRow["UnomH"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["NumberOtp"].ToString())) rtp3_Transforms2.NumberOtp = Convert.ToInt32(dataRow["NumberOtp"].ToString()); else rtp3_Transforms2.NumberOtp = null;
                        if (!String.IsNullOrEmpty(dataRow["DeltaUotp"].ToString())) rtp3_Transforms2.DeltaUotp = Convert.ToDecimal(dataRow["DeltaUotp"].ToString()); else rtp3_Transforms2.DeltaUotp = null;
                        if (!String.IsNullOrEmpty(dataRow["Manufacturer"].ToString())) rtp3_Transforms2.Manufacturer = Convert.ToInt32(dataRow["Manufacturer"].ToString()); else rtp3_Transforms2.Manufacturer = null;
                        if (!String.IsNullOrEmpty(dataRow["Hide"].ToString())) rtp3_Transforms2.Hide = Convert.ToInt16(dataRow["Hide"].ToString()); else rtp3_Transforms2.Hide = null;
                        rtp3_Transforms2_List.Add(rtp3_Transforms2);                        
                    }
                    rtp3eshEntities.RTP3_Transforms2.AddRange(rtp3_Transforms2_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["Transforms3"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица Transforms3...");
                    command.CommandText = String.Concat(
                        "SELECT \"Ident\", \"Unom\", \"TypeTR\", \"Snom\", \"AddStr\", ",
                        "\"UnomB\", \"UnomC\", \"UnomH\", \"Pxx\", \"Ixx\", \"PkzBC\", \"PkzBH\", \"PkzCH\", \"UkBC\", \"UkBH\", \"UkCH\", ",
                        "\"SobmB\", \"SobmC\", \"SobmH\", \"NumberOtpB\", \"DeltaUotpB\", \"NumberOtpC\", \"DeltaUotpC\", \"Manufacturer\", \"Hide\" FROM \"Transforms3\"");
                    dataAdapter.Fill(dt);
                    List<RTP3_Transforms3> rtp3_Transforms3_List = new List<RTP3_Transforms3>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_Transforms3 rtp3_Transforms3 = new RTP3_Transforms3();
                        rtp3_Transforms3.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_Transforms3.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_Transforms3.Unom = Convert.ToDecimal(dataRow["Unom"].ToString());
                        rtp3_Transforms3.TypeTR = dataRow["TypeTR"].ToString();
                        rtp3_Transforms3.Snom = Convert.ToDecimal(dataRow["Snom"].ToString());
                        rtp3_Transforms3.AddStr = dataRow["AddStr"].ToString();
                        rtp3_Transforms3.UnomB = Convert.ToDecimal(dataRow["UnomB"].ToString());
                        rtp3_Transforms3.UnomC = Convert.ToDecimal(dataRow["UnomC"].ToString());
                        rtp3_Transforms3.UnomH = Convert.ToDecimal(dataRow["UnomH"].ToString());
                        rtp3_Transforms3.Pxx = Convert.ToDecimal(dataRow["Pxx"].ToString());                        
                        rtp3_Transforms3.Ixx = Convert.ToDecimal(dataRow["Ixx"].ToString());
                        rtp3_Transforms3.PkzBC = Convert.ToDecimal(dataRow["PkzBC"].ToString());
                        rtp3_Transforms3.PkzBH = Convert.ToDecimal(dataRow["PkzBH"].ToString());
                        rtp3_Transforms3.PkzCH = Convert.ToDecimal(dataRow["PkzCH"].ToString());
                        rtp3_Transforms3.UkBC = Convert.ToDecimal(dataRow["UkBC"].ToString());
                        rtp3_Transforms3.UkBH = Convert.ToDecimal(dataRow["UkBH"].ToString());
                        rtp3_Transforms3.UkCH = Convert.ToDecimal(dataRow["UkCH"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["SobmB"].ToString())) rtp3_Transforms3.SobmB = Convert.ToDecimal(dataRow["SobmB"].ToString()); else rtp3_Transforms3.SobmB = null;
                        if (!String.IsNullOrEmpty(dataRow["SobmC"].ToString())) rtp3_Transforms3.SobmC = Convert.ToDecimal(dataRow["SobmC"].ToString()); else rtp3_Transforms3.SobmC = null;
                        if (!String.IsNullOrEmpty(dataRow["SobmH"].ToString())) rtp3_Transforms3.SobmH = Convert.ToDecimal(dataRow["SobmH"].ToString()); else rtp3_Transforms3.SobmH = null;
                        if (!String.IsNullOrEmpty(dataRow["NumberOtpB"].ToString())) rtp3_Transforms3.NumberOtpB = Convert.ToInt32(dataRow["NumberOtpB"].ToString()); else rtp3_Transforms3.NumberOtpB = null;
                        if (!String.IsNullOrEmpty(dataRow["DeltaUotpB"].ToString())) rtp3_Transforms3.DeltaUotpB = Convert.ToDecimal(dataRow["DeltaUotpB"].ToString()); else rtp3_Transforms3.DeltaUotpB = null;
                        if (!String.IsNullOrEmpty(dataRow["NumberOtpC"].ToString())) rtp3_Transforms3.NumberOtpC = Convert.ToInt32(dataRow["NumberOtpC"].ToString()); else rtp3_Transforms3.NumberOtpC = null;
                        if (!String.IsNullOrEmpty(dataRow["DeltaUotpC"].ToString())) rtp3_Transforms3.DeltaUotpC = Convert.ToDecimal(dataRow["DeltaUotpC"].ToString()); else rtp3_Transforms3.DeltaUotpC = null;
                        if (!String.IsNullOrEmpty(dataRow["Manufacturer"].ToString())) rtp3_Transforms3.Manufacturer = Convert.ToInt32(dataRow["Manufacturer"].ToString()); else rtp3_Transforms3.Manufacturer = null;
                        if (!String.IsNullOrEmpty(dataRow["Hide"].ToString())) rtp3_Transforms3.Hide = Convert.ToInt16(dataRow["Hide"].ToString()); else rtp3_Transforms3.Hide = null;
                        rtp3_Transforms3_List.Add(rtp3_Transforms3);
                    }
                    rtp3eshEntities.RTP3_Transforms3.AddRange(rtp3_Transforms3_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["SwitchGN"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица SwitchGN...");
                    command.CommandText = "SELECT \"Ident\", \"Name\", \"PictureIdent\" FROM \"SwitchGN\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_SwitchGN> rtp3_SwitchGN_List = new List<RTP3_SwitchGN>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_SwitchGN rtp3_SwitchGN = new RTP3_SwitchGN();
                        rtp3_SwitchGN.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_SwitchGN.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_SwitchGN.Name = dataRow["Name"].ToString();
                        if (!String.IsNullOrEmpty(dataRow["PictureIdent"].ToString())) rtp3_SwitchGN.PictureIdent = Convert.ToInt32(dataRow["PictureIdent"].ToString()); else rtp3_SwitchGN.PictureIdent = null;
                        rtp3_SwitchGN_List.Add(rtp3_SwitchGN);
                    }
                    rtp3eshEntities.RTP3_SwitchGN.AddRange(rtp3_SwitchGN_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["SwitchGR"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица SwitchGR...");
                    command.CommandText = "SELECT \"Ident\", \"Unom\", \"SwitchGType\", \"Marka\", \"AddStr\", \"Inom\", \"Ioff\", \"Id\", \"It\", \"Toff\", \"Ton\", \"Manufacturer\", \"Hide\" FROM \"SwitchGR\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_SwitchGR> rtp3_SwitchGR_List = new List<RTP3_SwitchGR>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_SwitchGR rtp3_SwitchGR = new RTP3_SwitchGR();
                        rtp3_SwitchGR.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_SwitchGR.Ident = int.Parse(dataRow["Ident"].ToString());
                        rtp3_SwitchGR.Unom = Convert.ToDecimal(dataRow["Unom"].ToString());
                        rtp3_SwitchGR.SwitchGType = int.Parse(dataRow["SwitchGType"].ToString());
                        rtp3_SwitchGR.Marka = dataRow["Marka"].ToString();
                        rtp3_SwitchGR.AddStr = dataRow["AddStr"].ToString();                                                
                        if (!String.IsNullOrEmpty(dataRow["Inom"].ToString())) rtp3_SwitchGR.Inom = Convert.ToDecimal(dataRow["Inom"].ToString()); else rtp3_SwitchGR.Inom = null;
                        if (!String.IsNullOrEmpty(dataRow["Ioff"].ToString())) rtp3_SwitchGR.Ioff = Convert.ToDecimal(dataRow["Ioff"].ToString()); else rtp3_SwitchGR.Ioff = null;
                        if (!String.IsNullOrEmpty(dataRow["Id"].ToString())) rtp3_SwitchGR.Id = Convert.ToDecimal(dataRow["Id"].ToString()); else rtp3_SwitchGR.Id = null;
                        if (!String.IsNullOrEmpty(dataRow["It"].ToString())) rtp3_SwitchGR.It = Convert.ToDecimal(dataRow["It"].ToString()); else rtp3_SwitchGR.It = null;
                        if (!String.IsNullOrEmpty(dataRow["Toff"].ToString())) rtp3_SwitchGR.Toff = Convert.ToDecimal(dataRow["Toff"].ToString()); else rtp3_SwitchGR.Toff = null;
                        if (!String.IsNullOrEmpty(dataRow["Ton"].ToString())) rtp3_SwitchGR.Ton = Convert.ToDecimal(dataRow["Ton"].ToString()); else rtp3_SwitchGR.Ton = null;
                        if (!String.IsNullOrEmpty(dataRow["Manufacturer"].ToString())) rtp3_SwitchGR.Manufacturer = Convert.ToInt32(dataRow["Manufacturer"].ToString()); else rtp3_SwitchGR.Manufacturer = null;
                        if (!String.IsNullOrEmpty(dataRow["Hide"].ToString())) rtp3_SwitchGR.Hide = Convert.ToInt16(dataRow["Hide"].ToString()); else rtp3_SwitchGR.Hide = null;
                        rtp3_SwitchGR_List.Add(rtp3_SwitchGR);
                    }
                    rtp3eshEntities.RTP3_SwitchGR.AddRange(rtp3_SwitchGR_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["LVConsumers"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица LVConsumers...");
                    command.CommandText = "SELECT \"GUID\", \"onBalance\", \"Phase\", \"TypeLoad\", \"Pust\", \"CosFi\", \"TypicalLoadCurve\" FROM \"LVConsumers\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_LVConsumers> rtp3_LVConsumers_List = new List<RTP3_LVConsumers>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_LVConsumers rtp3_LVConsumers = new RTP3_LVConsumers();
                        rtp3_LVConsumers.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_LVConsumers.GUID = int.Parse(dataRow["GUID"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["onBalance"].ToString())) rtp3_LVConsumers.onBalance = Convert.ToInt16(dataRow["onBalance"].ToString()); else rtp3_LVConsumers.onBalance = null;
                        if (!String.IsNullOrEmpty(dataRow["Phase"].ToString())) rtp3_LVConsumers.Phase = Convert.ToInt16(dataRow["Phase"].ToString()); else rtp3_LVConsumers.Phase = null;
                        if (!String.IsNullOrEmpty(dataRow["TypeLoad"].ToString())) rtp3_LVConsumers.TypeLoad = Convert.ToInt32(dataRow["TypeLoad"].ToString()); else rtp3_LVConsumers.TypeLoad = null;
                        if (!String.IsNullOrEmpty(dataRow["Pust"].ToString())) rtp3_LVConsumers.Pust = Convert.ToDecimal(dataRow["Pust"].ToString()); else rtp3_LVConsumers.Pust = null;
                        if (!String.IsNullOrEmpty(dataRow["CosFi"].ToString())) rtp3_LVConsumers.CosFi = Convert.ToDecimal(dataRow["CosFi"].ToString()); else rtp3_LVConsumers.CosFi = null;
                        if (!String.IsNullOrEmpty(dataRow["TypicalLoadCurve"].ToString())) rtp3_LVConsumers.TypicalLoadCurve = Convert.ToInt32(dataRow["TypicalLoadCurve"].ToString()); else rtp3_LVConsumers.TypicalLoadCurve = null;
                        rtp3_LVConsumers_List.Add(rtp3_LVConsumers);
                    }
                    rtp3eshEntities.RTP3_LVConsumers.AddRange(rtp3_LVConsumers_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["LVConsumersA"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица LVConsumersA...");
                    command.CommandText = "SELECT \"GUID\", \"OwnerGUID\", \"Info\", \"TabOrder\", \"Name\" FROM \"LVConsumersA\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_LVConsumersA> rtp3_LVConsumersA_List = new List<RTP3_LVConsumersA>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_LVConsumersA rtp3_LVConsumersA = new RTP3_LVConsumersA();
                        rtp3_LVConsumersA.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_LVConsumersA.GUID = int.Parse(dataRow["GUID"].ToString());
                        rtp3_LVConsumersA.OwnerGUID = int.Parse(dataRow["OwnerGUID"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["Info"].ToString())) rtp3_LVConsumersA.Info = dataRow["Info"].ToString(); else rtp3_LVConsumersA.Info = null;
                        if (!String.IsNullOrEmpty(dataRow["TabOrder"].ToString())) rtp3_LVConsumersA.TabOrder = Convert.ToInt16(dataRow["TabOrder"].ToString()); else rtp3_LVConsumersA.TabOrder = null;
                        if (!String.IsNullOrEmpty(dataRow["Name"].ToString())) rtp3_LVConsumersA.Name = dataRow["Name"].ToString(); else rtp3_LVConsumersA.Name = null;
                        rtp3_LVConsumersA_List.Add(rtp3_LVConsumersA);
                    }
                    rtp3eshEntities.RTP3_LVConsumersA.AddRange(rtp3_LVConsumersA_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkedListBoxControlRTP3Tables.Items["LVConsumersInfo"].CheckState == CheckState.Checked)
                {
                    dt.Clear();
                    dt.Columns.Clear();
                    splashScreenManager1.SetWaitFormDescription("таблица LVConsumersInfo...");
                    command.CommandText = "SELECT \"GUID\", \"PointName\", \"Contract\", \"Address\", \"Telephone\", \"Customer\", \"ContractDate\", \"INN\", \"Power\", \"Kmeter\", \"MeterValue\" FROM \"LVConsumersInfo\"";
                    dataAdapter.Fill(dt);
                    List<RTP3_LVConsumersInfo> rtp3_LVConsumersInfo_List = new List<RTP3_LVConsumersInfo>();
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        RTP3_LVConsumersInfo rtp3_LVConsumersInfo = new RTP3_LVConsumersInfo();
                        rtp3_LVConsumersInfo.ID_DB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                        rtp3_LVConsumersInfo.GUID = int.Parse(dataRow["GUID"].ToString());
                        if (!String.IsNullOrEmpty(dataRow["PointName"].ToString())) rtp3_LVConsumersInfo.PointName = dataRow["PointName"].ToString(); else rtp3_LVConsumersInfo.PointName = null;
                        if (!String.IsNullOrEmpty(dataRow["Contract"].ToString())) rtp3_LVConsumersInfo.Contract = dataRow["Contract"].ToString(); else rtp3_LVConsumersInfo.Contract = null;
                        if (!String.IsNullOrEmpty(dataRow["Address"].ToString())) rtp3_LVConsumersInfo.Address = dataRow["Address"].ToString(); else rtp3_LVConsumersInfo.Address = null;
                        if (!String.IsNullOrEmpty(dataRow["Telephone"].ToString())) rtp3_LVConsumersInfo.Telephone = dataRow["Telephone"].ToString(); else rtp3_LVConsumersInfo.Telephone = null;
                        if (!String.IsNullOrEmpty(dataRow["Customer"].ToString())) rtp3_LVConsumersInfo.Customer = dataRow["Customer"].ToString(); else rtp3_LVConsumersInfo.Customer = null;
                        if (!String.IsNullOrEmpty(dataRow["ContractDate"].ToString())) rtp3_LVConsumersInfo.ContractDate = Convert.ToDateTime(dataRow["ContractDate"].ToString()); else rtp3_LVConsumersInfo.ContractDate = null;
                        if (!String.IsNullOrEmpty(dataRow["INN"].ToString())) rtp3_LVConsumersInfo.INN = dataRow["INN"].ToString(); else rtp3_LVConsumersInfo.INN = null;
                        if (!String.IsNullOrEmpty(dataRow["Power"].ToString())) rtp3_LVConsumersInfo.Power = dataRow["Power"].ToString(); else rtp3_LVConsumersInfo.Power = null;
                        if (!String.IsNullOrEmpty(dataRow["Kmeter"].ToString())) rtp3_LVConsumersInfo.Kmeter = Convert.ToDecimal(dataRow["Kmeter"].ToString()); else rtp3_LVConsumersInfo.Kmeter = null;
                        if (!String.IsNullOrEmpty(dataRow["MeterValue"].ToString())) rtp3_LVConsumersInfo.MeterValue = Convert.ToDecimal(dataRow["MeterValue"].ToString()); else rtp3_LVConsumersInfo.MeterValue = null;
                        rtp3_LVConsumersInfo_List.Add(rtp3_LVConsumersInfo);
                    }
                    rtp3eshEntities.RTP3_LVConsumersInfo.AddRange(rtp3_LVConsumersInfo_List);
                    rtp3eshEntities.SaveChanges();
                }

                if (formRTP3dbLoadParams.checkEditFillConsumersInfo.CheckState == CheckState.Checked)
                {                     
                    splashScreenManager1.SetWaitFormDescription("таблица ConsumersInfo...");                                        
                    dt.Clear();
                    dt.Columns.Clear();
                    string line;
                    StreamReader sr = new StreamReader(@"sql_RTP3.txt");
                    command.CommandText = null;
                    while ((line = sr.ReadLine()) != null)
                    {
                        command.CommandText = String.Concat(command.CommandText, line, Environment.NewLine);
                    }
                    dataAdapter.Fill(dt);

                    List<ConsumersInfo> consumersInfo_List = new List<ConsumersInfo>();
                    int idDB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                    foreach (DataRow dataRow in dt.Rows)
                    {
                        ConsumersInfo consumersInfo = new ConsumersInfo();
                        consumersInfo.id_consumersinfo = Guid.NewGuid();
                        consumersInfo.ID_DB = idDB;                        

                        int consumerGUID = Convert.ToInt32(dataRow["GUID"]);
                        int centerGUID = Convert.ToInt32(dataRow["v_Centers_GUID"]);
                        int sectionGUID = Convert.ToInt32(dataRow["v_Sections_GUID"]);
                        int fiderGUID = Convert.ToInt32(dataRow["v_Fiders_GUID"]);
                        int? transforms2Ident = null;
                        if (!String.IsNullOrEmpty(dataRow["v_Transforms2_Ident"].ToString())) transforms2Ident = Convert.ToInt32(dataRow["v_Transforms2_Ident"]);
                        int? nodeGUID = null;
                        if (!String.IsNullOrEmpty(dataRow["v_Nodes_GUID"].ToString())) nodeGUID = Convert.ToInt32(dataRow["v_Nodes_GUID"]);
                        int? lvfiderGUID = null;
                        if (!String.IsNullOrEmpty(dataRow["v_LVFiders_GUID"].ToString())) lvfiderGUID = Convert.ToInt32(dataRow["v_LVFiders_GUID"]);

                        consumersInfo.GUIDRTP3 = consumerGUID;
                        consumersInfo.CenterGUID = centerGUID;
                        consumersInfo.SectionGUID = sectionGUID;
                        consumersInfo.FiderGUID = fiderGUID;
                        consumersInfo.Transforms2_Ident = transforms2Ident;
                        consumersInfo.NodeGUID = nodeGUID;
                        consumersInfo.LVFiderGUID = lvfiderGUID;

                        consumersInfo.ContractRTP3 = dataRow["Contract"].ToString();

                        /*consumersInfo.codeIESBK = pvalues[4];
                        consumersInfo.codeOKE = pvalues[5];
                        consumersInfo.addressInfo = pvalues[6];
                        if (pvalues[7].Equals("0") || String.IsNullOrEmpty(pvalues[7])) consumersInfo.phoneInfo = null;
                        else consumersInfo.phoneInfo = pvalues[7];
                        if (pvalues[8].Equals("0") || String.IsNullOrEmpty(pvalues[8])) consumersInfo.emailInfo = null;
                        else consumersInfo.emailInfo = pvalues[8];*/

                        consumersInfo_List.Add(consumersInfo);
                    }
                    rtp3eshEntities.ConsumersInfo.AddRange(consumersInfo_List);
                    rtp3eshEntities.SaveChanges();
                    
                    // -------------------------------

                }

                //--------------
                splashScreenManager1.SetWaitFormDescription("завершение...");
                connection.Close();
                splashScreenManager1.CloseWaitForm();

                MessageBox.Show(text: "Загрузка успешно завершена!", caption: "Информация");
            }
        } // private void barButtonItemRTP3_Admin_LoadRTP3DB_ItemClick(object sender, ItemClickEventArgs e)

        // "Информация об отключениях"
        private void barButtonBlackouts_ItemClick(object sender, ItemClickEventArgs e)
        {            
            splashScreenManager1.ShowWaitForm();
            splashScreenManager1.SetWaitFormDescription("построение дерева сети...");
            FormBlackouts form1 = new FormBlackouts(dbconnectionStringBlackouts, permissionAdmin);
            form1.Text = "Формирование перечня отключений";
            form1.MdiParent = this;
                        
            form1.Show();
            ribbonControl1.SelectedPage = ribbonControl1.MergedPages["Отключения"];
            splashScreenManager1.CloseWaitForm();
        }

        // привязка потребителей (загрузка из РТП-3)
        private void barButtonItem37_ItemClick(object sender, ItemClickEventArgs e)
        {
            /*OpenFileDialog openFileDialogRTP = new OpenFileDialog();
            openFileDialogRTP.Title = "Выберите файл базы данных РТП-3";
            openFileDialogRTP.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
            openFileDialogRTP.FileName = "";

            if (openFileDialogRTP.ShowDialog() == DialogResult.OK)
            {*/
            FormRTP3dbLoadParams formRTP3dbLoadParams = new FormRTP3dbLoadParams();

            if (formRTP3dbLoadParams.ShowDialog() == DialogResult.OK)
            {
                splashScreenManager1.ShowWaitForm();

                splashScreenManager1.SetWaitFormDescription("подключение к базе данных...");

                // создаем соединение и загружаем данные
                FbConnection connection = new FbConnection();                
                //connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", openFileDialogRTP.FileName, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");
                connection.ConnectionString = String.Concat("UserID=SYSDBA;Password=masterkey;Database=", formRTP3dbLoadParams.textEditRTP3DBFile.Text, ";DataSource=", Environment.MachineName, ";Charset=UTF8;");

                FbCommand command = new FbCommand();
                command.Connection = connection;

                // читаем запрос из файла (для теста, потом убрать!!!)
                command.CommandText = "";
                string line;
                StreamReader sr = new StreamReader(@"sql_RTP3.txt");
                while ((line = sr.ReadLine()) != null)
                {
                    command.CommandText = String.Concat(command.CommandText, line, Environment.NewLine);
                }
                
                FbDataAdapter dataAdapter = new FbDataAdapter();
                dataAdapter.SelectCommand = command;
                DataTable dt = new DataTable();
                connection.Open();

                splashScreenManager1.SetWaitFormDescription("загрузка данных...");
                dataAdapter.Fill(dt);

                connection.Close();

                // обновляем данные о привязке потребителей
                RTP3ESHEntities rtp3eshEntities = new RTP3ESHEntities();
                int idDB = (int)formRTP3dbLoadParams.comboBoxEditRTP3DB.EditValue;
                IQueryable<ConsumersInfo> consumersInfo_List = rtp3eshEntities.ConsumersInfo.Where(p => p.ID_DB == idDB).Select(p => p);

                foreach (DataRow dr in dt.Rows)
                {
                    int consumerGUID = Convert.ToInt32(dr["GUID"]);
                    int centerGUID = Convert.ToInt32(dr["v_Centers_GUID"]);
                    int sectionGUID = Convert.ToInt32(dr["v_Sections_GUID"]);
                    int fiderGUID = Convert.ToInt32(dr["v_Fiders_GUID"]);

                    int? transforms2Ident = null;
                    if (!String.IsNullOrEmpty(dr["v_Transforms2_Ident"].ToString())) transforms2Ident = Convert.ToInt32(dr["v_Transforms2_Ident"]);

                    int? nodeGUID = null;
                    if (!String.IsNullOrEmpty(dr["v_Nodes_GUID"].ToString())) nodeGUID = Convert.ToInt32(dr["v_Nodes_GUID"]);

                    int? lvfiderGUID = null;
                    if (!String.IsNullOrEmpty(dr["v_LVFiders_GUID"].ToString())) lvfiderGUID = Convert.ToInt32(dr["v_LVFiders_GUID"]);

                    try
                    {
                        ConsumersInfo consumerInfo = consumersInfo_List.First(p => /*p.ID_DB == idDB && */p.GUIDRTP3 == consumerGUID);
                        consumerInfo.CenterGUID = centerGUID;
                        consumerInfo.SectionGUID = sectionGUID;
                        consumerInfo.FiderGUID = fiderGUID;
                        consumerInfo.Transforms2_Ident = transforms2Ident;
                        consumerInfo.NodeGUID = nodeGUID;
                        consumerInfo.LVFiderGUID = lvfiderGUID;
                    }
                    catch (InvalidOperationException)
                    {

                    }
                }

                rtp3eshEntities.SaveChanges();

                //--------------
                splashScreenManager1.SetWaitFormDescription("завершение...");                
                splashScreenManager1.CloseWaitForm();

                MessageBox.Show(text: "Загрузка успешно завершена!", caption: "Информация");
            }
        } // private void barButtonItem37_ItemClick(object sender, ItemClickEventArgs e)

        // загрузка информационных полей из 1С Транспорт электроэнергии
        private void barButtonItemLoadDataFrom1C_ItemClick(object sender, ItemClickEventArgs e)
        {
            splashScreenManager1.ShowWaitForm();

            splashScreenManager1.SetWaitFormDescription("загрузка данных...");

            // создаем соединение с БД 1С ТЭ            
            string dbconnectionString = "Data Source=;Initial Catalog=transport;User ID=;Password=";
            SqlConnection connection = new SqlConnection(dbconnectionString);
            connection.Open();
            SqlDataAdapter adapter = new SqlDataAdapter();

            DataTable dt = new DataTable();

            RTP3ESHEntities rtp3eshEntities = new RTP3ESHEntities();            
            IQueryable<ConsumersInfo> consumersInfo_List = rtp3eshEntities.ConsumersInfo.Select(p => p);

            int i = 0;

            // запрос для ФЛ ----------ФЛ
            string queryStringFL = String.Concat(
                //"SELECT TOP 1000",
                "SELECT",
                " _Marked,", /*ПометкаУдаления*/
                "_Code,",   /*Код*/
                "_Fld1389,",    /*КодИЭСБК*/
                "_Fld1390,",    /*КодОКЭ*/
                "_Fld1402,",    /*ПредставлениеАдреса*/
                "_Fld6660,",    /*ТелефонДляУведомлений*/
                "_Fld6661,",    /*ЭлектроннаяПочтаДляУведомлений*/
                "_Fld6907",	/*ДомФИАСID*/
                " FROM _Reference82",
                //" WHERE _Code = '000158601'"
                " WHERE _Marked = 0"
            );
            adapter.SelectCommand = new SqlCommand(queryStringFL, connection);
            adapter.SelectCommand.CommandTimeout = 0;
            adapter.Fill(dt);            
            foreach (DataRow dr in dt.Rows)
            {                
                try
                {                    
                    string codeOKE = dr["_Fld1390"].ToString();
                    string codeIESBK = dr["_Fld1389"].ToString();
                    ConsumersInfo consumerInfo = consumersInfo_List.First(p => String.Equals(p.ContractRTP3, codeIESBK)); // ищем по коду ИЭСБК
                    
                    consumerInfo.codeIESBK = codeIESBK;
                    consumerInfo.codeOKE = codeOKE;
                    consumerInfo.addressInfo = dr["_Fld1402"].ToString();
                    string phoneInfo = dr["_Fld6660"].ToString();
                    if (phoneInfo.Equals("0") || String.IsNullOrEmpty(phoneInfo)) consumerInfo.phoneInfo = null;
                    else consumerInfo.phoneInfo = phoneInfo;
                    string emailInfo = dr["_Fld6661"].ToString();
                    if (emailInfo.Equals("0") || String.IsNullOrEmpty(emailInfo)) consumerInfo.emailInfo = null;
                    else consumerInfo.emailInfo = emailInfo;

                    if (!String.IsNullOrEmpty(dr["_Fld6907"].ToString()))
                    {
                        byte[] houseFIASid_bytes_bad = (byte[])dr["_Fld6907"];
                        byte[] houseFIASid_bytes_good =
                            new byte[]
                            {
                            houseFIASid_bytes_bad[15], houseFIASid_bytes_bad[14], houseFIASid_bytes_bad[13], houseFIASid_bytes_bad[12],
                            houseFIASid_bytes_bad[11], houseFIASid_bytes_bad[10], houseFIASid_bytes_bad[9], houseFIASid_bytes_bad[8],
                            houseFIASid_bytes_bad[0], houseFIASid_bytes_bad[1], houseFIASid_bytes_bad[2], houseFIASid_bytes_bad[3],
                            houseFIASid_bytes_bad[4], houseFIASid_bytes_bad[5], houseFIASid_bytes_bad[6], houseFIASid_bytes_bad[7]
                            };
                        Guid houseFIASid_guid = new Guid(houseFIASid_bytes_good);
                        consumerInfo.houseFIASid = houseFIASid_guid;
                    }
                }
                catch (InvalidOperationException)
                {

                }
                splashScreenManager1.SetWaitFormDescription("Обработка данных ФЛ (" + (++i).ToString() + " из " + dt.Rows.Count.ToString() + ")");
            }
            rtp3eshEntities.SaveChanges();

            i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    string codeOKE = dr["_Fld1390"].ToString();
                    string codeIESBK = dr["_Fld1389"].ToString();
                    
                    string codeCode = dr["_Code"].ToString();
                    ConsumersInfo consumerInfo = consumersInfo_List.First(p => String.Equals(p.ContractRTP3, codeCode)); // ищем по внутреннему коду 1С

                    consumerInfo.codeIESBK = codeIESBK;
                    consumerInfo.codeOKE = codeOKE;
                    consumerInfo.addressInfo = dr["_Fld1402"].ToString();
                    string phoneInfo = dr["_Fld6660"].ToString();
                    if (phoneInfo.Equals("0") || String.IsNullOrEmpty(phoneInfo)) consumerInfo.phoneInfo = null;
                    else consumerInfo.phoneInfo = phoneInfo;
                    string emailInfo = dr["_Fld6661"].ToString();
                    if (emailInfo.Equals("0") || String.IsNullOrEmpty(emailInfo)) consumerInfo.emailInfo = null;
                    else consumerInfo.emailInfo = emailInfo;

                    /*string houseFIASid_str = BitConverter.ToString((byte[])dr["_Fld6907"]).Replace("-", "");
                    Guid houseFIASid_guid = new Guid(houseFIASid_str.Replace("-", ""));
                    consumerInfo.houseFIASid = houseFIASid_guid;*/

                    if (!String.IsNullOrEmpty(dr["_Fld6907"].ToString()))
                    {
                        byte[] houseFIASid_bytes_bad = (byte[])dr["_Fld6907"];
                        byte[] houseFIASid_bytes_good =
                            new byte[]
                            {
                            houseFIASid_bytes_bad[15], houseFIASid_bytes_bad[14], houseFIASid_bytes_bad[13], houseFIASid_bytes_bad[12],
                            houseFIASid_bytes_bad[11], houseFIASid_bytes_bad[10], houseFIASid_bytes_bad[9], houseFIASid_bytes_bad[8],
                            houseFIASid_bytes_bad[0], houseFIASid_bytes_bad[1], houseFIASid_bytes_bad[2], houseFIASid_bytes_bad[3],
                            houseFIASid_bytes_bad[4], houseFIASid_bytes_bad[5], houseFIASid_bytes_bad[6], houseFIASid_bytes_bad[7]
                            };
                        Guid houseFIASid_guid = new Guid(houseFIASid_bytes_good);
                        consumerInfo.houseFIASid = houseFIASid_guid;
                    }
                }
                catch (InvalidOperationException)
                {

                }
                splashScreenManager1.SetWaitFormDescription("Обработка данных ФЛ (" + (++i).ToString() + " из " + dt.Rows.Count.ToString() + ")");
            }
            rtp3eshEntities.SaveChanges();
            //----------------------ФЛ
            
            dt.Clear();
            dt.Columns.Clear();

            // запрос для МКД ----------МКД
            string queryStringMKD = String.Concat(
                //"SELECT TOP 1000",
                "SELECT",
                " dtTUMKD._Marked,", /*ПометкаУдаления*/
                "dtTUMKD._Code,",   /*Код*/
                "dtTUMKD._Fld4998,",    /*КодИЭСБК*/
                "dtTUMKD._Fld4999,",    /*КодОКЭ*/
                "dtTUMKD._Fld6743,",    /*ТелефонДляУведомлений*/
                "dtTUMKD._Fld6744,",    /*ЭлектроннаяПочтаДляУведомлений*/

                "dtMKD._Fld796,",   /*ПредставлениеАдреса*/
                "dtMKD._Fld6855",	/*ДомФИАСID*/

                " FROM _Reference4990 dtTUMKD",
                " LEFT JOIN _Reference48 dtMKD",
                " ON dtTUMKD._Fld5001RRef = dtMKD._IDRRef",

                " WHERE dtTUMKD._Marked = 0"                
            );
            adapter.SelectCommand = new SqlCommand(queryStringMKD, connection);
            adapter.SelectCommand.CommandTimeout = 0;
            adapter.Fill(dt);
            i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    string codeOKE = dr["_Fld4999"].ToString();
                    ConsumersInfo consumerInfo = consumersInfo_List.First(p => String.Equals(p.ContractRTP3, codeOKE)); // ищем по коду ОКЭ

                    consumerInfo.codeIESBK = dr["_Fld4998"].ToString();
                    consumerInfo.codeOKE = codeOKE;
                    consumerInfo.addressInfo = dr["_Fld796"].ToString();
                    string phoneInfo = dr["_Fld6743"].ToString();
                    if (phoneInfo.Equals("0") || String.IsNullOrEmpty(phoneInfo)) consumerInfo.phoneInfo = null;
                    else consumerInfo.phoneInfo = phoneInfo;
                    string emailInfo = dr["_Fld6744"].ToString();
                    if (emailInfo.Equals("0") || String.IsNullOrEmpty(emailInfo)) consumerInfo.emailInfo = null;
                    else consumerInfo.emailInfo = emailInfo;

                    if (!String.IsNullOrEmpty(dr["_Fld6855"].ToString()))
                    {
                        byte[] houseFIASid_bytes_bad = (byte[])dr["_Fld6855"];
                        byte[] houseFIASid_bytes_good =
                            new byte[]
                            {
                            houseFIASid_bytes_bad[15], houseFIASid_bytes_bad[14], houseFIASid_bytes_bad[13], houseFIASid_bytes_bad[12],
                            houseFIASid_bytes_bad[11], houseFIASid_bytes_bad[10], houseFIASid_bytes_bad[9], houseFIASid_bytes_bad[8],
                            houseFIASid_bytes_bad[0], houseFIASid_bytes_bad[1], houseFIASid_bytes_bad[2], houseFIASid_bytes_bad[3],
                            houseFIASid_bytes_bad[4], houseFIASid_bytes_bad[5], houseFIASid_bytes_bad[6], houseFIASid_bytes_bad[7]
                            };
                        Guid houseFIASid_guid = new Guid(houseFIASid_bytes_good);
                        consumerInfo.houseFIASid = houseFIASid_guid;
                    }
                }
                catch (InvalidOperationException)
                {

                }
                splashScreenManager1.SetWaitFormDescription("Обработка данных МКД (" + (++i).ToString() + " из " + dt.Rows.Count.ToString() + ")");
            }
            rtp3eshEntities.SaveChanges();
            //----------------------МКД

            dt.Clear();
            dt.Columns.Clear();

            // запрос для ЮЛ ----------ЮЛ
            string queryStringYL = String.Concat(
                //"SELECT TOP 1000",
                "SELECT",
                " _Marked,", /*ПометкаУдаления*/
                "_Code,",   /*Код*/
                "_Fld1454,",    /*КодИЭСБК*/
                "_Fld1455,",    /*КодОКЭ*/
                "_Fld1461,",    /*ПредставлениеАдреса*/
                "_Fld6662,",    /*ТелефонДляУведомлений*/
                "_Fld6663,",    /*ЭлектроннаяПочтаДляУведомлений*/
                "_Fld6932",	/*ДомФИАСID*/
                " FROM _Reference83",
                " WHERE _Marked = 0"
            );            
            adapter.SelectCommand = new SqlCommand(queryStringYL, connection);
            adapter.SelectCommand.CommandTimeout = 0;
            adapter.Fill(dt);
            i = 0;
            foreach (DataRow dr in dt.Rows)
            {                
                try
                {
                    string codeOKE = dr["_Fld1455"].ToString();
                    ConsumersInfo consumerInfo = consumersInfo_List.First(p => String.Equals(p.ContractRTP3, codeOKE)); // ищем по коду ОКЭ
                    
                    consumerInfo.codeIESBK = dr["_Fld1454"].ToString();
                    consumerInfo.codeOKE = codeOKE;
                    consumerInfo.addressInfo = dr["_Fld1461"].ToString();
                    string phoneInfo = dr["_Fld6662"].ToString();
                    if (phoneInfo.Equals("0") || String.IsNullOrEmpty(phoneInfo)) consumerInfo.phoneInfo = null;
                    else consumerInfo.phoneInfo = phoneInfo;
                    string emailInfo = dr["_Fld6663"].ToString();
                    if (emailInfo.Equals("0") || String.IsNullOrEmpty(emailInfo)) consumerInfo.emailInfo = null;
                    else consumerInfo.emailInfo = emailInfo;

                    if (!String.IsNullOrEmpty(dr["_Fld6932"].ToString()))
                    {
                        byte[] houseFIASid_bytes_bad = (byte[])dr["_Fld6932"];
                        byte[] houseFIASid_bytes_good =
                            new byte[]
                            {
                            houseFIASid_bytes_bad[15], houseFIASid_bytes_bad[14], houseFIASid_bytes_bad[13], houseFIASid_bytes_bad[12],
                            houseFIASid_bytes_bad[11], houseFIASid_bytes_bad[10], houseFIASid_bytes_bad[9], houseFIASid_bytes_bad[8],
                            houseFIASid_bytes_bad[0], houseFIASid_bytes_bad[1], houseFIASid_bytes_bad[2], houseFIASid_bytes_bad[3],
                            houseFIASid_bytes_bad[4], houseFIASid_bytes_bad[5], houseFIASid_bytes_bad[6], houseFIASid_bytes_bad[7]
                            };
                        Guid houseFIASid_guid = new Guid(houseFIASid_bytes_good);
                        consumerInfo.houseFIASid = houseFIASid_guid;
                    }
                }
                catch (InvalidOperationException)
                {

                }
                splashScreenManager1.SetWaitFormDescription("Обработка данных ЮЛ (" + (++i).ToString() + " из " + dt.Rows.Count.ToString() + ")");
            }
            rtp3eshEntities.SaveChanges();
            //----------------------ЮЛ

            connection.Close();

            splashScreenManager1.SetWaitFormDescription("завершение...");
            splashScreenManager1.CloseWaitForm();
            MessageBox.Show(text: "Загрузка успешно завершена!", caption: "Информация");
        } // private void barButtonItemLoadDataFrom1C_ItemClick(object sender, ItemClickEventArgs e)
    }
}
