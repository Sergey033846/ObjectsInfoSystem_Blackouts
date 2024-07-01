using rtp3esh_bd;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraEditors;
using DevExpress.Spreadsheet;

namespace ObjectsInfoSystem
{
    public partial class FormLoadData : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        public FormLoadData()
        {
            InitializeComponent();
        }

        private void FormLoadData_Load(object sender, EventArgs e)
        {

        }

        // "выбрать вид загрузки"
        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormLoadTypeSelect form1 = null;
            form1 = new FormLoadTypeSelect();
            form1.Text = "Выберите вид загрузки данных";

            // загрузка названий таблиц
            form1.comboBoxEdit1.Properties.Items.Clear();

            // все что в комментах снежинка слэш убрать
            DataSetPnrmMapSrc DataSetLoad = new DataSetPnrmMapSrc();
            DataSetIESBK DataSetIESBKLoad = new DataSetIESBK();
            //DataSet1TableAdapters.tblMeteringPointTableAdapter tblMeteringPointTableAdapter = new DataSet1TableAdapters.tblMeteringPointTableAdapter();
            //DataSet1TableAdapters.tblCounterValueTableAdapter tblCounterValueTableAdapter = new DataSet1TableAdapters.tblCounterValueTableAdapter();
            
            // инициализация значений подставляемых свойств
            for (int i = 0; i < DataSetLoad.Tables.Count; i++)
                form1.comboBoxEdit1.Properties.Items.Add(DataSetLoad.Tables[i].TableName.ToString());

            // ПАНОРАМА полная
            form1.comboBoxEdit7.Properties.Items.Clear();
            form1.comboBoxEdit6.Properties.Items.Clear();
            form1.textEdit1.Text = "";
            form1.textEdit2.Text = "";
            form1.textEdit3.Text = "";
            DataSetPnrmMapSrcTableAdapters.tblContractorTableAdapter tblContractorTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblContractorTableAdapter();
            DataSetPnrmMapSrcTableAdapters.tblFilialTableAdapter tblFilialTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblFilialTableAdapter();
            DataSetPnrmMapSrcTableAdapters.tblMapSourceTableAdapter tblMapSourceTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblMapSourceTableAdapter();
            tblContractorTableAdapter.Fill(DataSetLoad.tblContractor);
            tblFilialTableAdapter.Fill(DataSetLoad.tblFilial);
            tblMapSourceTableAdapter.Fill(DataSetLoad.tblMapSource);
            for (int i = 0; i < DataSetLoad.tblContractor.Count; i++) form1.comboBoxEdit7.Properties.Items.Add(DataSetLoad.tblContractor.Rows[i]["captioncontr"].ToString());
            for (int i = 0; i < DataSetLoad.tblFilial.Count; i++) form1.comboBoxEdit6.Properties.Items.Add(DataSetLoad.tblFilial.Rows[i]["captionfilial"].ToString());

            // определяем максимальное значение idmapsrc--- пока так без sql            
            string coordsort = "idmapsrc DESC";
            DataRow[] MapSourceRows = DataSetLoad.tblMapSource.Select("", coordsort);
            int pnrmMapSrcmaxvalue = Convert.ToInt32(MapSourceRows[0]["idmapsrc"].ToString());
            //------------------------------------------------
            //form1.textEdit3.Text = (DataSetLoad.tblMapSource.Count+1).ToString(); // инф поле - потом убрать?
            form1.textEdit3.Text = (pnrmMapSrcmaxvalue+1).ToString();
            //---------------------------------------------

            if (form1.ShowDialog(this) == DialogResult.OK)
            {
                splashScreenManager1.ShowWaitForm();

                IWorkbook workbook = this.spreadsheetControl1.Document;
                Worksheet worksheet = workbook.Worksheets[0];

                workbook.History.IsEnabled = false;

                // определяем количество колонок в листе
                int maxcol = 0;
                Cell spcell = worksheet[0, 0];

                // определяем количество колонок в листе (считаем по 1-ой строке)
                while (!String.IsNullOrWhiteSpace(spcell.Value.ToString()))
                {
                    maxcol++;
                    spcell = worksheet[0, maxcol];
                }

                // определяем количество строк в листе (считаем по 1-му столбцу)
                int maxrow = 0;                
                spcell = worksheet[0, 0];
                while (!String.IsNullOrWhiteSpace(spcell.Value.ToString()))
                {
                    maxrow++;
                    spcell = worksheet[maxrow, 0];
                }
                //-------------------------------

                this.spreadsheetControl1.BeginUpdate();
                this.spreadsheetControl1.Hide();

                string[] svalues = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

                // загружаем данные в зависимости от вида загрузки
                switch (form1.listBoxLoadType.SelectedValue.ToString())
                {
                    case "координаты по адресу": // без внесения инфы в БД
                        //tblMeteringPointTableAdapter.Fill(DataSetLoad.tblMeteringPoint);

                        spreadsheetControl1.Hide();

                        // начинаем с первой строки
                        for (int i = 1; i < maxrow; i++)
                        {
                            //for (int j = 0; j < maxcol; j++)
                            for (int j = 0; j < 3; j++)
                                {
                                spcell = worksheet[i, j];
                                svalues[j] = spcell.Value.ToString();
                            }

                            /*DataRow mpRow = DataSetLoad.tblMeteringPoint.FindByid_counter_serial(svalues[3]);
                            if (mpRow == null)
                            {
                                worksheet.Rows[i].FillColor = Color.Red;
                            }
                            else
                            {*/
                                // адресная строка запроса для геокодирования (первый столбец)
                                //string addr_find = svalues[0]+","+ svalues[1]+","+ svalues[2];
                                string addr_find = "Иркутская область ,"+svalues[0];

                            //Запрос к API геокодирования Google ----------------------------------------------------------
                            // !!!!!!!!!!!!!!! бесплатно 2500 в день
                            string url = string.Format(
                                    "http://maps.googleapis.com/maps/api/geocode/xml?address={0}&sensor=true_or_false&language=ru",
                                    Uri.EscapeDataString(addr_find));

                                //Выполняем запрос к универсальному коду ресурса (URI).
                                System.Net.HttpWebRequest request = (System.Net.HttpWebRequest)System.Net.WebRequest.Create(url);

                                //Получаем ответ от интернет-ресурса.
                                System.Net.WebResponse response = request.GetResponse();

                                //Экземпляр класса System.IO.Stream
                                //для чтения данных из интернет-ресурса.
                                System.IO.Stream dataStream = response.GetResponseStream();

                                //Инициализируем новый экземпляр класса
                                //System.IO.StreamReader для указанного потока.
                                System.IO.StreamReader sreader = new System.IO.StreamReader(dataStream);

                                //Считывает поток от текущего положения до конца.           
                                string responsereader = sreader.ReadToEnd();

                                //Закрываем поток ответа.
                                response.Close();

                                System.Xml.XmlDocument xmldoc = new System.Xml.XmlDocument();

                                xmldoc.LoadXml(responsereader);

                                //Переменные широты и долготы.
                                double latitude = 0.0;
                                double longitude = 0.0;

                                if (xmldoc.GetElementsByTagName("status")[0].ChildNodes[0].InnerText == "OK")
                                {

                                    //Получение широты и долготы.
                                    System.Xml.XmlNodeList nodes = xmldoc.SelectNodes("//location");
                                                                        
                                    //Получаем широту и долготу.
                                    foreach (System.Xml.XmlNode node in nodes)
                                    {
                                        latitude = System.Xml.XmlConvert.ToDouble(node.SelectSingleNode("lat").InnerText.ToString());
                                        longitude = System.Xml.XmlConvert.ToDouble(node.SelectSingleNode("lng").InnerText.ToString());
                                    }
                                }

                                //----------------------------------

                                /*mpRow["geo1coord"] = latitude.ToString();
                                mpRow["geo2coord"] = longitude.ToString();
                                tblMeteringPointTableAdapter.Update(mpRow);*/

                                worksheet[i, 1].SetValue(latitude.ToString());
                                worksheet[i, 2].SetValue(longitude.ToString());                                                                
                           // }

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }

                        spreadsheetControl1.Show();

                        break; // case "координаты по адресу":

                    case "ПАНОРАМА полная":
                        // СДЕЛАТЬ ПРОВЕРКУ НА ОТСУТСТВИЕ ПУСТЫХ ПОЛЕЙ
                        //---- добавляем mapsource-----
                        //int idmapsrc = DataSetLoad.tblMapSource.Count + 1; было раньше
                        int idmapsrc = Convert.ToInt32(form1.textEdit3.Text);
                        string captionmapsrc = form1.textEdit1.Text;
                        string commentmapsrc = form1.textEdit2.Text;

                        DataRow[] findrows = DataSetLoad.tblContractor.Select("captioncontr = '" + form1.comboBoxEdit7.SelectedItem.ToString() + "'");                        
                        int idcontractor = Convert.ToInt32(findrows[0]["idcontractor"].ToString());

                        findrows = DataSetLoad.tblFilial.Select("captionfilial = '"+ form1.comboBoxEdit6.SelectedItem.ToString()+"'");
                        int idfilial = Convert.ToInt32(findrows[0]["idfilial"].ToString());

                        tblMapSourceTableAdapter.Insert(idmapsrc, captionmapsrc, commentmapsrc, idcontractor, idfilial, null,DateTime.Now);
                        tblMapSourceTableAdapter.Update(DataSetLoad.tblMapSource);
                        //-----------------------------
                        DataSetPnrmMapSrcTableAdapters.tblPanoramaObjectTableAdapter tblPanoramaObjectTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaObjectTableAdapter();
                        DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();
                        DataSetPnrmMapSrcTableAdapters.tblPanoramaSemanticTableAdapter tblPanoramaSemanticTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaSemanticTableAdapter();
                        DataSetPnrmMapSrcTableAdapters.tblPanoramaObjSemValuesTableAdapter tblPanoramaObjSemValuesTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaObjSemValuesTableAdapter();

                        spreadsheetControl1.Hide();

                        int first_notsys_column = form1.checkEditFromDbf.Checked ? 13 : 15;
                        int max_sys_column = 15;
                        int back_delta_column = form1.checkEditFromDbf.Checked ? 2 : 0;

                        // 0-ая строка содержит заголовки полей, коды семантики
                        string[] columnnames = new string[maxcol];
                        for (int j = 0; j < maxcol; j++) columnnames[j] = worksheet[0, j].Value.ToString();

                        string[] pvalues = new string[maxcol];

                        bool is_file_correct = false;
                        if (form1.checkEditFromDbf.Checked) // проверяем на полноту файла, если нет, то не загружаем
                        {
                            // если присутствуют все колонки
                            if (columnnames[0] == "N" && columnnames[1] == "OBJECT" && columnnames[2] == "SUBJECT" &&
                                columnnames[3] == "POINT" && columnnames[4] == "NAME" && columnnames[5] == "OBJECTKEY" &&
                                columnnames[6] == "LAYER" && columnnames[7] == "LOCAL" && columnnames[8] == "LENGHT" &&
                                columnnames[9] == "SQUARE" && columnnames[10] == "X" && columnnames[11] == "Y" &&
                                columnnames[12] == "H" && columnnames[maxcol-2] == "LINKSHEET" && columnnames[maxcol-1] == "LINKOBJECT") 
                            {
                                is_file_correct = true;
                            }
                        }

                        if (is_file_correct)
                        {
                            // добавляем новые семантики при необходимости
                            tblPanoramaSemanticTableAdapter.Fill(DataSetLoad.tblPanoramaSemantic);
                            for (int j = first_notsys_column; j < maxcol - back_delta_column; j++)
                            {
                                DataRow semrow = DataSetLoad.tblPanoramaSemantic.FindByidpnrmSEMidmapsrc(columnnames[j], idmapsrc);
                                if (semrow == null)
                                {
                                    tblPanoramaSemanticTableAdapter.Insert(columnnames[j], null, idmapsrc);
                                    tblPanoramaSemanticTableAdapter.Update(DataSetLoad.tblPanoramaSemantic);
                                }
                            }
                            //--------------------------------------------

                            // инициализация номеров колонок (зависит от "dbf или нет")
                            int[] colid = new int[max_sys_column];
                            if (form1.checkEditFromDbf.Checked)
                            {
                                colid[0] = 0; colid[1] = 1; colid[2] = 2; colid[3] = 3; colid[4] = 10; colid[5] = 11; colid[6] = 12;
                                colid[7] = 4; colid[8] = 5; colid[9] = 6; colid[10] = 7; colid[11] = 8; colid[12] = 9;
                                colid[13] = maxcol - 1; colid[14] = maxcol - 2;
                            }
                            else
                            {
                                for (int i = 0; i < max_sys_column; i++) colid[i] = i;
                            }
                            //---------------------------------------------------------

                            for (int i = 1; i < maxrow; i++)
                            {
                                for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                                int idpnrmOBJECT = Convert.ToInt32(pvalues[colid[1]]);
                                int pnrmSUBJECT = Convert.ToInt32(pvalues[colid[2]]);
                                string pnrmNAME = String.IsNullOrWhiteSpace(pvalues[colid[7]]) ? null : pvalues[colid[7]];
                                string pnrmOBJECTKEY = String.IsNullOrWhiteSpace(pvalues[colid[8]]) ? null : pvalues[colid[8]];
                                string pnrmLAYER = String.IsNullOrWhiteSpace(pvalues[colid[9]]) ? null : pvalues[colid[9]];
                                string pnrmLOCAL = String.IsNullOrWhiteSpace(pvalues[colid[10]]) ? null : pvalues[colid[10]];

                                // преобразовываем из вормата Панорамы 123.45 в 123,45 LENGTH и SQUARE (из dbf не требуется)
                                string pnrmLENGTHstr = String.IsNullOrWhiteSpace(pvalues[colid[11]]) ? null : pvalues[colid[11]];
                                //double? pnrmLENGTH = null;  if (pnrmLENGTHstr != null) pnrmLENGTH = Convert.ToDouble(pnrmLENGTHstr.Replace('.', ','));
                                double? pnrmLENGTH = null; if (pnrmLENGTHstr != null) pnrmLENGTH = Convert.ToDouble(pnrmLENGTHstr);

                                string pnrmSQUAREstr = String.IsNullOrWhiteSpace(pvalues[colid[12]]) ? null : pvalues[colid[12]];
                                //double? pnrmSQUARE = null; if (pnrmSQUAREstr != null) pnrmSQUARE = Convert.ToDouble(pnrmSQUAREstr.Replace('.', ','));
                                double? pnrmSQUARE = null; if (pnrmSQUAREstr != null) pnrmSQUARE = Convert.ToDouble(pnrmSQUAREstr);
                                //--------------------------------------------------------------------

                                int? pnrmLINKOBJECT = null; if (!String.IsNullOrWhiteSpace(pvalues[colid[13]])) pnrmLINKOBJECT = Convert.ToInt32(pvalues[colid[13]]);
                                string pnrmLINKSHEET = String.IsNullOrWhiteSpace(pvalues[colid[14]]) ? null : pvalues[colid[14]];

                                int pnrmPOINT = Convert.ToInt32(pvalues[colid[3]]);
                                int idpnrmN = Convert.ToInt32(pvalues[colid[0]]);
                                string pnrmX = String.IsNullOrWhiteSpace(pvalues[colid[4]]) ? null : pvalues[colid[4]];
                                string pnrmY = String.IsNullOrWhiteSpace(pvalues[colid[5]]) ? null : pvalues[colid[5]];
                                string pnrmH = String.IsNullOrWhiteSpace(pvalues[colid[6]]) ? null : pvalues[colid[6]];
                                string coordALT = null;
                                string coordLONG = null;

                                int? idpnrmLOCALtype = null; // установку значений делаю через SQL Server
                                int? idpnrmGROUPLAYER = null; // установку значений делаю через SQL Server
                                //----------------------------------

                                // оптимизируем добавление -----
                                if (pnrmLOCAL != null)
                                {
                                    /*tblPanoramaObjectTableAdapter.Fill(DataSetLoad.tblPanoramaObject);
                                    DataRow objrow = DataSetLoad.tblPanoramaObject.FindByidpnrmOBJECTidmapsrc(idpnrmOBJECT, idmapsrc);
                                    if (objrow == null)
                                    {*/
                                    tblPanoramaObjectTableAdapter.Insert(idpnrmOBJECT, idmapsrc, pnrmNAME, pnrmOBJECTKEY, pnrmLAYER,
                                        pnrmLOCAL, pnrmLENGTH, pnrmSQUARE, pnrmLINKOBJECT, pnrmLINKSHEET, idpnrmLOCALtype, idpnrmGROUPLAYER);
                                    //tblPanoramaObjectTableAdapter.Update(DataSetLoad.tblPanoramaObject);

                                    // добавляем свойства семантики (с 15-й колонки)                            
                                    for (int j = first_notsys_column; j < maxcol - back_delta_column; j++)
                                    {
                                        string idpnrmSEM = columnnames[j];
                                        string pnrmSEMvalue = String.IsNullOrWhiteSpace(pvalues[j]) ? null : pvalues[j];
                                        if (!String.IsNullOrWhiteSpace(pnrmSEMvalue))
                                        {
                                            tblPanoramaObjSemValuesTableAdapter.Insert(idpnrmSEM, idmapsrc, idpnrmOBJECT, pnrmSEMvalue);
                                            //tblPanoramaObjSemValuesTableAdapter.Update(DataSetLoad.tblPanoramaObjSemValues);
                                        }
                                    }
                                    // }
                                }
                                //---------------------------------
                                tblPanoramaCoordsTableAdapter.Insert(idpnrmOBJECT, idpnrmN, idmapsrc, pnrmPOINT, pnrmX, pnrmY, pnrmH, coordALT, coordLONG, pnrmSUBJECT, null, null);
                                //tblPanoramaCoordsTableAdapter.Update(DataSetLoad.tblPanoramaCoords);

                                //--------------------

                                worksheet[i, 0].FillColor = Color.Green;
                                //----------------------------------

                                splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                            } // for (int i = 1; i < maxrow; i++)

                            tblPanoramaObjectTableAdapter.Update(DataSetLoad.tblPanoramaObject);
                            tblPanoramaCoordsTableAdapter.Update(DataSetLoad.tblPanoramaCoords);
                            tblPanoramaObjSemValuesTableAdapter.Update(DataSetLoad.tblPanoramaObjSemValues);
                        } // if (is_file_correct)
                        else
                        {
                            MessageBox.Show("Некорректный файл!");
                        }

                        spreadsheetControl1.Show();

                        break; // case "ПАНОРАМА полная":

                        //----------------------------------------------------------------

                    case "ПАНОРАМА WGS84":
                        tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();
                        tblPanoramaCoordsTableAdapter.Fill(DataSetLoad.tblPanoramaCoords);

                        spreadsheetControl1.Hide();                        
                        
                        for (int i = 1; i < maxrow; i++)
                        {
                            pvalues = new string[maxcol];
                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                            int idpnrmOBJECT = Convert.ToInt32(pvalues[2]);
                            int pnrmSUBJECT = Convert.ToInt32(pvalues[3]);
                            int pnrmPOINT = Convert.ToInt32(pvalues[1]);
                            idmapsrc = Convert.ToInt32(pvalues[0]);
                            string pnrmX = pvalues[4];
                            string pnrmY = pvalues[5];
                            string[] coordsWGS84 = pvalues[7].Split(',');
                            string coordALTstr = coordsWGS84[0];
                            //string coordALT = coordALTstr.Replace('.', ',');
                            double coordALT = Convert.ToDouble(coordALTstr.Replace('.', ','));
                            string coordLONGstr = coordsWGS84[1].Substring(1);
                            //string coordLONG = coordLONGstr.Replace('.', ',');
                            double coordLONG = Convert.ToDouble(coordLONGstr.Replace('.', ','));

                            DataRow coordrow = DataSetLoad.tblPanoramaCoords.FindByidpnrmOBJECTidmapsrcpnrmPOINTpnrmSUBJECT(idpnrmOBJECT, idmapsrc, pnrmPOINT, pnrmSUBJECT);
                            if (coordrow != null)
                            {
                                coordrow["coordALT"] = coordALT;
                                coordrow["coordLONG"] = coordLONG;
                                tblPanoramaCoordsTableAdapter.Update(coordrow);
                                worksheet[i,0].FillColor = Color.Green;
                            }
                            
                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        } // for (int i = 1; i < maxrow; i++)                        


                        /*
                        // делаем через поиск X,Y
                        tblPanoramaCoordsTableAdapter = new DataSet1TableAdapters.tblPanoramaCoordsTableAdapter();
                        tblPanoramaCoordsTableAdapter.Fill(DataSetLoad.tblPanoramaCoords);

                        spreadsheetControl1.Hide();

                        pvalues = new string[maxcol];

                        for (int i = 1; i < maxrow; i++)
                        {                                                        
                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();
                                                        
                            string pnrmX = pvalues[4];
                            string pnrmY = pvalues[5];

                            string[] coordsWGS84 = pvalues[7].Split(',');

                            string coordALTstr = coordsWGS84[0];                            
                            string coordALT = coordALTstr.Replace('.', ',');

                            string coordLONGstr = coordsWGS84[1].Substring(1);                            
                            string coordLONG = coordLONGstr.Replace('.', ',');

                            string fexpr = "pnrmX = '" + pnrmX + "' AND pnrmY = '" + pnrmY + "' AND coordALT is NULL";
                            DataRow[] coordrows = DataSetLoad.tblPanoramaCoords.Select(fexpr);

                            foreach (DataRow coordrow in coordrows)
                            {
                                coordrow["coordALT"] = coordALT;
                                coordrow["coordLONG"] = coordLONG;
                                tblPanoramaCoordsTableAdapter.Update(coordrow);
                                //worksheet.Rows[i].FillColor = Color.Green;
                            }

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        } // for (int i = 1; i < maxrow; i++)*/

                        //----------------------------------------------------------

                        //tblPanoramaCoordsTableAdapter.Update(DataSetLoad.tblPanoramaCoords); // непрозрачно, в конце долго ждать

                        spreadsheetControl1.Show();

                        break; // case "ПАНОРАМА WGS84":

                    //----------------------------------------------------------------

                    case "ПАНОРАМА WGS84 цветами":
                        tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();
                        tblPanoramaCoordsTableAdapter.Fill(DataSetLoad.tblPanoramaCoords);

                        spreadsheetControl1.Hide();

                        int maxcoordrowsinwrksh = 8000;
                        int coordrowsincolumn = 2000; // текущая виртуальная "колонка координат" со служебными полями
                        int maxcolumninwrksh = maxcoordrowsinwrksh / coordrowsincolumn;

                        bool enddataflag = false;

                        for (int wrksh = 0; wrksh < workbook.Worksheets.Count; wrksh++)
                        {
                            Worksheet worksheet2 = workbook.Worksheets[wrksh];

                            for (int col = 0; col < maxcolumninwrksh; col++)
                            {
                                if (enddataflag) break;

                                for (int i = 0; i < coordrowsincolumn; i++)
                                {
                                    // читаем строку значений
                                    pvalues = new string[8];
                                    for (int j = 0; j < 8; j++) pvalues[j] = worksheet2[i, col*8+j].Value.ToString();

                                    if (String.IsNullOrWhiteSpace(pvalues[0]))
                                    {
                                        enddataflag = true;
                                        break;
                                    }

                                    int idpnrmOBJECT = Convert.ToInt32(pvalues[2]);
                                    int pnrmSUBJECT = Convert.ToInt32(pvalues[3]);
                                    int pnrmPOINT = Convert.ToInt32(pvalues[1]);
                                    idmapsrc = Convert.ToInt32(pvalues[0]);
                                    string pnrmX = pvalues[4];
                                    string pnrmY = pvalues[5];
                                    string[] coordsWGS84 = pvalues[7].Split(',');
                                    string coordALTstr = coordsWGS84[0];
                                    //string coordALT = coordALTstr.Replace('.', ',');
                                    double coordALT = Convert.ToDouble(coordALTstr.Replace('.', ','));
                                    string coordLONGstr = coordsWGS84[1].Substring(1);
                                    //string coordLONG = coordLONGstr.Replace('.', ',');
                                    double coordLONG = Convert.ToDouble(coordLONGstr.Replace('.', ','));

                                    DataRow coordrow = DataSetLoad.tblPanoramaCoords.FindByidpnrmOBJECTidmapsrcpnrmPOINTpnrmSUBJECT(idpnrmOBJECT, idmapsrc, pnrmPOINT, pnrmSUBJECT);
                                    if (coordrow != null)
                                    {
                                        coordrow["coordALT"] = coordALT;
                                        coordrow["coordLONG"] = coordLONG;

                                        coordrow["coordALTfloat"] = coordALT;
                                        coordrow["coordLONGfloat"] = coordLONG;

                                        tblPanoramaCoordsTableAdapter.Update(coordrow);
                                        //worksheet2[i, col*8].FillColor = Color.Green;
                                    }

                                    splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (wrksh+1).ToString() + "(" + workbook.Worksheets.Count.ToString() + ")-" + col.ToString() + "-" + i.ToString()+ ")");
                                } // for (int i = 1; i < maxrow; i++)
                            } // for (int col = 0; col < maxcolumninwrksh; col++)
                        } // for (int wrksh = 0; wrksh < workbook.Worksheets.Count; wrksh++)

                            spreadsheetControl1.Show();

                        break; // case "ПАНОРАМА WGS84 цветами":

                        //---------------------------------------------------------

                    case "ПАНОРАМА точка":
                        tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();
                        tblPanoramaCoordsTableAdapter.Fill(DataSetLoad.tblPanoramaCoords);

                        for (int i = 0; i < DataSetLoad.tblPanoramaCoords.Rows.Count; i++)
                        {
                            string coordALT = DataSetLoad.tblPanoramaCoords.Rows[i]["coordALT"].ToString().Replace('.',',');
                            string coordLONG = DataSetLoad.tblPanoramaCoords.Rows[i]["coordLONG"].ToString().Replace('.', ',');

                            DataSetLoad.tblPanoramaCoords.Rows[i]["coordALT"] = coordALT;
                            DataSetLoad.tblPanoramaCoords.Rows[i]["coordLONG"] = coordLONG;
                            tblPanoramaCoordsTableAdapter.Update(DataSetLoad.tblPanoramaCoords.Rows[i]);                                

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }                       

                        break; // case "ПАНОРАМА точка":

                        //-----------------------------------------------------------

                    case "ПАНОРАМА семантика": // загрузка описаний значений
                        tblPanoramaSemanticTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaSemanticTableAdapter();
                        tblPanoramaSemanticTableAdapter.Fill(DataSetLoad.tblPanoramaSemantic);

                        tblMapSourceTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblMapSourceTableAdapter();
                        tblMapSourceTableAdapter.Fill(DataSetLoad.tblMapSource);

                        spreadsheetControl1.Hide();
                        
                        /*pvalues = new string[maxcol];

                        for (int i = 1; i < maxrow; i++)
                        {                            
                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();
                            
                            string idpnrmSEM = pvalues[0];
                            string pnrmSEMcaption = pvalues[1];
                            idmapsrc = Convert.ToInt32(pvalues[2]);

                            DataRow findrow = DataSetLoad.tblPanoramaSemantic.FindByidpnrmSEMidmapsrc(idpnrmSEM,idmapsrc);
                            if (findrow != null)
                            {
                                findrow["pnrmSEMcaption"] = pnrmSEMcaption;
                                tblPanoramaSemanticTableAdapter.Update(findrow);
                                worksheet.Rows[i].FillColor = Color.Green;
                            }

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }*/

                        pvalues = new string[maxcol];

                        //int idmapsrcCOUNT = DataSetLoad.tblMapSource.Rows.Count;
                        //int idmapsrcCOUNT = 123; // Аланс
                        /*int[] idmapsrcMAS = { 1,2,3,4,5,6,7,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,
                                            32,33,34,35,36,37,38,39,40,45,46,47,48,49,50,51,52,53,54,55,56,57,59,60,
                                            115,116,117,118,119,120,121,122,123}; // Аланс*/

                        //for (int k = 1; k <= idmapsrcCOUNT; k++)

                        foreach (DataRow mapSrcRow in DataSetLoad.tblMapSource.Rows)
                        {
                            idmapsrc = (int)mapSrcRow["idmapsrc"];

                            DataRow[] semRows = DataSetLoad.tblPanoramaSemantic.Select("idmapsrc = " + idmapsrc.ToString());

                            for (int i = 1; i < maxrow; i++)
                            {
                                for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                                string idpnrmSEM = pvalues[0];
                                string pnrmSEMcaption = pvalues[1];
                                //idmapsrc = k;
                                
                                DataRow findrow = DataSetLoad.tblPanoramaSemantic.FindByidpnrmSEMidmapsrc(idpnrmSEM, idmapsrc);
                                if (findrow != null)
                                {
                                    findrow["pnrmSEMcaption"] = pnrmSEMcaption;
                                    tblPanoramaSemanticTableAdapter.Update(findrow);
                                    //worksheet.Rows[i].FillColor = Color.Green;
                                }

                                splashScreenManager1.SetWaitFormDescription("Обработка данных - " + (idmapsrc).ToString());
                            }
                        }

                        spreadsheetControl1.Show();

                        break; // case "ПАНОРАМА семантика":

                    //-----------------------------------------------------------

                    case "ИЭСБК Группы и свойства ФЛ": // загрузка видов макетов и свойств лицевых счетов ФЛ ИЭСБК
                        DataSetIESBKTableAdapters.tblIESBKtemplateTableAdapter tblIESBKtemplateTableAdapter = new DataSetIESBKTableAdapters.tblIESBKtemplateTableAdapter();
                        DataSetIESBKTableAdapters.tblIESBKlspropTableAdapter tblIESBKlspropTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlspropTableAdapter();
                        
                        spreadsheetControl1.Hide();

                        pvalues = new string[maxcol];

                        /*tblIESBKtemplateTableAdapter.Insert("служебный","служебный");
                        tblIESBKtemplateTableAdapter.Update(DataSetIESBKLoad.tblIESBKtemplate);*/

                        for (int i = 1; i < maxrow; i++)
                        {
                            tblIESBKtemplateTableAdapter.Fill(DataSetIESBKLoad.tblIESBKtemplate);
                            tblIESBKlspropTableAdapter.Fill(DataSetIESBKLoad.tblIESBKlsprop);

                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                            DataRow findrow = DataSetIESBKLoad.tblIESBKtemplate.FindBytemplateid(pvalues[0]);
                            if (findrow == null)
                            {
                                tblIESBKtemplateTableAdapter.Insert(pvalues[0],pvalues[0]);
                                tblIESBKtemplateTableAdapter.Update(DataSetIESBKLoad.tblIESBKtemplate);
                            }

                            int? numcolumninfile = null;
                            //string lspropertieid = String.Concat(pvalues[0],"_",pvalues[2]);
                            tblIESBKlspropTableAdapter.Insert(pvalues[2], pvalues[0],numcolumninfile,pvalues[1]);
                            tblIESBKlspropTableAdapter.Update(DataSetIESBKLoad.tblIESBKlsprop);
                            
                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }

                        spreadsheetControl1.Show();

                        break; // case "ИЭСБК Группы и свойства ФЛ":

                    //-----------------------------------------------------------

                    case "ИЭСБК Объемы ФЛ": // загрузка объема потребления ФЛ ИЭСБК
                        // гружу с 3-ей строки, поэтому поменял алгоритма расчета maxrow                        
                        // определяем количество строк в листе (считаем по 1-му столбцу)
                        maxrow = 0+2; // !!!
                        spcell = worksheet[2, 0]; // !!!
                        while (!String.IsNullOrWhiteSpace(spcell.Value.ToString()))
                        {
                            maxrow++;
                            spcell = worksheet[maxrow, 0];
                        }
                        //-------------------------------

                        tblIESBKtemplateTableAdapter = new DataSetIESBKTableAdapters.tblIESBKtemplateTableAdapter();
                        tblIESBKlspropTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlspropTableAdapter();
                        tblIESBKtemplateTableAdapter.Fill(DataSetIESBKLoad.tblIESBKtemplate);
                        tblIESBKlspropTableAdapter.Fill(DataSetIESBKLoad.tblIESBKlsprop);

                        DataSetIESBKTableAdapters.tblIESBKlsTableAdapter tblIESBKlsTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlsTableAdapter();
                        DataSetIESBKTableAdapters.tblIESBKlspropvalueTableAdapter tblIESBKlspropvalueTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlspropvalueTableAdapter();
                        tblIESBKlsTableAdapter.Fill(DataSetIESBKLoad.tblIESBKls);
                        //tblIESBKlspropvalueTableAdapter.Fill(DataSetIESBKLoad.tblIESBKlspropvalue);

                        spreadsheetControl1.Hide();
                        
                        //int maxcolIESBK = 44; // макет 01, 02, 2016
                        //int maxcolIESBK = 36; // макет 11, 12, 2015
                        //int maxcolIESBK = 18; // макет 10, 2015
                        //int maxcolIESBK = 17; // макет 01-06, 2015

                        int maxcolIESBK = 0;
                        //int maxcolIESBK = 49; // макет 03, 2016
                        //int maxcolIESBK = 50; // макет 08, 2016 (по повышающему коэффициенту)

                        int[] ncolpropmas1;
                        
                        if (form1.checkEditIsPovKoef.Checked) // если установлен флаг "повышающий коэффициент"
                        {
                            maxcolIESBK = 50; // макет 08, 2016 (по повышающему коэффициенту)

                            // макет 08, 2016 - 50 ("по повышающему коэффициенту")
                            ncolpropmas1 = new int[] { 1,2,3,4,5,6,7,8,9,10,50,11,12,13,14,15,16,17,18,19,51,52,20,21,22,23,24,
                                                25,26,27,28,55,29,30,53,31,54,32,33,34,35,36,37,38,39,40,41,42,43,44 };
                        }
                        else
                        {
                            maxcolIESBK = 49; // макет 03, 2016

                            // макет 03, 2016 - 49
                            ncolpropmas1 = new int[] { 1,2,3,4,5,6,7,8,9,10,50,11,12,13,14,15,16,17,18,19,51,52,20,21,22,23,24,
                                                25,26,27,28,29,30,53,31,54,32,33,34,35,36,37,38,39,40,41,42,43,44 };
                        }

                        pvalues = new string[maxcolIESBK];

                        // макет 01, 02, 2016 - 44 свойства
                        string[] propcaptionmas1 = { "№ п/п", "Лицевой счет", "ЛСвАСУСЭиРП", "ЛС ОКЭ / БЭСК", "ФИО", "Тип ПУ", "Заводской номер", "КТ", "Фазность ПУ", "Класс точности",
                                                    "Район", "Населенный пункт", "Улица", "Дом", "Номер квартиры", "Количество проживающих", "Количество комнат", "Тип плиты",
                                                    "Дата заключения договора", "Дата  снятия контрольного показания ", "Дата предыдущего показания", "Предыдущее показание прибора учета",
                                                    "Вид предыдущего показания", "Дата последнего показания", "Показание прибора учета", "Вид последнего показания", "Полезный отпуск, кВт*ч",
                                                    "Без ИПУ по базовому нормативу", "с ИПУ, Расход по прибору", "с ИПУ, Расход по среднемесячному (нет переданных показаний 6 месяцев)",
                                                    "с ИПУ, Расход по нормативу (нет переданных показаний свыше 6 месяцев)", "Корректировки, Ручные", "Корректировки, Прочие",
                                                    "Расход по актам безучетного потребления", "Расход ОДН по ОДПУ", "Расход ОДН по нормативу", "Сетевое Предприятие",
                                                    "РЭС", "Подстанция", "Линия", "Код линии", "КТП", "Трансформатор", "Линия04",
                                                    
                                                    // макет 11, 12, 2015 - +3 = 47 свойств
                                                    "По среднему (макет 11, 12, 2015)", "По договорным нагрузкам (макет 11, 12, 2015)", "Остальные начисления (макет 11, 12, 2015)",

                                                    // макет 10 2015 - +2 = 49 свойств
                                                    "Отделение ОКЭ (или код отделения) (макет 10 2015)", "Начисление ОДН, кВт*ч (макет 10 2015)",

                                                    // макет 03 2016 - +5 = 54 свойств
                                                    "Признак АСКУЭ (макет 03 2016)", "Состояние ЛС (макет 03 2016)", "Тип строения (макет 03 2016)",
                                                    "Кол-во расчетных периодов (среднемесячное) (макет 03 2016)", "Кол-во расчетных периодов (норматив) (макет 03 2016)",

                                                    // макет 08 2016 (только по Ангарскому отделению) - +1 = 55 свойств
                                                    "По повышающему коэффициенту"
                                                    };

                        /*// макет 01, 02, 2016 - 44
                        int[] ncolpropmas1 = { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
                                             25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44 };*/

                        /*// макет 11, 12, 2015 - 36
                        int[] ncolpropmas1 = { 1,2,5,3,4,6,7,9,10,8,16,17,18,12,13,14,15,21,22,24,25,27,28,29,45,31,30,34,32,46,35,36,33,47,23,26 };*/

                        /*// макет 10, 2015 - 18
                        int[] ncolpropmas1 = { 1,48,5,2,3,4,19,21,22,24,25,27,28,29,30,31,49,26 };*/

                        /*// макет 01-06, 2015 - 17
                        int[] ncolpropmas1 = { 1, 48, 5, 2, 3, 4, 19, 21, 22, 24, 25, 27, 28, 29, 30, 31, 49 };*/

                        /*// макет 03, 2016 - 49
                        int[] ncolpropmas1 = { 1,2,3,4,5,6,7,8,9,10,50,11,12,13,14,15,16,17,18,19,51,52,20,21,22,23,24,
                                                25,26,27,28,29,30,53,31,54,32,33,34,35,36,37,38,39,40,41,42,43,44 };*/

                        // макет 08, 2016 - 50 ("по повышающему коэффициенту")
                        /*int[] ncolpropmas1 = { 1,2,3,4,5,6,7,8,9,10,50,11,12,13,14,15,16,17,18,19,51,52,20,21,22,23,24,
                                                25,26,27,28,55,29,30,53,31,54,32,33,34,35,36,37,38,39,40,41,42,43,44 };*/

                        string lstypeid = "ФЛ";                        

                        //!!! поставил с 3-й строки (с нуля для макетов с 03 2016)
                        for (int i = 0; i < maxrow; i++) // пропускаем строку названий (убрал, снова с 0)
                        {                            
                            for (int j = 0; j < maxcolIESBK; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                            string otdelenieIESBK = form1.comboBoxEdit2.SelectedItem.ToString();

                            // грузим 2015, для 2016 pvalues[2]
                            string codeIESBK = pvalues[2].Replace(" ",""); // ЛСвАСУСЭиРП по письму от ИЭСБК !!!!!!!!!!!!!!!!!!!!!!!!!!!

                            DataRow findrow = DataSetIESBKLoad.tblIESBKls.FindBycodeIESBKlstypeidotdelenieid(codeIESBK, lstypeid, otdelenieIESBK);
                            if (findrow == null)
                            {
                                tblIESBKlsTableAdapter.Insert(codeIESBK,lstypeid,otdelenieIESBK, DateTime.Now, 1); 
                                tblIESBKlsTableAdapter.Update(DataSetIESBKLoad.tblIESBKls);
                                //tblIESBKlsTableAdapter.Fill(DataSetIESBKLoad.tblIESBKls); // по Ангарскому отд были дубли
                            }

                            string periodyear = form1.comboBoxEdit3.SelectedItem.ToString();
                            string periodmonth = form1.comboBoxEdit4.SelectedItem.ToString();
                            //string codeIESBK = pvalues[2];

                            string templateid;// = "служебный";
                            //string propcaption = "Отделение ИЭСБК";

                            // отсекаем дубли л/с в файле (вообще, нужна первичная обработка)
                            /*string queryString = "select min(tblPanoramaCoords.coordALT) as Min_coordALT, min(tblPanoramaCoords.coordLONG) as Min_coordLONG," +
                "max(tblPanoramaCoords.coordALT) as Max_coordALT, max(tblPanoramaCoords.coordLONG) as Max_coordLONG " +
                "from tblPanoramaCoords where idmapsrc = " + idmapsrc.ToString();
                            DataTable tableLSPROPVALUES = new DataTable();
                            SelectDataFromSQL(tableLSPROPVALUES, dbconnectionString, queryString);*/

                            /* findrow = DataSetIESBKLoad.tblIESBKlspropvalue.FindByperiodyearperiodmonthcodeIESBKlstypeidlspropcaptiontemplateid(periodyear, periodmonth, codeIESBK, lstypeid, propcaption, templateid);
                             if (findrow == null)
                             {*/

                                /*string propvalue = form1.comboBoxEdit2.SelectedItem.ToString();
                                tblIESBKlspropvalueTableAdapter.Insert(propvalue, periodyear, periodmonth, codeIESBK, lstypeid, propcaption, templateid);
                                //tblIESBKlspropvalueTableAdapter.Update(DataSetIESBKLoad.tblIESBKlspropvalue);*/

                                /*propcaption = "Дата загрузки в БД (системное)";
                                propvalue = DateTime.Now.ToShortDateString();
                                tblIESBKlspropvalueTableAdapter.Insert(propvalue, periodyear, periodmonth, codeIESBK, lstypeid, propcaption, templateid);
                                //tblIESBKlspropvalueTableAdapter.Update(DataSetIESBKLoad.tblIESBKlspropvalue);*/

                                templateid = "3.3.2";
                                for (int j = 0; j < maxcolIESBK; j++)
                                {                                    
                                    //string lspropertieid = propcaptionmas1[ncolpropmas1[j]-1];
                                    string lspropertieid = ncolpropmas1[j].ToString();

                                    string propvalue = null;
                                    if (pvalues[j].Length > 100) propvalue = pvalues[j].Substring(0, 100);
                                    else propvalue = String.IsNullOrWhiteSpace(pvalues[j]) ? null : pvalues[j];
                                    tblIESBKlspropvalueTableAdapter.Insert(propvalue, periodyear, periodmonth, codeIESBK, lstypeid, lspropertieid, templateid,otdelenieIESBK);
                                    //tblIESBKlspropvalueTableAdapter.Update(DataSetIESBKLoad.tblIESBKlspropvalue);
                                } 

                                tblIESBKlspropvalueTableAdapter.Update(DataSetIESBKLoad.tblIESBKlspropvalue);
                           // } // if (findrow == null)

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }

                        spreadsheetControl1.Show();

                        break; // case "ИЭСБК Объемы ФЛ":

                    //--------------------------------------------------------------

                    case "ИЭСБК Отделения": // загрузка отделений ИЭСБК
                        DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter tblIESBKotdelenieTableAdapter = new DataSetIESBKTableAdapters.tblIESBKotdelenieTableAdapter();                        

                        spreadsheetControl1.Hide();

                        pvalues = new string[maxcol];

                        for (int i = 1; i < maxrow; i++)
                        {
                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                            tblIESBKotdelenieTableAdapter.Insert(pvalues[0], pvalues[1]);
                            tblIESBKotdelenieTableAdapter.Update(DataSetIESBKLoad.tblIESBKotdelenie);

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }

                        spreadsheetControl1.Show();

                        break; // case "ИЭСБК Отделения":

                    case "ИЭСБК желтые": // установка признака "желтый = не действует (Анита)"
                        tblIESBKlsTableAdapter = new DataSetIESBKTableAdapters.tblIESBKlsTableAdapter();
                        tblIESBKlsTableAdapter.Fill(DataSetIESBKLoad.tblIESBKls);

                        spreadsheetControl1.Hide();

                        pvalues = new string[maxcol];

                        for (int i = 1; i < maxrow; i++)
                        {
                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                            //DataRow findrow = DataSetIESBKLoad.tblIESBKls.FindBycodeIESBKlstypeidotdelenieid(pvalues[2], "ФЛ", pvalues[1]);
                            DataRow[] findrowsall = DataSetIESBKLoad.tblIESBKls.Select("codeIESBK = '" + pvalues[0] + "'");
                            foreach (DataRow findrow in findrowsall)
                            {
                                /*if (findrow != null)
                                {*/
                                    //findrow["isvalid"] = Convert.ToInt32(pvalues[3]);
                                    findrow["isvalid"] = (int)0;
                                    tblIESBKlsTableAdapter.Update(findrow);

                                    worksheet[i, 0].FillColor = Color.Green;
                                //}
                            }

                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }

                        spreadsheetControl1.Show();

                        break; // case "ИЭСБК желтые":

                    //-----------------------------------------------------------

                    case "РТП3_ЭСХ потребители": // загрузка информации о потребителях
                        RTP3ESHEntities rtp3eshEntities = new RTP3ESHEntities();
                        List<ConsumersInfo> consumersInfo_List = new List<ConsumersInfo>();

                        spreadsheetControl1.Hide();
                                                
                        pvalues = new string[maxcol];

                        for (int i = 1; i < maxrow; i++)
                        {
                            for (int j = 0; j < maxcol; j++) pvalues[j] = worksheet[i, j].Value.ToString();

                            if (!String.IsNullOrEmpty(pvalues[3]))
                            {
                                ConsumersInfo consumersInfo = new ConsumersInfo();
                                consumersInfo.id_consumersinfo = Guid.NewGuid();
                                consumersInfo.ID_DB = Convert.ToInt32(pvalues[1]);
                                consumersInfo.GUIDRTP3 = Convert.ToInt32(pvalues[2]);
                                consumersInfo.ContractRTP3 = pvalues[3];
                                consumersInfo.codeIESBK = pvalues[4];
                                consumersInfo.codeOKE = pvalues[5];
                                consumersInfo.addressInfo = pvalues[6];
                                if (pvalues[7].Equals("0") || String.IsNullOrEmpty(pvalues[7])) consumersInfo.phoneInfo = null;
                                else consumersInfo.phoneInfo = pvalues[7];
                                if (pvalues[8].Equals("0") || String.IsNullOrEmpty(pvalues[8])) consumersInfo.emailInfo = null;
                                else consumersInfo.emailInfo = pvalues[8];
                                consumersInfo_List.Add(consumersInfo);

                                /*rtp3eshEntities.ConsumersInfo.Add(consumersInfo);
                                rtp3eshEntities.SaveChanges();*/
                            }
                            splashScreenManager1.SetWaitFormDescription("Обработка данных (" + (i + 1).ToString() + " из " + maxrow.ToString() + ")");
                        }

                        rtp3eshEntities.ConsumersInfo.AddRange(consumersInfo_List);
                        rtp3eshEntities.SaveChanges();

                        spreadsheetControl1.Show();

                        break; // case "РТП3_ЭСХ потребители":

                        //-----------------------------------------------------------
                } // switch (form1.listBoxLoadType.SelectedValue.ToString())             

                //-------------------------------

                this.spreadsheetControl1.EndUpdate();
                splashScreenManager1.CloseWaitForm();

                this.spreadsheetControl1.Show();
            } // if (form1.ShowDialog(this) == DialogResult.OK)
        }

        private void spreadsheetControl1_BeforeImport(object sender, SpreadsheetBeforeImportEventArgs e)
        {
            this.Text = "Загрузка данных - ";
            this.Text += e.Options.SourceUri.ToString();
        }
    }
}