using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ObjectsInfoSystem
{
    public partial class FormSetObjectFromXML : DevExpress.XtraEditors.XtraForm
    {
        //public int srcFileType; // 0 - GPX, 1 - SAS.Planet

        public FormMapGeoDraw formMain;

        public FormSetObjectFromXML()
        {
            InitializeComponent();
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            // устанавливаем значения имен столбцов в зависимости от вида источника
            string columnNAME = "";
            string columnLON = "";
            string columnLAT = "";
            if (formMain.srcFileType == 0) // GPX
            {
                columnNAME = "name";
                columnLON = "lon";
                columnLAT = "lat";
            }
            else
            if (formMain.srcFileType == 1) // SAS.Planet
            {
                columnNAME = "name";
                columnLON = "lonL";
                columnLAT = "latT";
            }

            // "парсим" строку, разделители запятые, внутри - тире (минус)            
            string srcPoint = memoEditSrcPoints.Text;
            string[] words;

            // удаляем пробелы и разбиваем
            srcPoint = srcPoint.Replace(" ", "");
            words = srcPoint.Split(new char[] { ',', '.' }, 300);

            // разбиваем каждое слово
            //GMapOverlay overlayXML = MC_GMap.GetOverlayIndexByName(gMapControl1, "43"); // исправить!!!

            // проверяем на корректность введенную информацию
            bool flag_error = false;
            foreach (string word in words)
            {
                string[] points = word.Split(new char[] { '-' }, 2);

                if (points.Count() > 0)
                {
                    if (!String.IsNullOrWhiteSpace(points[0]))
                    {
                        //int point1 = Convert.ToInt32(points[0]);
                        string point1 = points[0];
                        DataRow[] fRows = formMain.xmlTable.Select(String.Concat(columnNAME, " = '", points[0], "'"));
                        if (fRows.Count() != 1)
                        {
                            memoEditTest.Text += "точка " + points[0] + " не корректна!" + Environment.NewLine;
                            flag_error = true;
                        }

                        if (points.Count() == 2)
                        {
                            if (!String.IsNullOrWhiteSpace(points[1]))
                            {
                                int point2int = Convert.ToInt32(points[1]);
                                fRows = formMain.xmlTable.Select(String.Concat(columnNAME, " = '", points[1], "'"));
                                if (fRows.Count() != 1)
                                {
                                    memoEditTest.Text += "точка " + points[1] + " не корректна!" + Environment.NewLine;
                                    flag_error = true;
                                }

                                int point1int = Convert.ToInt32(points[0]);
                                if (point2int <= point1int) // добавить дополнительно проверку в цикле от нач к конечному!!!
                                {
                                    memoEditTest.Text += "нарушена последовательность точек!!!" + Environment.NewLine;
                                    flag_error = true;
                                }
                            } // if (!String.IsNullOrWhiteSpace(points[1]))
                            else flag_error = true;

                        } // if (points.Count() == 2)

                    } // if (!String.IsNullOrWhiteSpace(points[0]))
                    else flag_error = true;

                } // if (points.Count() > 0)    
                else
                {
                    memoEditTest.Text += "введите последовательность точек!!!" + Environment.NewLine;
                    flag_error = true;
                }

            } // foreach (string word in words)
            //-------------------------------------------------------------------------

            // если парсинг прошел успешно, делаем то же самое, но уже с записью

            DataSetPnrmMapSrc DataSetLoad = new DataSetPnrmMapSrc();

            DataSetPnrmMapSrcTableAdapters.tblPanoramaObjectTableAdapter tblPanoramaObjectTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaObjectTableAdapter();
            DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter tblPanoramaCoordsTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter();

            if (flag_error == false)
            {
                // добавляем объект к слою

                // формируем параметры объекта
                // найти последнее значение idpnrmOBJECT!!!
                string queryStringObj =
                    String.Concat("SELECT MAX(idpnrmOBJECT) as maxobj FROM objectsoke.dbo.tblPanoramaObject WHERE idmapsrc = 900"); // ПЕРЕДЕЛАТЬ!!!
                DataTable TableMaxIdPnrmObject = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(TableMaxIdPnrmObject, formMain.dbconnectionString, queryStringObj);
                int last_idpnrmOBJECT = Convert.ToInt32(TableMaxIdPnrmObject.Rows[0]["maxobj"].ToString());
                TableMaxIdPnrmObject.Dispose();

                int idpnrmOBJECT;
                if (last_idpnrmOBJECT < 900000) idpnrmOBJECT = 900000 + 1;
                else idpnrmOBJECT = last_idpnrmOBJECT + 1;
                                
                int idmapsrc = 900; // исправить в будущем!!!

                string pnrmLOCAL = listBoxControlLOCAL.SelectedItem.ToString();                                
                int idpnrmLOCALtype = Convert.ToInt32(formMain.xmlTableLOCAL.Select(String.Concat("pnrmLOCALtypecapt = '", pnrmLOCAL, "'"))[0]["idpnrmLOCALtype"].ToString());

                string pnrmLAYER = listBoxControlLAYER.SelectedItem.ToString();
                int idpnrmGROUPLAYER = Convert.ToInt32(formMain.xmlTableLAYER.Select(String.Concat("pnrmGROUPLAYERcapt = '", pnrmLAYER, "'"))[0]["idpnrmGROUPLAYER"].ToString());
                //------------------------------------------

                if (pnrmLOCAL == "линейный" || pnrmLOCAL == "площадной")
                {
                    tblPanoramaObjectTableAdapter.Insert(idpnrmOBJECT, idmapsrc, "загружено", null, pnrmLAYER, pnrmLOCAL, null, null, null, null, idpnrmLOCALtype, idpnrmGROUPLAYER);
                    tblPanoramaObjectTableAdapter.Update(DataSetLoad.tblPanoramaObject);
                }

                int pointID = 0;
                int pnrmSUBJECT = 0;
                foreach (string word in words)
                {
                    string[] points = word.Split(new char[] { '-' }, 2);
                    
                    if (points.Count() > 0)
                    {
                        //int point1 = Convert.ToInt32(points[0]);
                        string point1 = points[0];
                        DataRow[] fRows = formMain.xmlTable.Select(String.Concat(columnNAME, " = '", points[0], "'"));

                        string coordALT = fRows[0][columnLAT].ToString().Replace(".",",");
                        string coordLONG = fRows[0][columnLON].ToString().Replace(".", ",");
                        double coordALTfloat = Convert.ToDouble(coordALT);
                        double coordLONGfloat = Convert.ToDouble(coordLONG);

                        if (pnrmLOCAL == "точечный")
                        {
                            tblPanoramaObjectTableAdapter.Insert(idpnrmOBJECT, idmapsrc, "загружено", null, pnrmLAYER, pnrmLOCAL, null, null, null, null, idpnrmLOCALtype, idpnrmGROUPLAYER);
                            tblPanoramaObjectTableAdapter.Update(DataSetLoad.tblPanoramaObject);
                        }
                        tblPanoramaCoordsTableAdapter.Insert(idpnrmOBJECT, pointID, idmapsrc, pointID++, null, null, null, coordALT, coordLONG, pnrmSUBJECT, coordALTfloat, coordLONGfloat);
                        if (pnrmLOCAL == "точечный") idpnrmOBJECT++;

                        if (points.Count() == 2)
                        {
                            int point2 = Convert.ToInt32(points[1]);
                            int point1int = Convert.ToInt32(points[0]);

                            for (int i = point1int+1; i <= point2; i++)
                            {
                                fRows = formMain.xmlTable.Select(String.Concat(columnNAME, " = '", i.ToString(), "'"));

                                coordALT = fRows[0][columnLAT].ToString().Replace(".", ",");
                                coordLONG = fRows[0][columnLON].ToString().Replace(".", ",");
                                coordALTfloat = Convert.ToDouble(coordALT);
                                coordLONGfloat = Convert.ToDouble(coordLONG);

                                if (pnrmLOCAL == "точечный")
                                {
                                    tblPanoramaObjectTableAdapter.Insert(idpnrmOBJECT, idmapsrc, "загружено", null, pnrmLAYER, pnrmLOCAL, null, null, null, null, idpnrmLOCALtype, idpnrmGROUPLAYER);
                                    tblPanoramaObjectTableAdapter.Update(DataSetLoad.tblPanoramaObject);
                                }
                                tblPanoramaCoordsTableAdapter.Insert(idpnrmOBJECT, pointID, idmapsrc, pointID++, null, null, null, coordALT, coordLONG, pnrmSUBJECT, coordALTfloat, coordLONGfloat);
                                if (pnrmLOCAL == "точечный") idpnrmOBJECT++;
                            }

                        } // if (points.Count() == 2)
                                                    
                    } // if (points.Count() > 0)             

                } // foreach (string word in words)

                tblPanoramaCoordsTableAdapter.Update(DataSetLoad.tblPanoramaCoords);
            }

            //-------------------------------------------------------------------------

            if (flag_error) memoEditTest.Text += "слишком много ошибок!!!" + Environment.NewLine;
            else DialogResult = DialogResult.OK;
        }

        private void FormSetObjectFromXML_Load(object sender, EventArgs e)
        {

        }
    }
}