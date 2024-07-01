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
using DevExpress.Spreadsheet;

namespace ObjectsInfoSystem
{
    public partial class FormCoordZonesForLoad : DevExpress.XtraEditors.XtraForm
    {
        public DataSetPnrmMapSrc DataSetLoad;
        public DataSetPnrmMapSrcTableAdapters.tblPanoramaCoordsTableAdapter tblPanoramaCoordsTableAdapter;

        public FormCoordZonesForLoad()
        {
            InitializeComponent();
        }

        public void GetNULLWGS84CoordsByZone (int zone)
        {
            FormLoadData form1 = null;
            form1 = new FormLoadData();
            form1.Text = "Выгрузка данных - "+zone.ToString() + " зона";
                        
            //splashScreenManager1.ShowWaitForm();
            IWorkbook workbook = form1.spreadsheetControl1.Document;
            workbook.History.IsEnabled = false;
            //Worksheet worksheet = workbook.Worksheets[0];

            form1.spreadsheetControl1.BeginUpdate();
            form1.spreadsheetControl1.Hide();

            //DataSet1 DataSetLoad = new DataSet1();
            //DataSet1TableAdapters.tblPanoramaCoordsTableAdapter tblPanoramaCoordsTableAdapter = new DataSet1TableAdapters.tblPanoramaCoordsTableAdapter();
            //tblPanoramaCoordsTableAdapter.Fill(DataSetLoad.tblPanoramaCoords);
            DataRow[] coordrows = DataSetLoad.tblPanoramaCoords.Select("coordALT is NULL AND SUBSTRING(pnrmY,1,1) = '"+zone.ToString()+"'");
            //DataRow[] coordrows = DataSetLoad.tblPanoramaCoords.Select("SUBSTRING(pnrmY,1,1) = '" + zone.ToString() + "'");

            /*int row = 0;
            worksheet[row, 0].Value = "IDMAPSRC";
            worksheet[row, 1].Value = "PNRMPOINT";
            worksheet[row, 2].Value = "IDPNRMOBJECT";
            worksheet[row, 3].Value = "PNRMSUBJECT";
            worksheet[row, 4].Value = "X";
            worksheet[row, 5].Value = "Y";
            worksheet[row, 6].Value = "X,Y";
            worksheet[row, 7].Value = "ALT,LONGT";            */

            int maxcoordrowsinwrksh = 8000;
            int coordrowsincolumn = 2000; // текущая виртуальная "колонка координат" со служебными полями
            int maxcolumninwrksh = maxcoordrowsinwrksh / coordrowsincolumn;
            int maxCount = coordrows.Count();
            int maxwrksh = maxCount / maxcoordrowsinwrksh;
            if (maxwrksh * maxcoordrowsinwrksh < maxCount) maxwrksh++;

            for (int wrksh = 0; wrksh < maxwrksh - 1; wrksh++)
                workbook.Worksheets.Add();

            for (int j = 0; j < maxwrksh; j++)
            {
                Worksheet worksheet = workbook.Worksheets[j];

                for (int column = 0; column < maxcoordrowsinwrksh / coordrowsincolumn; column++)
                {
                    for (int row = 0; row < coordrowsincolumn; row++)
                    {
                        int rowbd = j * maxcoordrowsinwrksh + column * coordrowsincolumn + row;

                        if (rowbd >= maxCount) break;

                        worksheet[row, (column * 8) + 0].Value = coordrows[rowbd]["idmapsrc"].ToString();
                        worksheet[row, (column * 8) + 1].Value = coordrows[rowbd]["pnrmPOINT"].ToString();
                        worksheet[row, (column * 8) + 2].Value = coordrows[rowbd]["idpnrmOBJECT"].ToString();
                        worksheet[row, (column * 8) + 3].Value = coordrows[rowbd]["pnrmSUBJECT"].ToString();
                        worksheet[row, (column * 8) + 4].Value = coordrows[rowbd]["pnrmX"].ToString();
                        worksheet[row, (column * 8) + 5].Value = coordrows[rowbd]["pnrmY"].ToString();
                        worksheet[row, (column * 8) + 6].Value = 
                            String.Concat(worksheet[row, (column * 8) + 4].Value.ToString().Replace(",","."), 
                            ",", 
                            worksheet[row, (column * 8) + 5].Value.ToString().Replace(",", "."));
                    }
                    worksheet.Columns[column * 8 + 6].FillColor = Color.Orange;
                    worksheet.Columns[column * 8 + 7].FillColor = Color.DeepSkyBlue;
                }

                worksheet.Columns.AutoFit(0, 8 * maxcolumninwrksh);
            }

            form1.spreadsheetControl1.EndUpdate();
            form1.spreadsheetControl1.Show();
            //splashScreenManager1.CloseWaitForm();

            form1.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(3);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(4);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(5);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(6);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(7);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            GetNULLWGS84CoordsByZone(8);
        }
    }
}