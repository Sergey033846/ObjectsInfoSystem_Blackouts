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

using GMap.NET;
using GMap.NET.WindowsForms;

namespace ObjectsInfoSystem
{
    public partial class FormSemanticValueInfo : DevExpress.XtraEditors.XtraForm
    {
        public string dbconnectionString; // строка подключения к бд

        public FormMapGeoDraw formMain;

        public int idmapsrc;
        public string idpnrmOBJECT;

        public FormSemanticValueInfo()
        {
            InitializeComponent();
        }

        // удаление объекта
        private void buttonDelObj_Click(object sender, EventArgs e)
        {
            // переделать на конкэт и 1 подключение sql

            // удаляем координаты
            string queryString = "delete FROM objectsoke.dbo.tblPanoramaCoords WHERE "+
                "idmapsrc = " + idmapsrc.ToString() + " AND idpnrmOBJECT = " + idpnrmOBJECT;
            MC_SQLDataProvider.DeleteSQLQuery(dbconnectionString, queryString);

            // удаляем объект
            queryString = "delete FROM objectsoke.dbo.tblPanoramaObject WHERE " +
                "idmapsrc = " + idmapsrc.ToString() + " AND idpnrmOBJECT = " + idpnrmOBJECT;
            MC_SQLDataProvider.DeleteSQLQuery(dbconnectionString, queryString);

            PointLatLng prevPosition = formMain.gMapControl1.Position;
            double prevZoom = formMain.gMapControl1.Zoom;
            
            formMain.LoadEraseDataFromMapSrc(false, 900);
            formMain.LoadEraseDataFromMapSrc(true, 900);

            formMain.gMapControl1.Position = prevPosition;
            formMain.gMapControl1.Zoom = prevZoom;

            this.Close();
        }

        private void buttonChangeLayer_Click(object sender, EventArgs e)
        {
            // меняем объект
            string pnrmLAYER = this.comboBoxNewLAYER.SelectedItem.ToString();
            int idpnrmGROUPLAYER = Convert.ToInt32(formMain.xmlTableLAYER.Select(String.Concat("pnrmGROUPLAYERcapt = '", pnrmLAYER, "'"))[0]["idpnrmGROUPLAYER"].ToString());

            // как сделать несколько сетов в одном запросе?
            string queryString =
                "UPDATE objectsoke.dbo.tblPanoramaObject " + 
                "SET pnrmLAYER = '" + pnrmLAYER + "'" +                
                " WHERE " +
                "idmapsrc = " + idmapsrc.ToString() + " AND idpnrmOBJECT = " + idpnrmOBJECT;
            MC_SQLDataProvider.UpdateSQLQuery(dbconnectionString, queryString);

            queryString =
                "UPDATE objectsoke.dbo.tblPanoramaObject " + 
                "SET idpnrmGROUPLAYER = " + idpnrmGROUPLAYER.ToString() +
                " WHERE " +
                "idmapsrc = " + idmapsrc.ToString() + " AND idpnrmOBJECT = " + idpnrmOBJECT;
            MC_SQLDataProvider.UpdateSQLQuery(dbconnectionString, queryString);

            PointLatLng prevPosition = formMain.gMapControl1.Position;
            double prevZoom = formMain.gMapControl1.Zoom;
                        
            formMain.LoadEraseDataFromMapSrc(false, 900);
            formMain.LoadEraseDataFromMapSrc(true, 900);

            formMain.gMapControl1.Position = prevPosition;
            formMain.gMapControl1.Zoom = prevZoom;

            this.Close();
        }
    }
}