using GMap.NET;
using GMap.NET.WindowsForms;

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ObjectsInfoSystem.MyClasses.gmap;

namespace ObjectsInfoSystem.MyClasses
{
    class MC_MapSource
    {
        public int IdMapSrc { get; set; }
            
        public DataTable tableObj;  // таблица объектов карты
        public DataTable tableCoord; // таблица координат объектов карты
        public DataTable tableSEMvalues; // таблица значений семантики карты

        private DataTable tableALTLONG;
        private DataTable tableMAXSubject;

        public DataTable tableLAYERname;
        public DataTable tableLEGEND;

            //public DataTable tableLEGEND; // таблица свойств легенды
        public double ALTmin, ALTmax, LONGmin, LONGmax; // границы прямоугольника координат карты

        public string dbconnectionString;

        public MC_MapSource(int idmapsrc)
        {
            dbconnectionString = "Data Source=SERVERPFB;Initial Catalog=objectsoke;User ID=sa;Password=SqL1310198"; // переделать!!!!

            SqlConnection SQLconnection = new SqlConnection(dbconnectionString);
            SQLconnection.Open();

            List<SqlParameter> listSQLparams = new List<SqlParameter>();
            listSQLparams.Add(new SqlParameter("@idmapsrc", idmapsrc));
              
            // загружаем объекты -----------------      
            /*string queryStringObj = "SELECT tblPO.idpnrmOBJECT,tblPO.idmapsrc,tblPO.pnrmNAME,tblPO.pnrmOBJECTKEY,tblPO.pnrmLAYER,tblPO.pnrmLOCAL," + 
                                    "tblPO.pnrmLENGTH,tblPO.pnrmSQUARE,tblPO.pnrmLINKOBJECT,tblPO.pnrmLINKSHEET,tblPO.idpnrmLOCALtype,tblPO.idpnrmGROUPLAYER," +                                    
		                            "tblPL.PENcolor,tblPL.PENthickness,tblPL.PENstyle,tblPL.FILLSOLIDcolor,tblPL.FILLSOLIDalpha," +
		                            "tblPL.FILLHATCHforecolor,tblPL.FILLHATCHbackcolor,tblPL.FILLHATCHstyle ,tblPL.POINTradius,tblPL.POINTtype " +
                                    "FROM objectsoke.dbo.tblPanoramaObject tblPO " +                                    
                                    "INNER JOIN objectsoke.dbo.tblPanoramaLegend tblPL ON (tblPO.idpnrmGROUPLAYER = tblPL.idpnrmGROUPLAYER) and(tblPO.idpnrmLOCALtype = tblPL.idpnrmLOCALtype)" +
                                    "WHERE(tblPO.idmapsrc = " + idmapsrc.ToString() + ") AND (tblPL.idlgnddest = 1)";*/
            string queryStringObj = 
                String.Concat("SELECT tblPO.idpnrmOBJECT,tblPO.idmapsrc,tblPO.pnrmNAME,tblPO.pnrmOBJECTKEY,tblPO.pnrmLAYER,tblPO.pnrmLOCAL,",
                "tblPO.pnrmLENGTH,tblPO.pnrmSQUARE,tblPO.pnrmLINKOBJECT,tblPO.pnrmLINKSHEET,tblPO.idpnrmLOCALtype,tblPO.idpnrmGROUPLAYER",                
                " FROM objectsoke.dbo.tblPanoramaObject tblPO",                
                " WHERE tblPO.idmapsrc = ", idmapsrc.ToString());
            tableObj = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableObj, queryStringObj);                  
            
            /*listSQLparams.Add(new SqlParameter("@idmapsrc", idmapsrc));
            MyFuncSQL.CallStoredProcedure(SQLconnection, tableObj, "dbo.myprocGetPnrmObjectsByIdmapsrc", listSQLparams);*/
            //------------------------------------
            
            // загружаем координаты -------------- (подумать про убрать idmapsrc из запроса)
            string queryStringCoord = String.Concat("select idpnrmOBJECT, idmapsrc, pnrmPOINT, pnrmSUBJECT, coordALTfloat, coordLONGfloat from tblPanoramaCoords where tblPanoramaCoords.idmapsrc = ", idmapsrc.ToString());
            tableCoord = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableCoord, queryStringCoord);
            //------------------------------------

            // определяем мин/макс широту/долготу ---------------------
            string queryString =
                String.Concat("select min(tblPanoramaCoords.coordALTfloat) as Min_coordALT, min(tblPanoramaCoords.coordLONGfloat) as Min_coordLONG,",
                "max(tblPanoramaCoords.coordALTfloat) as Max_coordALT, max(tblPanoramaCoords.coordLONGfloat) as Max_coordLONG ",
                "from tblPanoramaCoords where idmapsrc = ", idmapsrc.ToString());
            tableALTLONG = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableALTLONG, queryString);
            
            // формируем и заполняем селектор слоев, создаем слои компонента GMAP            
            
            string queryStringLAYERname =
                String.Concat("SELECT DISTINCT tblPO.idpnrmGROUPLAYER, tblPGL.pnrmGROUPLAYERcapt ",
                "FROM objectsoke.dbo.tblPanoramaObject tblPO ",
                "INNER JOIN objectsoke.dbo.tblPanoramaGroupLayer tblPGL ON tblPO.idpnrmGROUPLAYER = tblPGL.idpnrmGROUPLAYER ",
                "WHERE tblPO.idmapsrc = ", idmapsrc.ToString(),
                " ORDER BY tblPGL.pnrmGROUPLAYERcapt ASC");
            /*string queryStringLAYERname =
                String.Concat("SELECT DISTINCT tblPO.idpnrmGROUPLAYER, tblPGL.pnrmGROUPLAYERcapt,",
                              "tblPL.PENcolor, tblPL.PENthickness, tblPL.PENstyle, tblPL.FILLSOLIDcolor, tblPL.FILLSOLIDalpha,",
                              "tblPL.FILLHATCHforecolor, tblPL.FILLHATCHbackcolor, tblPL.FILLHATCHstyle, tblPL.POINTradius, tblPL.POINTtype",
                              " FROM objectsoke.dbo.tblPanoramaObject tblPO",
                              " INNER JOIN objectsoke.dbo.tblPanoramaGroupLayer tblPGL ON tblPO.idpnrmGROUPLAYER = tblPGL.idpnrmGROUPLAYER",
                              " INNER JOIN objectsoke.dbo.tblPanoramaLegend tblPL ON(tblPO.idpnrmGROUPLAYER = tblPL.idpnrmGROUPLAYER) and(tblPO.idpnrmLOCALtype = tblPL.idpnrmLOCALtype)",
                              " WHERE (tblPO.idmapsrc = ", idmapsrc.ToString(), ") AND (tblPL.idlgnddest = 1)",
                              " ORDER BY tblPGL.pnrmGROUPLAYERcapt ASC");*/
            tableLAYERname = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLAYERname, queryStringLAYERname);

            // загружаем семантику ------------------------------------------------------------------------------------
            string querySEMvalues = String.Concat("select * from ((tblPanoramaObject tblPanoramaObject ",
                "inner join tblPanoramaObjSemValues tblPanoramaObjSemValues on((tblPanoramaObjSemValues.idpnrmOBJECT = tblPanoramaObject.idpnrmOBJECT) and (tblPanoramaObjSemValues.idmapsrc = tblPanoramaObject.idmapsrc))) ",
                "inner join tblPanoramaSemantic tblPanoramaSemantic on((tblPanoramaSemantic.idpnrmSEM = tblPanoramaObjSemValues.idpnrmSEM) and(tblPanoramaSemantic.idmapsrc = tblPanoramaObjSemValues.idmapsrc))) ",
                "where tblPanoramaObjSemValues.idmapsrc = ", idmapsrc.ToString());
            tableSEMvalues = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableSEMvalues, querySEMvalues);
            //---------------------------------------------------------------------------------------------------------

            // пусть пока так (грузится вся легенда с фильтром по слоям карты)
            string queryString1 = String.Concat(
                "SELECT * FROM objectsoke.dbo.tblPanoramaLegend",
                " WHERE idpnrmGROUPLAYER in",
                    " (SELECT DISTINCT idpnrmGROUPLAYER FROM objectsoke.dbo.tblPanoramaObject tblPO WHERE tblPO.idmapsrc = ", idmapsrc.ToString(), ") ",
                " ORDER BY idpnrmGROUPLAYER ASC");
            tableLEGEND = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, tableLEGEND, queryString1);

            // получаем таблицу max Subject-ов - тестим производительность            
            tableMAXSubject = new DataTable();
            MC_SQLDataProvider.CallStoredProcedure(SQLconnection, tableMAXSubject, "dbo.myprocGetPnrmTableMaxSubject", listSQLparams);
                        
            SQLconnection.Close();

            IdMapSrc = idmapsrc;
        }

        public void DrawMapSource(GMapControl gmc)
        {
            // создаем массив стилей (УБРАТЬ в МОДУЛЬ!!!)
            DashStyle[] dashStylesPen = { DashStyle.Dash, DashStyle.DashDot, DashStyle.DashDotDot, DashStyle.Dot, DashStyle.Solid };

            double latitude = 0.0;
            double longitude = 0.0;

            // почему это здесь, а не в конструкторе?
            ALTmin = (double)tableALTLONG.Rows[0]["Min_coordALT"];
            ALTmax = (double)tableALTLONG.Rows[0]["Max_coordALT"];
            LONGmin = (double)tableALTLONG.Rows[0]["Min_coordLONG"];
            LONGmax = (double)tableALTLONG.Rows[0]["Max_coordLONG"];

            foreach (DataRow legendRow in tableLEGEND.Select("idlgnddest = 1")) // !!!!! idlgnddest
            {
                DataRow[] tableObjRows = tableObj.Select(
                    String.Concat("idpnrmGROUPLAYER = ", legendRow["idpnrmGROUPLAYER"].ToString(),
                                  " AND idpnrmLOCALtype = ", legendRow["idpnrmLOCALtype"].ToString()));

                // формируем стили -----------                
                Pen layerPen = new Pen(Color.FromArgb((int)legendRow["PENcolor"]), (int)legendRow["PENthickness"]);
                layerPen.DashStyle = dashStylesPen[(int)legendRow["PENstyle"]];

                Brush layerBrush = new SolidBrush(Color.FromArgb((int)legendRow["FILLSOLIDalpha"], Color.FromArgb((int)legendRow["FILLSOLIDcolor"])));

                int pointType = (int)legendRow["POINTtype"];
                int pointRadius = (int)legendRow["POINTradius"];

                GMapOverlay overlayCurrent = MC_GMap.GetOverlayIndexByName(gmc, legendRow["idpnrmGROUPLAYER"].ToString());

                // бежим по объектам
                foreach (DataRow objRow in tableObjRows)
                {
                    int idpnrmOBJECT = (int)objRow["idpnrmOBJECT"]; 
                    string idpnrmOBJECTstr = String.Concat("M", IdMapSrc.ToString(),"-",idpnrmOBJECT.ToString());
                    string idpnrmOBJECTstrFilter = String.Concat("idpnrmOBJECT = ", idpnrmOBJECT);
                    int idpnrmLOCALtype = (int)objRow["idpnrmLOCALtype"];

                    DataRow[] objCoordRows = tableMAXSubject.Select(idpnrmOBJECTstrFilter);
                    //DataRow[] objCoordsRows = tableMAXSubject.Select("idpnrmOBJECT = " + idpnrmOBJECTstr);
                    int pnrmSUBJECTmaxvalue = (objCoordRows.Count() > 0) ? (int)objCoordRows[0]["maxSUBJECT"] : 0;
                    //------------------------------------------------

                    string tempSUBJECTstr = String.Concat(idpnrmOBJECTstrFilter, " AND pnrmSUBJECT = ");
                    string coordsort = "pnrmPOINT ASC";

                    for (int sbj = 0; sbj < pnrmSUBJECTmaxvalue + 1; sbj++)
                    {
                        List<GMap.NET.PointLatLng> points = new List<GMap.NET.PointLatLng>();

                        string coordexpr = String.Concat(tempSUBJECTstr, sbj);
                        //string coordsort = "pnrmPOINT ASC";                                        
                        objCoordRows = tableCoord.Select(coordexpr, coordsort);
                                                
                        if (idpnrmLOCALtype != 5) // если объект не точечный
                        {
                            foreach (DataRow objCoordRow in objCoordRows)
                            {
                                latitude = (double)objCoordRow["coordALTfloat"];
                                longitude = (double)objCoordRow["coordLONGfloat"];

                                //if (idpnrmLOCALtype == 2 || idpnrmLOCALtype == 3) points.Add(new GMap.NET.PointLatLng(latitude, longitude));

                                points.Add(new GMap.NET.PointLatLng(latitude, longitude));
                            } // foreach (DataRow objCoordRow in objCoordsRows)
                        }
                        else
                        if (objCoordRows.Count() > 0)
                        {
                            latitude = (double)objCoordRows[0]["coordALTfloat"];
                            longitude = (double)objCoordRows[0]["coordLONGfloat"];
                        }

                        switch (idpnrmLOCALtype)
                        {
                            case 2: // линейный                            
                                GMapRoute r = new GMapRoute(points, idpnrmOBJECTstr);
                                r.IsHitTestVisible = true;

                                r.Stroke = layerPen;

                                //r.Tag = GetPnrmObjectProperties(idpnrmOBJECTstr);
                                r.Tag = idpnrmOBJECTstr;

                                //gmc.Overlays[layerid].Routes.Add(r);
                                overlayCurrent.Routes.Add(r);
                                break;

                            case 3: // площадной                            
                                GMap.NET.WindowsForms.GMapPolygon polygon = new GMap.NET.WindowsForms.GMapPolygon(points, idpnrmOBJECTstr);

                                polygon.Fill = layerBrush;
                                polygon.Stroke = layerPen;

                                polygon.IsHitTestVisible = true;

                                polygon.Tag = idpnrmOBJECTstr;

                                //gmc.Overlays[layerid].Polygons.Add(polygon);
                                overlayCurrent.Polygons.Add(polygon);
                                break;

                            case 5: // точечный
                                MC_GMapMarkerShape markerP = new MC_GMapMarkerShape(idpnrmOBJECTstr, new GMap.NET.PointLatLng(latitude, longitude), layerPen, layerBrush, (my_MarkerType)pointType, pointRadius);
                                
                                markerP.Tag = idpnrmOBJECTstr;

                                //gmc.Overlays[layerid].Markers.Add(markerG);
                                overlayCurrent.Markers.Add(markerP);

                                break;

                            case 4: // подпись
                                MC_GMapMarkerText markerT = new MC_GMapMarkerText(idpnrmOBJECTstr, new GMap.NET.PointLatLng(latitude, longitude), 
                                    GetObjectTooltipString(idpnrmOBJECT), 0, new SolidBrush(Color.Blue));

                                markerT.Tag = idpnrmOBJECTstr;
                                                                
                                overlayCurrent.Markers.Add(markerT);

                                //GMap.NET.WindowsForms.Markers.GMarkerGoogleType marker_type = GMap.NET.WindowsForms.Markers.GMarkerGoogleType.orange_small;
                                //GMap.NET.WindowsForms.Markers.GMarkerGoogle markerT = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(new GMap.NET.PointLatLng(latitude, longitude), marker_type);

                                // формируем текст маркера-подписи --------
                                /*string markerText = "";

                                float angleText = 30;
                                latitude = (double)objCoordsRows[0]["coordALTfloat"];
                                longitude = (double)objCoordsRows[0]["coordLONGfloat"];
                                PointLatLng pointTextstart = new GMap.NET.PointLatLng(latitude, longitude);

                                GMapMarkerText markerT = new GMapMarkerText(pointTextstart, markerText, angleText);

                                //markerT.ToolTip = new GMap.NET.WindowsForms.ToolTips.GMapBaloonToolTip(markerT);
                                //markerT.ToolTipMode = GMap.NET.WindowsForms.MarkerTooltipMode.Always;
                                markerT.ToolTipMode = GMap.NET.WindowsForms.MarkerTooltipMode.Never;

                                markerT.Tag = idpnrmOBJECT;
                                //markerT.ToolTipText = (i + 1).ToString();                            

                                markerT.IsHitTestVisible = true;

                                gmc.Overlays[layerid].Markers.Add(markerT);*/

                                break;
                        }

                    } // for (int sbj = 0; sbj < pnrmSUBJECTmaxvalue + 1; sbj++)

                } // for (int i = 0; i < objRows.Count(); i++)            
            } // foreach (DataRow legendRow in tableLEGEND.Rows)

            /*swatch.Stop();
            memoEditSemantic.Text += swatch.Elapsed.ToString() + Environment.NewLine;
            // диагностика конец*/

            // Устанавливаем центральную позицию карты
            if (IdMapSrc == 6) gmc.Position = new GMap.NET.PointLatLng(52.0323866448363, 105.406642646227);
            else if (IdMapSrc == 7) gmc.Position = new GMap.NET.PointLatLng(52.1369961023287, 105.278848418755);
            else gmc.Position = new GMap.NET.PointLatLng((ALTmax + ALTmin) / 2, (LONGmax + LONGmin) / 2);

            //Указываем, что при загрузке карты будет использоваться
            //17ти кратное приближение.
            gmc.Zoom = 17;

            //Обновляем карту.
            gmc.Refresh();                        
        }

        public void EraseMapSource(GMapControl gmc)
        {
            string idMapSrcstr = String.Concat("M", IdMapSrc.ToString(), "-");

            // бежим по слоям
            foreach (GMapOverlay gmcOverlay in gmc.Overlays) 
            {
                // бежим по объектам слоя
                for (int i = 0; i < gmcOverlay.Markers.Count; i++)
                {
                    GMapMarker gmcObject = gmcOverlay.Markers[i];
                    if (gmcObject.Tag.ToString().Contains(idMapSrcstr))
                    {
                        gmcOverlay.Markers.Remove(gmcObject);
                        i--;
                    }
                }

                for (int i = 0; i < gmcOverlay.Routes.Count; i++)
                {
                    GMapRoute gmcObject = gmcOverlay.Routes[i];
                    if (gmcObject.Tag.ToString().Contains(idMapSrcstr))
                    {
                        gmcOverlay.Routes.Remove(gmcObject);
                        i--;
                    }
                }

                for (int i = 0; i < gmcOverlay.Polygons.Count; i++)
                {
                    GMapPolygon gmcObject = gmcOverlay.Polygons[i];
                    if (gmcObject.Tag.ToString().Contains(idMapSrcstr))
                    {
                        gmcOverlay.Polygons.Remove(gmcObject);
                        i--;
                    }
                }
            } 

            // обновляем карту            
            gmc.Refresh();
        }

        public string GetObjectTooltipString(int idpnrmOBJECT)
        {
            StringBuilder tooltipString = new StringBuilder();
            
            DataRow[] objSEMvaluesrows = tableSEMvalues.Select(String.Concat("idpnrmOBJECT = ", idpnrmOBJECT.ToString()));

            foreach (DataRow objSEMvaluesrow in objSEMvaluesrows)
            {                
                tooltipString.Append(objSEMvaluesrow["pnrmSEMvalue"].ToString());
            }

            return tooltipString.ToString();
        }
    }
}
