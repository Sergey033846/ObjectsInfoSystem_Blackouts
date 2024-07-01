using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;

using GMap.NET;
using GMap.NET.WindowsForms;

using ObjectsInfoSystem.MyClasses;
using ObjectsInfoSystem.MyClasses.gmap;

namespace ObjectsInfoSystem
{
    public partial class FormMapGeoDraw : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        // пользовательские переменные -------------
        int permissionAdmin; // признак прав доступа администратора
        
        public int srcFileType; // default = -1, GPX = 0, SAS.Planet = 1

        //public int current_idmapsrc; // текущая загруженная карта, -1 - если ничего не загружено
        public string dbconnectionString; // строка подключения к бд

        private MC_SQLDataProvider SQLDataProvider;
        private MC_MapLegend MapLegend;

        public TreeNode nodemapsrcloaded; // ссылка на текущий загруженный узел дерева Филиал/Карта
        //public double ALTmin, ALTmax, LONGmin, LONGmax; // границы прямоугольника координат карты

        //private MC_MapSource MapSrc;

        private IDictionary<int, MC_MapSource> mapSrcDictionary;
        private IDictionary<int, CheckState> layerVisibleDictionary;

        public DataSet xmlDataSet;
        public DataTable xmlTable;
        public DataTable xmlTableLAYER;
        public DataTable xmlTableLOCAL;

        //public DataTable tableObj; // таблица объектов карты
        //public DataTable tableCoord; // таблица координат объектов карты
        //public DataTable tableSEMvalues; // таблица значений семантики карты
        //public DataTable tableLEGEND; // таблица свойств легенды

        //FormSemanticValueInfo formShowSEMvalues_temp; // форма вывода значений семантики
        FormLegendShow formShowLegend;
        //------------------------------------------
                
        //-------------------------------
        // класс "пользовательский маркер-подпись"
        
        // заполнение таблиц информационной формы при нажатии на карту
        private void ShowObjectSEMvalues(string idpnrmOBJECTstr)
        {
            FormSemanticValueInfo formShowSEMvalues = new FormSemanticValueInfo();

            // убрать в конструктор!!!!
            formShowSEMvalues.Text = "Свойства объекта";
            formShowSEMvalues.formMain = this;

            if (permissionAdmin == 100 || permissionAdmin == 2)
            {
                formShowSEMvalues.buttonDelObj.Enabled = true;
                formShowSEMvalues.buttonChangeLayer.Enabled = true;                
            }
            else
            {
                formShowSEMvalues.buttonDelObj.Enabled = false;
                formShowSEMvalues.buttonChangeLayer.Enabled = false;
            }
            
            foreach (DataRow drow in xmlTableLAYER.Rows)
            {
                formShowSEMvalues.comboBoxNewLAYER.Properties.Items.Add(drow["pnrmGROUPLAYERcapt"].ToString());
            }
            //------------------------------------------

            int position = idpnrmOBJECTstr.IndexOf("-");                        
            /*if (position < 0)
                continue;*/
            /*Console.WriteLine("Key: {0}, Value: '{1}'",
                              pair.Substring(0, position),
                              pair.Substring(position + 1));*/
            int mapSrc = Convert.ToInt32(idpnrmOBJECTstr.Substring(1, position-1)); // не с 0, т.к. пропускаем букву M
            string idpnrmOBJECT = idpnrmOBJECTstr.Substring(position + 1);

            formShowSEMvalues.dbconnectionString = dbconnectionString;
            formShowSEMvalues.idmapsrc = mapSrc;
            formShowSEMvalues.idpnrmOBJECT = idpnrmOBJECT;

            MC_MapSource MapSrc = mapSrcDictionary[mapSrc];
            string fexpr2 = String.Concat("idpnrmOBJECT = ", idpnrmOBJECT);
            DataRow[] objrows = MapSrc.tableObj.Select(fexpr2);
            DataRow[] objSEMvaluesrows = MapSrc.tableSEMvalues.Select(fexpr2);
            
            for (int i = 0; i < objSEMvaluesrows.Count(); i++)
            {
                formShowSEMvalues.dataGridView1.Rows.Add(objSEMvaluesrows[i]["pnrmSEMcaption"].ToString(), objSEMvaluesrows[i]["pnrmSEMvalue"].ToString());                
            }

            for (int i = 0; i < MapSrc.tableObj.Columns.Count; i++)            
            {
                formShowSEMvalues.dataGridView2.Rows.Add(MapSrc.tableObj.Columns[i].Caption, objrows[0][i].ToString());                
            }

            //
            formShowSEMvalues.Show();            
        }

        //---------------------------
        private GMap.NET.MapProviders.GMapProvider ReturnGMapProviderByString(string MapProviderStr)
        {
            if (MapProviderStr.Equals("OpenStreetMap")) return GMap.NET.MapProviders.GMapProviders.OpenStreetMap;
            if (MapProviderStr.Equals("OpenCycleMap")) return GMap.NET.MapProviders.GMapProviders.OpenCycleMap;
            if (MapProviderStr.Equals("OpenCycleLandscapeMap")) return GMap.NET.MapProviders.GMapProviders.OpenCycleLandscapeMap;
            if (MapProviderStr.Equals("OpenCycleTransportMap")) return GMap.NET.MapProviders.GMapProviders.OpenCycleTransportMap;
            if (MapProviderStr.Equals("OpenStreetMapQuest")) return GMap.NET.MapProviders.GMapProviders.OpenStreetMapQuest;
            if (MapProviderStr.Equals("OpenSeaMapHybrid")) return GMap.NET.MapProviders.GMapProviders.OpenSeaMapHybrid;
            if (MapProviderStr.Equals("WikiMapiaMap")) return GMap.NET.MapProviders.GMapProviders.WikiMapiaMap;
            if (MapProviderStr.Equals("GoogleMap")) return GMap.NET.MapProviders.GMapProviders.GoogleMap;
            if (MapProviderStr.Equals("GoogleTerrainMap")) return GMap.NET.MapProviders.GMapProviders.GoogleTerrainMap;
            if (MapProviderStr.Equals("OviMap")) return GMap.NET.MapProviders.GMapProviders.OviMap;
            if (MapProviderStr.Equals("OviTerrainMap")) return GMap.NET.MapProviders.GMapProviders.OviTerrainMap;
            if (MapProviderStr.Equals("YandexMap")) return GMap.NET.MapProviders.GMapProviders.YandexMap;
            if (MapProviderStr.Equals("ArcGIS_World_Street_Map")) return GMap.NET.MapProviders.GMapProviders.ArcGIS_World_Street_Map;
            if (MapProviderStr.Equals("ArcGIS_World_Topo_Map")) return GMap.NET.MapProviders.GMapProviders.ArcGIS_World_Topo_Map;

            if (MapProviderStr.Equals("BingSatelliteMap")) return GMap.NET.MapProviders.GMapProviders.BingSatelliteMap;
            if (MapProviderStr.Equals("GoogleSatelliteMap")) return GMap.NET.MapProviders.GMapProviders.GoogleSatelliteMap;
            if (MapProviderStr.Equals("OviSatelliteMap")) return GMap.NET.MapProviders.GMapProviders.OviSatelliteMap;
            if (MapProviderStr.Equals("YandexSatelliteMap")) return GMap.NET.MapProviders.GMapProviders.YandexSatelliteMap;
            if (MapProviderStr.Equals("ArcGIS_Imagery_World_2D_Map")) return GMap.NET.MapProviders.GMapProviders.ArcGIS_Imagery_World_2D_Map;

            if (MapProviderStr.Equals("BingHybridMap")) return GMap.NET.MapProviders.GMapProviders.BingHybridMap;
            if (MapProviderStr.Equals("GoogleHybridMap")) return GMap.NET.MapProviders.GMapProviders.GoogleHybridMap;
            if (MapProviderStr.Equals("OviHybridMap")) return GMap.NET.MapProviders.GMapProviders.OviHybridMap;
            if (MapProviderStr.Equals("YandexHybridMap")) return GMap.NET.MapProviders.GMapProviders.YandexHybridMap;
            return GMap.NET.MapProviders.GMapProviders.EmptyProvider;
        }
        //---------------------------

        public FormMapGeoDraw(string constr, int permissionAdmin_parent)
        {
            InitializeComponent();

            //MapSrc = null;

            xmlDataSet = new DataSet();

            mapSrcDictionary = new Dictionary<int, MC_MapSource>();
            layerVisibleDictionary = new Dictionary<int, CheckState>();

            dbconnectionString = constr;

            SQLDataProvider = new MC_SQLDataProvider(dbconnectionString);
            MapLegend = new MC_MapLegend(SQLDataProvider);

            permissionAdmin = permissionAdmin_parent;

            // настройка интерфейса в зависимости от прав доступа
            //ribbonPageGroupOptions.Visible = (permissionAdmin == 100);
            ribbonPageGroupOptions.Visible = (permissionAdmin == 100 || permissionAdmin == 2); // добавление прав для редактирования легенды для ОДС

            ribbonPageGroupImport.Visible = (permissionAdmin != 1);

            if (permissionAdmin == 100) // если администратор
            {
                dockPanel4.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible; // панель отладки    
            }            
            else
            {
                dockPanel4.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }

            // заполняем таблицу данных о легенде карты            
            /*string queryString1 = "SELECT * FROM objectsoke.dbo.tblPanoramaLegend ORDER BY idpnrmGROUPLAYER ASC";
            tableLEGEND = new DataTable();
            MyFuncSQL.SelectDataFromSQL(tableLEGEND, dbconnectionString, queryString1);*/

            formShowLegend = null;
            formShowLegend = new FormLegendShow();
            formShowLegend.Text = "Легенда карты";
            //----------------------------------------
                        
            // для XML создаем таблицы слоев и типов объектов
            SqlConnection SQLconnection = new SqlConnection(dbconnectionString);
            SQLconnection.Open();

            string queryStringObj =
                String.Concat("SELECT idpnrmGROUPLAYER,pnrmLAYERcaptsrc,pnrmGROUPLAYERcapt FROM objectsoke.dbo.tblPanoramaGroupLayer " + 
                              "WHERE idpnrmGROUPLAYER <> 43"); // ПЕРЕДЕЛАТЬ!!!
            xmlTableLAYER = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, xmlTableLAYER, queryStringObj);

            queryStringObj =
                String.Concat("SELECT idpnrmLOCALtype,pnrmLOCALtypecapt FROM objectsoke.dbo.tblPanoramaLOCALType");
            xmlTableLOCAL = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(SQLconnection, xmlTableLOCAL, queryStringObj);

            SQLconnection.Close();

            srcFileType = -1;
        }

        private void FormMapGeoDraw_Load(object sender, EventArgs e)
        {
            // подправить
            string queryString = 
                String.Concat("select tblMapSource.idmapsrc, tblMapSource.captionmapsrc, tblMapSource.idcontractor, tblMapSource.idfilial, ",
                "tblMapSource.commentmapsrc, tblMapSource.datecreated, tblMapSource.dateloaded, tblFilial.captionfilial, tblContractor.captioncontr ",
                "from((dbo.tblMapSource tblMapSource ",
                "inner join dbo.tblFilial tblFilial on(tblFilial.idfilial = tblMapSource.idfilial)) ",
                "inner join dbo.tblContractor tblContractor on(tblContractor.idcontractor = tblMapSource.idcontractor))");
            DataTable table = new DataTable();            
            MC_SQLDataProvider.SelectDataFromSQL(table, dbconnectionString, queryString);

            string queryStringLAYERname =
                //String.Concat("SELECT idpnrmGROUPLAYER,pnrmLAYERcaptsrc,pnrmGROUPLAYERcapt ",
                String.Concat("SELECT idpnrmGROUPLAYER ",
                "FROM objectsoke.dbo.tblPanoramaGroupLayer ",
                " ORDER BY idpnrmGROUPLAYER ASC");
            DataTable tableGROUPLAYERname = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableGROUPLAYERname, dbconnectionString, queryStringLAYERname);

            // заполняем дерево электронных карт----------------------------
            DataSetPnrmMapSrc DataSetLoad = new DataSetPnrmMapSrc();            
            DataSetPnrmMapSrcTableAdapters.tblFilialTableAdapter tblFilialTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblFilialTableAdapter();
            tblFilialTableAdapter.Fill(DataSetLoad.tblFilial);

            treeViewFilials.Nodes.Clear();
            for (int i = 0; i < DataSetLoad.tblFilial.Rows.Count; i++)
            {
                string captionfilial = DataSetLoad.tblFilial.Rows[i]["captionfilial"].ToString();

                DataRow[] mapsrcrows = table.Select(String.Concat("captionfilial = '", captionfilial, "'"));

                if (mapsrcrows.Count() > 0)
                {
                    TreeNode[] nodesmapsrc = new TreeNode[mapsrcrows.Count()];
                    for (int j = 0; j < mapsrcrows.Count(); j++)
                    {
                        string mapsrcdatecreatedstr = "";
                        if (!String.IsNullOrWhiteSpace(mapsrcrows[j]["datecreated"].ToString()))
                        {
                            DateTime mapsrcdatecreated = Convert.ToDateTime(mapsrcrows[j]["datecreated"].ToString());
                            mapsrcdatecreatedstr = mapsrcdatecreated.ToShortDateString();
                        }
                        string captionnodemapsrc = 
                            String.Concat(mapsrcrows[j]["captionmapsrc"].ToString(), " (", mapsrcrows[j]["captioncontr"].ToString(), 
                                ", ", mapsrcdatecreatedstr, ")");
                        nodesmapsrc[j] = new TreeNode(captionnodemapsrc);
                        nodesmapsrc[j].Tag = mapsrcrows[j];
                    }

                    TreeNode nodefilial = new TreeNode(captionfilial, nodesmapsrc);
                    nodefilial.Tag = "filialcaption";

                    treeViewFilials.Nodes.Add(nodefilial);                    
                }
            } // for (int i = 0; i < DataSetLoad.tblFilial.Rows.Count; i++)
            treeViewFilials.Sort();          
            //----------------------------------
            
            //Настройки для компонента GMap.
            gMapControl1.Bearing = 0;

            //CanDragMap - Если параметр установлен в True,
            //пользователь может перетаскивать карту
            ///с помощью правой кнопки мыши.
            gMapControl1.CanDragMap = true;

            //Указываем, что перетаскивание карты осуществляется
            //с использованием левой клавишей мыши.
            //По умолчанию - правая.
            gMapControl1.DragButton = MouseButtons.Left;

            gMapControl1.GrayScaleMode = true;
            
            gMapControl1.MarkersEnabled = true;

            //Указываем значение максимального приближения.
            gMapControl1.MaxZoom = 36;// 18;

            //Указываем значение минимального приближения.
            gMapControl1.MinZoom = 2;

            //Устанавливаем центр приближения/удаления
            //курсор мыши.
            gMapControl1.MouseWheelZoomType = GMap.NET.MouseWheelZoomType.MousePositionAndCenter;

            //Отказываемся от негативного режима.
            gMapControl1.NegativeMode = false;

            //Разрешаем полигоны.
            gMapControl1.PolygonsEnabled = true;

            //Разрешаем маршруты
            gMapControl1.RoutesEnabled = true;

            //Скрываем внешнюю сетку карты
            //с заголовками.
            gMapControl1.ShowTileGridLines = false;

            //Указываем, что при загрузке карты будет использоваться
            //6х кратное приближение.
            gMapControl1.Zoom = 6;

            // инициализация панелей выбора "подложки"
            listBoxMapProvider1.Hide();
            listBoxMapProvider2.Hide();
            listBoxMapProvider3.Hide();

            listBoxMapProvider1.Items.Clear();
            listBoxMapProvider2.Items.Clear();
            listBoxMapProvider3.Items.Clear();

            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.EmptyProvider);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.ArcGIS_World_Street_Map);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.ArcGIS_World_Topo_Map);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.GoogleMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.GoogleTerrainMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OpenCycleLandscapeMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OpenCycleMap);            
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OpenCycleTransportMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OpenSeaMapHybrid);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OpenStreetMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OpenStreetMapQuest);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OviMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.OviTerrainMap);
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.WikiMapiaMap);                        
            listBoxMapProvider1.Items.Add(GMap.NET.MapProviders.GMapProviders.YandexMap);

            listBoxMapProvider2.Items.Add(GMap.NET.MapProviders.GMapProviders.EmptyProvider);
            listBoxMapProvider2.Items.Add(GMap.NET.MapProviders.GMapProviders.ArcGIS_Imagery_World_2D_Map);
            listBoxMapProvider2.Items.Add(GMap.NET.MapProviders.GMapProviders.BingSatelliteMap);
            listBoxMapProvider2.Items.Add(GMap.NET.MapProviders.GMapProviders.GoogleSatelliteMap);
            listBoxMapProvider2.Items.Add(GMap.NET.MapProviders.GMapProviders.OviSatelliteMap);
            listBoxMapProvider2.Items.Add(GMap.NET.MapProviders.GMapProviders.YandexSatelliteMap);

            listBoxMapProvider3.Items.Add(GMap.NET.MapProviders.GMapProviders.EmptyProvider);
            listBoxMapProvider3.Items.Add(GMap.NET.MapProviders.GMapProviders.BingHybridMap);
            listBoxMapProvider3.Items.Add(GMap.NET.MapProviders.GMapProviders.GoogleHybridMap);
            listBoxMapProvider3.Items.Add(GMap.NET.MapProviders.GMapProviders.OviHybridMap);
            listBoxMapProvider3.Items.Add(GMap.NET.MapProviders.GMapProviders.YandexHybridMap);            
            
            //GMap.NET.MapProviders.GMapProvider[] mp1 = new 
            /*foreach (GMap.NET.MapProviders.GMapProvider GMapProvider in GMap.NET.MapProviders.GMapProviders.List)
            {
                listBoxMapProvider1.Items.Add(GMapProvider.ToString());                
            }*/

            listBoxMapProvider1.Show();
            listBoxMapProvider2.Show();
            listBoxMapProvider3.Show();
            //--------------------------------------- 

            //GMap.NET.MapProviders.GMapProvider.UserAgent = "Oblkommunenergo";
            gMapControl1.MapProvider = GMap.NET.MapProviders.GMapProviders.GoogleMap;            
            listBoxMapProvider1.SelectedIndex = 3;
            GMap.NET.GMaps.Instance.Mode = GMap.NET.AccessMode.ServerOnly;

            gMapControl1.Position = new PointLatLng(54.5368793388085, 103.579102531985);

            //Если вы используете интернет через прокси сервер,
            //указываем свои учетные данные.
            /*GMap.NET.MapProviders.GMapProvider.WebProxy = new System.Net.WebProxy("192.168.3.241",true);
            GMap.NET.MapProviders.GMapProvider.WebProxy.Credentials = new System.Net.NetworkCredential("", "");*/

            GMap.NET.MapProviders.GMapProvider.WebProxy = System.Net.WebRequest.GetSystemWebProxy();
            //GMap.NET.MapProviders.GMapProvider.WebProxy.Credentials = System.Net.CredentialCache.DefaultCredentials;
            GMap.NET.MapProviders.GMapProvider.WebProxy.Credentials = System.Net.CredentialCache.DefaultNetworkCredentials;

            /*foreach (DataRow layerRow in layerRows)
            {
                checkedListBoxControl1.Items.Add(layerRow["layer"]).ToString();
            }*/

            // инициализация пользовательских свойств класса --------------------------------------- 
            //current_idmapsrc = -1;
            nodemapsrcloaded = null;

            //gMapControl1.ScalePen.Color = Color.White;
            //gMapControl1.Font.S            
            //gMapControl1.MapScaleInfoEnabled = true;

            // начальная статистика
            this.barButtonMapProvider.Caption = "География: " + gMapControl1.MapProvider.ToString();

            //--------------------------            
            // заполняем gmc "все слоями"
            foreach (DataRow rowlayer in tableGROUPLAYERname.Rows)
            {
                int layerid = (int)rowlayer["idpnrmGROUPLAYER"];
                //string layercapt = rowlayer["pnrmGROUPLAYERcapt"].ToString();

                GMapOverlay polyOverlay = new GMapOverlay(layerid.ToString());
                polyOverlay.IsVisibile = false;
                gMapControl1.Overlays.Add(polyOverlay);

                layerVisibleDictionary.Add(layerid, CheckState.Unchecked);
            }
        } // private void FormMapGeoDraw_Load(object sender, EventArgs e)
        
        // нажатие на маркер
        private void gMapControl1_OnMarkerClick(GMap.NET.WindowsForms.GMapMarker item, MouseEventArgs e)
        {
            //ShowObjectSEMvalues((int)item.Tag);
            ShowObjectSEMvalues((string)item.Tag);
        }

        // нажатие на полигон
        private void gMapControl1_OnPolygonClick(GMapPolygon item, MouseEventArgs e)
        {
            ShowObjectSEMvalues((string)item.Tag);            
        }

        // нажатие на route
        private void gMapControl1_OnRouteClick(GMapRoute item, MouseEventArgs e)
        {
            //ShowObjectSEMvalues((string)item.Tag);

            /*// "бубен" - накладывающиеся в ОДНОМ И ТОМ ЖЕ слое объекты не "суммируются" по нажатию на них
            foreach (GMapOverlay overlay1 in gMapControl1.Overlays)
            {
                foreach (GMapRoute route1 in overlay1.Routes)
                {
                    if (route1.IsMouseOver)
                    {
                        ShowObjectSEMvalues((string)route1.Tag);
                    }
                }
            } // конец "бубна"*/
        }

        // загрузка данных карты
        public void LoadEraseDataFromMapSrc(bool isLOAD, int idmapsrc)
        {
            splashScreenManager1.ShowWaitForm();
            
            /*SqlConnection SQLconnection = new SqlConnection(dbconnectionString);
            SQLconnection.Open();*/
            
            //диагностика
            System.Diagnostics.Stopwatch swatch = new System.Diagnostics.Stopwatch();
            
            // запоминаем состояние checkbox-а слоев
            for (int i = 0; i < checkedListBoxControl1.Items.Count; i++)
            {
                layerVisibleDictionary[(int)checkedListBoxControl1.Items[i].Value] = checkedListBoxControl1.Items[i].CheckState;
            }
            //--------------------------------------

            MC_MapSource MapSrc;
            if (isLOAD) // загрузка
            {
                swatch.Start();

                MapSrc = new MC_MapSource(idmapsrc);

                swatch.Stop();
                memoEditSemantic.Text += "load - " + swatch.Elapsed.ToString() + Environment.NewLine;
                swatch.Reset();

                mapSrcDictionary.Add(idmapsrc, MapSrc);

                swatch.Start();
                MapSrc.DrawMapSource(gMapControl1);
                swatch.Stop();
                memoEditSemantic.Text += "paint - " + swatch.Elapsed.ToString() + Environment.NewLine;

                gridControl1.DataSource = MapSrc.tableObj;
            }
            else // удаление
            {
                MapSrc = mapSrcDictionary[idmapsrc];                        
                MapSrc.EraseMapSource(gMapControl1);
                mapSrcDictionary.Remove(idmapsrc);
            }
            
            //----------------------------
                                    
            this.checkedListBoxControl1.Items.Clear();
                        
            formShowLegend.dataGridView1.Rows.Clear();
            formShowLegend.dataGridView2.Rows.Clear();
            formShowLegend.dataGridView3.Rows.Clear();

            // объединяем таблицы слоев      
            DataTable tableLAYERname = new DataTable();
            DataTable tableLEGENDmain = new DataTable();
            //IEnumerable<DataRow> layerLEGENDrows = new DataTable().AsEnumerable(); // пустая начальная таблица

            if (mapSrcDictionary.Count > 0)
            {
                string idmapsrcENUM = "(";
                int i = 0;
                foreach (var mapSrcInDict in mapSrcDictionary)
                {
                    //idmapsrcENUM = String.Concat(idmapsrcENUM, mapSrcInDict.Value.IdMapSrc.ToString());
                    idmapsrcENUM += mapSrcInDict.Value.IdMapSrc.ToString();
                    i++;
                    if (i < mapSrcDictionary.Count) idmapsrcENUM += ",";

                    // для легенды
                    //layerLEGENDrows = layerLEGENDrows.Union(mapSrcInDict.Value.tableLEGEND.AsEnumerable());
                }
                idmapsrcENUM += ")";

                // попробовать через коллекции для сравнения
                string queryStringLAYERname =
                    String.Concat("SELECT DISTINCT tblPO.idpnrmGROUPLAYER, tblPGL.pnrmGROUPLAYERcapt ",
                    "FROM objectsoke.dbo.tblPanoramaObject tblPO ",
                    "INNER JOIN objectsoke.dbo.tblPanoramaGroupLayer tblPGL ON tblPO.idpnrmGROUPLAYER = tblPGL.idpnrmGROUPLAYER ",
                    "WHERE tblPO.idmapsrc IN ", idmapsrcENUM,
                    " ORDER BY tblPGL.pnrmGROUPLAYERcapt ASC");                
                MC_SQLDataProvider.SelectDataFromSQL(tableLAYERname, dbconnectionString, queryStringLAYERname);

                //layerLEGENDrows = layerLEGENDrows.Distinct(new DataRowComparer<DataRow>());
                string queryString1 = String.Concat(
                "SELECT * FROM objectsoke.dbo.tblPanoramaLegend",
                //" WHERE idpnrmGROUPLAYER IN ", idmapsrcENUM,                    
                " ORDER BY idpnrmGROUPLAYER ASC");                
                MC_SQLDataProvider.SelectDataFromSQL(tableLEGENDmain, dbconnectionString, queryString1);
            } // if (mapSrcDictionary.Count > 0)

            // добавляем служебный слой XML
            checkedListBoxControl1.Items.Add(43, "ЗАГРУЖЕНО", layerVisibleDictionary[43], true); // ИСПРАВИТЬ
            //-----------------------------

            //foreach (DataRow rowlayer in MapSrc.tableLAYERname.Rows)
            foreach (DataRow rowlayer in tableLAYERname.Rows)
            {
                int layerid = (int)rowlayer["idpnrmGROUPLAYER"];
                string layerid_str = layerid.ToString();
                string layercapt = rowlayer["pnrmGROUPLAYERcapt"].ToString();
                //checkedListBoxControl1.Items.Add(layercapt);
                checkedListBoxControl1.Items.Add(layerid, layercapt, layerVisibleDictionary[layerid], true);

                /*// посмотреть!
                polyOverlay = new GMapOverlay(layerid);
                polyOverlay.IsVisibile = false;
                gMapControl1.Overlays.Add(polyOverlay);*/

                // создание легенды карты
                //layerLEGENDrows = MapSrc.tableLEGEND.Select(String.Concat("idpnrmGROUPLAYER = ", layerid));
                DataRow[] layerLEGENDrows = tableLEGENDmain.Select(String.Concat("idpnrmGROUPLAYER = ", layerid_str));

                foreach (DataRow layerLEGENDrow in layerLEGENDrows)
                {
                    int idpnrmLOCALtype = (int)layerLEGENDrow["idpnrmLOCALtype"];

                    // переделать на заполнение из БД!!! + названия вкладок (стандарт, спутник, гибрид)
                    string pnrmLOCALtypecapt = string.Empty;
                    if (idpnrmLOCALtype == 2) pnrmLOCALtypecapt = "линейный";
                    else if (idpnrmLOCALtype == 3) pnrmLOCALtypecapt = "площадной";
                    else if (idpnrmLOCALtype == 4) pnrmLOCALtypecapt = "подпись";
                    else if (idpnrmLOCALtype == 5) pnrmLOCALtypecapt = "точечный";

                    int idlgnddest = (int)layerLEGENDrow["idlgnddest"];

                    Image legendRowImg = new Bitmap(50, 20);
                    Graphics g = Graphics.FromImage(legendRowImg);

                    // формируем стиль контура
                    int line_width = (int)layerLEGENDrow["PENthickness"];
                    Pen layerPen = new Pen(Color.FromArgb((int)layerLEGENDrow["PENcolor"]), line_width);
                    int pen_style = (int)layerLEGENDrow["PENstyle"];
                    if (pen_style == 0) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;// сделать через массив!!!
                    else if (pen_style == 1) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDot;
                    else if (pen_style == 2) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDotDot;
                    else if (pen_style == 3) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                    else if (pen_style == 4) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;

                    // формируем стиль заливки
                    Brush layerBrush = new SolidBrush(Color.FromArgb((int)layerLEGENDrow["FILLSOLIDalpha"], Color.FromArgb((int)layerLEGENDrow["FILLSOLIDcolor"])));

                    // стиль точечного объекта
                    int point_type = (int)layerLEGENDrow["POINTtype"];

                    // отрисовка легенды в зависимости от типа объекта
                    int dx = 5;
                    switch (idpnrmLOCALtype)
                    {
                        case 2: // линейный 
                            g.DrawLine(layerPen, new Point(dx, (legendRowImg.Height - line_width) / 2), new Point(legendRowImg.Width - dx, (legendRowImg.Height - line_width) / 2));
                            break;

                        case 3: // площадной
                            Rectangle pointrect = new Rectangle(dx, dx, legendRowImg.Width - dx - 5, legendRowImg.Height - dx - 5);
                            g.FillRectangle(layerBrush, pointrect);
                            g.DrawRectangle(layerPen, pointrect);
                            break;

                        case 4: // подпись
                            break;

                        case 5: // точечный
                            dx = 8;
                            pointrect = new Rectangle((legendRowImg.Width - dx) / 2, (legendRowImg.Height - dx) / 2, dx, dx);
                            switch (point_type)
                            {
                                case 0: // квадрат
                                    g.FillRectangle(layerBrush, pointrect);
                                    g.DrawRectangle(layerPen, pointrect);
                                    break;
                                case 1: // круг
                                    g.FillEllipse(layerBrush, pointrect);
                                    g.DrawEllipse(layerPen, pointrect);
                                    break;
                            }
                            break;
                    }

                    layerPen.Dispose();
                    layerBrush.Dispose();
                    g.Dispose();
                    //------------------------------------------------                   
                    
                    if (idlgnddest == 1)
                        formShowLegend.dataGridView1.Rows.Add(layercapt, pnrmLOCALtypecapt, legendRowImg);
                    else if (idlgnddest == 2)
                        formShowLegend.dataGridView2.Rows.Add(layercapt, pnrmLOCALtypecapt, legendRowImg);
                    else if (idlgnddest == 3)
                        formShowLegend.dataGridView3.Rows.Add(layercapt, pnrmLOCALtypecapt, legendRowImg);
                } // foreach (DataRow layerLEGENDrow in layerLEGENDrows)

                formShowLegend.dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
            }
            //---------------------------------------------------------------------------------------------------------

            //SQLconnection.Close();

            tableLAYERname.Dispose();
            tableLEGENDmain.Dispose();

            splashScreenManager1.CloseWaitForm();
        } // public void LoadEraseDataFromMapSrc(bool isLOAD, int idmapsrc)
                
        // check layers
        private void checkedListBoxControl1_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            /*if (e.State == CheckState.Checked) gMapControl1.Overlays[e.Index].IsVisibile = true;
            else gMapControl1.Overlays[e.Index].IsVisibile = false;            
            gMapControl1.Refresh();*/

            int idOverlay = (int)(checkedListBoxControl1.Items[e.Index].Value);
            layerVisibleDictionary[idOverlay] = e.State;

            GMapOverlay overlayCurrent = MC_GMap.GetOverlayIndexByName(gMapControl1, idOverlay.ToString());
            if (e.State == CheckState.Checked) overlayCurrent.IsVisibile = true;
            else overlayCurrent.IsVisibile = false;
            gMapControl1.Refresh();
        }
                
        private void gMapControl1_Load(object sender, EventArgs e)
        {
            
        }

        // zoom+
        private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
        {
            gMapControl1.Zoom++;
            //gMapControl1.Zoom += 0.1;
        }

        // zoom-
        private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
        {
            gMapControl1.Zoom--;
            //gMapControl1.Zoom -= 0.1;
        }

        // процедура загрузки карты по двойному нажатию в Навигаторе
        private void LoadDataFromMapSrcByClick()
        {
            /*if (treeViewFilials.SelectedNode != null)
            {
                TreeNode selectednode = treeViewFilials.SelectedNode;
                
                if (selectednode.Tag as string != "filialcaption")
                {
                    DataRow selectedrow = selectednode.Tag as DataRow;
                    //int idmapsrc = Convert.ToInt32(selectedrow["idmapsrc"].ToString());
                    int idmapsrc = (int)selectedrow["idmapsrc"];
                    if (current_idmapsrc != idmapsrc)
                    {
                        Text = selectedrow["captionmapsrc"].ToString();

                        if (current_idmapsrc != -1)
                        {
                            nodemapsrcloaded.NodeFont = new Font(this.Font, 0);
                            nodemapsrcloaded.ForeColor = Color.Black;
                        }

                        selectednode.NodeFont = new Font(this.Font, FontStyle.Bold);
                        selectednode.ForeColor = Color.Blue;
                        selectednode.Checked = true;

                        nodemapsrcloaded = selectednode;
                        //----------------------------------
                        LoadDataFromMapSrc(idmapsrc);                        
                    }
                    else MessageBox.Show("Данная карта уже загружена");
                }
            }*/
        }

        private void LoadDataFromMapSrcByClick(TreeNode checkednode)
        {               
            if (checkednode.Tag as string != "filialcaption")
            {
                DataRow selectedrow = checkednode.Tag as DataRow;                    
                int idmapsrc = (int)selectedrow["idmapsrc"];
                
                /*Text = selectedrow["captionmapsrc"].ToString();
                                        
                if (mapSrcDictionary.Count > 0)
                {
                    nodemapsrcloaded.NodeFont = new Font(this.Font, 0);
                    nodemapsrcloaded.ForeColor = Color.Black;
                }

                checkednode.NodeFont = new Font(this.Font, FontStyle.Bold);
                checkednode.ForeColor = Color.Blue;
                
                nodemapsrcloaded = checkednode;*/

                if (checkednode.Checked)
                {
                    //DeleteDataFromMapSrc(idmapsrc);
                    LoadEraseDataFromMapSrc(false, idmapsrc);

                    //checkednode.NodeFont = new Font(this.Font, 0);
                    checkednode.ForeColor = Color.Black;
                }
                else
                {
                    LoadEraseDataFromMapSrc(true, idmapsrc);

                    //checkednode.NodeFont = new Font(this.Font, FontStyle.Bold);
                    checkednode.ForeColor = Color.Blue;

                    Text = selectedrow["captionmapsrc"].ToString(); // подправить на случай множественного выбора
                }
            }       
        }

        private void gMapControl1_Paint(object sender, PaintEventArgs e)
        {
            this.barButtonStatusZoom.Caption = "Масштаб: " + gMapControl1.Zoom.ToString();
            //memoEditSemantic.Text += gMapControl1.Position.ToString();
        }

        // выбор провайдера-подложки
        private void listBoxMapProvider1_SelectedValueChanged(object sender, EventArgs e)
        {
            if ((sender as DevExpress.XtraEditors.ListBoxControl).SelectedIndex != -1)
            {
                gMapControl1.MapProvider = ReturnGMapProviderByString((sender as DevExpress.XtraEditors.ListBoxControl).SelectedValue.ToString()); /// ???

                // исправляем косяк со смещением слоев при изменении подложки ---------------
                for (int i = 0; i < gMapControl1.Overlays.Count; i++)
                {
                    bool visible_status_prev = gMapControl1.Overlays[i].IsVisibile;
                    gMapControl1.Overlays[i].IsVisibile = false;
                    gMapControl1.Overlays[i].IsVisibile = visible_status_prev;
                }
                //-------------------------------

                this.barButtonMapProvider.Caption = "География: " + gMapControl1.MapProvider.ToString();
            }
        }

        private void listBoxMapProvider2_SelectedValueChanged(object sender, EventArgs e)
        {
            if ((sender as DevExpress.XtraEditors.ListBoxControl).SelectedIndex != -1)
            {
                gMapControl1.MapProvider = ReturnGMapProviderByString((sender as DevExpress.XtraEditors.ListBoxControl).SelectedValue.ToString());


                // исправляем косяк со смещением слоев при изменении подложки ---------------
                for (int i = 0; i < gMapControl1.Overlays.Count; i++)
                {
                    bool visible_status_prev = gMapControl1.Overlays[i].IsVisibile;
                    gMapControl1.Overlays[i].IsVisibile = false;
                    gMapControl1.Overlays[i].IsVisibile = visible_status_prev;
                }
                //-------------------------------

                this.barButtonMapProvider.Caption = "География: " + gMapControl1.MapProvider.ToString();
            }
        }

        private void listBoxMapProvider3_SelectedValueChanged(object sender, EventArgs e)
        {
            if ((sender as DevExpress.XtraEditors.ListBoxControl).SelectedIndex != -1)
            {
                gMapControl1.MapProvider = ReturnGMapProviderByString((sender as DevExpress.XtraEditors.ListBoxControl).SelectedValue.ToString());

                // исправляем косяк со смещением слоев при изменении подложки ---------------
                for (int i = 0; i < gMapControl1.Overlays.Count; i++)
                {
                    bool visible_status_prev = gMapControl1.Overlays[i].IsVisibile;
                    gMapControl1.Overlays[i].IsVisibile = false;
                    gMapControl1.Overlays[i].IsVisibile = visible_status_prev;
                }
                //-------------------------------

                this.barButtonMapProvider.Caption = "География: " + gMapControl1.MapProvider.ToString();
            }
        }

        // сохранить изображение
        private void barButtonItem7_ItemClick(object sender, ItemClickEventArgs e)
        {               
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                gMapControl1.ToImage().Save(saveFileDialog1.FileName, System.Drawing.Imaging.ImageFormat.Tiff);
            }
        }

        // открытие формы просмотра/редактирования легенды
        private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormLayerLegendEdit form1 = null;
            form1 = new FormLayerLegendEdit(this.dbconnectionString, permissionAdmin);
            
            form1.ShowDialog();
        }
        

        // обработчик события "изменение активной страницы"
        private void xtraTabControl5_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            // сбрасываем селекторы ГИС-подложек
            listBoxMapProvider1.SelectedIndex = -1;
            listBoxMapProvider2.SelectedIndex = -1;
            listBoxMapProvider3.SelectedIndex = -1;
        }

        // Легенда карты
        private void barButtonItemLegendShow_ItemClick(object sender, ItemClickEventArgs e)
        {
            /*FormLegendShow form1 = null;
            form1 = new FormLegendShow(this.dbconnectionString);

            form1.Show();*/
            formShowLegend.Show();
        }

        // popup загрузить mapsrc
        private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
        {
            LoadDataFromMapSrcByClick();
        }

        // popup свойства mapsrc
        private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
        {            
            if (treeViewFilials.SelectedNode != null)
            {
                TreeNode selectednode = treeViewFilials.SelectedNode;
                DataRow selectedrow = selectednode.Tag as DataRow;
                MessageBox.Show("idmapsrc"+selectedrow["idmapsrc"].ToString());
            }
        }

        // двойное нажатие в окне филиал-карты
        private void treeViewFilials_DoubleClick(object sender, EventArgs e)
        {
            //barButtonItem3_ItemClick(sender, e as ItemClickEventArgs);                   
        }

        private void treeViewFilials_BeforeCheck(object sender, TreeViewCancelEventArgs e)
        {            
            LoadDataFromMapSrcByClick(e.Node); 
        }

        // загружаем из XML
        private void barButtonItem5_ItemClick(object sender, ItemClickEventArgs e)
        {
            //xmlDataSet = new DataSet();
            openFileDialogXML.Filter = "gpx (*.gpx)|*.gpx";

            if (openFileDialogXML.ShowDialog() == DialogResult.OK)
            {
                srcFileType = 0;
                                
                foreach (String file in openFileDialogXML.FileNames)
                {
                    //xmlDataSet.ReadXmlSchema(file);
                    xmlDataSet.ReadXml(file);
                }

                xmlTable = xmlDataSet.Tables["wpt"];

                /*xmlDataSet.ReadXml("e:\\2\\point17.gpx");
                xmlTable = xmlDataSet.Tables["wpt"];*/

                // удаляем все пробелы в названиях точек                
                foreach (DataRow xmlRow in xmlTable.Rows)
                {
                    xmlRow["name"] = xmlRow["name"].ToString().Replace(" ", "");
                }                
                //gridControl1.DataSource = xmlTable; // для отладки

                // загружаем данные в XML-слой
                GMapOverlay overlayCurrent = MC_GMap.GetOverlayIndexByName(gMapControl1, "43"); // исправить!!!

                // бежим по объектам
                //foreach (DataRow objRow in tableObjRows)
                foreach (DataRow xmlRow in xmlTable.Rows)
                {
                    //int idpnrmOBJECT = 900000 + Convert.ToInt32(xmlRow["name"].ToString());
                    //string idpnrmOBJECTstr = String.Concat("M", IdMapSrc.ToString(), "-", idpnrmOBJECT.ToString());
                    string idpnrmOBJECTstr = xmlRow["name"].ToString();

                    //int idpnrmLOCALtype = 5; // точечный ИСПРАВИТЬ!!!

                    double xmlLAT = Convert.ToDouble(xmlRow["lat"].ToString().Replace('.', ','));
                    double xmlLON = Convert.ToDouble(xmlRow["lon"].ToString().Replace('.', ','));

                    //GMap.NET.WindowsForms.Markers.GMarkerGoogleType marker_type = GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red;
                    GMap.NET.WindowsForms.Markers.GMarkerGoogleType marker_type = GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small;
                    GMap.NET.WindowsForms.Markers.GMarkerGoogle markerXML = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(new GMap.NET.PointLatLng(xmlLAT, xmlLON), marker_type);
                    markerXML.IsHitTestVisible = false;

                    markerXML.ToolTip = new GMap.NET.WindowsForms.ToolTips.GMapRoundedToolTip(markerXML);
                    //markerXML.ToolTip = new GMap.NET.WindowsForms.ToolTips.GMapBaloonToolTip(markerXML);
                    markerXML.ToolTipMode = GMap.NET.WindowsForms.MarkerTooltipMode.Always;
                    markerXML.ToolTipText = idpnrmOBJECTstr;

                    markerXML.Tag = idpnrmOBJECTstr;

                    overlayCurrent.Markers.Add(markerXML);

                } // foreach (DataRow xmlRow in xmlTable.Rows)

                overlayCurrent.IsVisibile = true;
                layerVisibleDictionary[43] = CheckState.Checked;
                checkedListBoxControl1.Items[0].CheckState = CheckState.Checked; // ПЕРЕДЕЛАТЬ!!! сейчас - первый элемент в списке 
                // добавляем служебный слой XML
                //checkedListBoxControl1.Items.Add(43, "ЗАГРУЖЕНО", layerVisibleDictionary[43], true); // ИСПРАВИТЬ                      
            } // if (openFileDialogXML.ShowDialog() == DialogResult.OK)
            //------------------------------------
        }

        // кнопка "Создание объекта из XML"
        private void barButtonItem8_ItemClick_1(object sender, ItemClickEventArgs e)
        {
            FormSetObjectFromXML form1 = null;
            form1 = new FormSetObjectFromXML();
            form1.formMain = this;

            // заполняем ComboBox-ы
            foreach (DataRow drow in xmlTableLAYER.Rows)
            {
                //form1.comboBoxEditLAYER.Properties.Items.Add(drow["pnrmGROUPLAYERcapt"]);
                form1.listBoxControlLAYER.Items.Add(drow["pnrmGROUPLAYERcapt"].ToString());
            }

            // забил вручную для "защиты от дурака" - чтобы не выбрали лишнего
            /*foreach (DataRow drow in xmlTableLOCAL.Rows)
            {
                form1.comboBoxEditLOCAL.Properties.Items.Add(drow["pnrmLOCALtypecapt"]);
            }*/

            form1.listBoxControlLAYER.SelectedIndex = 0;
            form1.listBoxControlLOCAL.SelectedIndex = 0;

            //form1.ShowDialog();

            if (form1.ShowDialog() == DialogResult.OK)
            {
                PointLatLng prevPosition = gMapControl1.Position;
                double prevZoom = gMapControl1.Zoom;

                LoadEraseDataFromMapSrc(false, 900);
                LoadEraseDataFromMapSrc(true, 900);

                gMapControl1.Position = prevPosition;
                gMapControl1.Zoom = prevZoom;
            } // if (form1.ShowDialog() == DialogResult.OK)
        }

        // удаление импортированных точек
        private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)
        {
            // очистка слоя
            GMapOverlay gmcOverlay = MC_GMap.GetOverlayIndexByName(gMapControl1, "43"); // исправить!!!

            for (int i = 0; i < gmcOverlay.Markers.Count; i++)
            {
                GMapMarker gmcObject = gmcOverlay.Markers[i];
                /*if (gmcObject.Tag.ToString().Contains(idMapSrcstr))
                {*/
                    gmcOverlay.Markers.Remove(gmcObject);
                    i--;
                //}
            }

            // очистка таблицы и датасета
            //xmlTable.Clear();
            xmlDataSet.Clear();

            xmlDataSet.Dispose();

            xmlDataSet = new DataSet();
        }

        // Тест загрузки файла marks.sml SAS.Planet
        private void barButtonItemSASPlanet_ItemClick(object sender, ItemClickEventArgs e)
        {
            //DataSet sasDataSet = new DataSet();
            //DataTable sasTable;// = new DataTable();
            
            openFileDialogXML.Filter = "sml (*.sml)|*.sml";

            if (openFileDialogXML.ShowDialog() == DialogResult.OK)
            {
                srcFileType = 1;

                foreach (String file in openFileDialogXML.FileNames)
                {
                    xmlDataSet.ReadXml(file);
                }

                xmlTable = xmlDataSet.Tables["ROW"];

                foreach (DataRow xmlRow in xmlTable.Rows)
                {
                    xmlRow["name"] = xmlRow["name"].ToString().Replace(" ", "");
                }

                /*xmlDataSet.ReadXml("e:\\2\\point17.gpx");
                xmlTable = xmlDataSet.Tables["wpt"];*/

                //gridControl1.DataSource = xmlTable;

                // загружаем данные в XML-слой
                GMapOverlay overlayCurrent = MC_GMap.GetOverlayIndexByName(gMapControl1, "43"); // исправить!!!

                // бежим по объектам
                //foreach (DataRow objRow in tableObjRows)
                foreach (DataRow xmlRow in xmlTable.Rows)
                {
                    //int idpnrmOBJECT = 900000 + Convert.ToInt32(xmlRow["name"].ToString());
                    //string idpnrmOBJECTstr = String.Concat("M", IdMapSrc.ToString(), "-", idpnrmOBJECT.ToString());
                    string idpnrmOBJECTstr = xmlRow["name"].ToString();

                    //int idpnrmLOCALtype = 5; // точечный ИСПРАВИТЬ!!!

                    double xmlLAT = Convert.ToDouble(xmlRow["latT"].ToString().Replace('.', ','));
                    double xmlLON = Convert.ToDouble(xmlRow["lonL"].ToString().Replace('.', ','));

                    //GMap.NET.WindowsForms.Markers.GMarkerGoogleType marker_type = GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red;
                    GMap.NET.WindowsForms.Markers.GMarkerGoogleType marker_type = GMap.NET.WindowsForms.Markers.GMarkerGoogleType.red_small;
                    GMap.NET.WindowsForms.Markers.GMarkerGoogle markerXML = new GMap.NET.WindowsForms.Markers.GMarkerGoogle(new GMap.NET.PointLatLng(xmlLAT, xmlLON), marker_type);
                    markerXML.IsHitTestVisible = false;

                    markerXML.ToolTip = new GMap.NET.WindowsForms.ToolTips.GMapRoundedToolTip(markerXML);
                    //markerXML.ToolTip = new GMap.NET.WindowsForms.ToolTips.GMapBaloonToolTip(markerXML);
                    markerXML.ToolTipMode = GMap.NET.WindowsForms.MarkerTooltipMode.Always;
                    markerXML.ToolTipText = idpnrmOBJECTstr;

                    markerXML.Tag = idpnrmOBJECTstr;

                    overlayCurrent.Markers.Add(markerXML);

                } // foreach (DataRow xmlRow in xmlTable.Rows)

                overlayCurrent.IsVisibile = true;
                layerVisibleDictionary[43] = CheckState.Checked;
                checkedListBoxControl1.Items[0].CheckState = CheckState.Checked; // ПЕРЕДЕЛАТЬ!!! сейчас - первый элемент в списке 
                // добавляем служебный слой XML
                //checkedListBoxControl1.Items.Add(43, "ЗАГРУЖЕНО", layerVisibleDictionary[43], true); // ИСПРАВИТЬ                      
            } // if (openFileDialogXML.ShowDialog() == DialogResult.OK)
            //------------------------------------
        }

        private void gMapControl1_MouseClick(object sender, MouseEventArgs e)
        {
            // "бубен" - накладывающиеся в ОДНОМ И ТОМ ЖЕ слое объекты не "суммируются" по нажатию на них
            foreach (GMapOverlay overlay1 in gMapControl1.Overlays)
            {
                foreach (GMapRoute route1 in overlay1.Routes)
                {
                    if (route1.IsMouseOver)
                    {
                        ShowObjectSEMvalues((string)route1.Tag);
                    }
                }
            } // конец "бубна"
        }

        // тестируем отображение WGS-координат под курсором мыши при перемещении по GMapControl-у
        private void gMapControl1_MouseMove(object sender, MouseEventArgs e)
        {
            PointLatLng currentPoint = gMapControl1.FromLocalToLatLng(e.X, e.Y);

            barTextMouseWGSCoords.Caption = String.Concat("Широта: ", currentPoint.Lat.ToString(), " Долгота: ", currentPoint.Lng.ToString());            
        }

        // сохранение в изображение выбранной территории (в т.ч. "за кадром")
        private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
        {
            /*PointLatLng pointLeftTop = new PointLatLng(52.2414401001363, 104.238924980164);
            PointLatLng pointRightBottom = new PointLatLng(52.2313216917683, 104.263343811035);*/

            RectLatLng areaSelected = gMapControl1.SelectedArea;
            PointLatLng pointLeftTop = areaSelected.LocationTopLeft;
            PointLatLng pointRightBottom = areaSelected.LocationRightBottom;

            double deltaLat = pointLeftTop.Lat - pointRightBottom.Lat;
            double deltaLng = pointRightBottom.Lng - pointLeftTop.Lng;

            // запоминаем предыдущие параметры
            PointLatLng prevPosition = gMapControl1.Position;
            double prevZoom = gMapControl1.Zoom;
            string prevCopyright = gMapControl1.MapProvider.Copyright;

            // "фотографируем"
            //gMapControl1.Hide();
            gMapControl1.ShowCenter = false;
            gMapControl1.MapProvider.Copyright = "";
            gMapControl1.SelectedArea = new RectLatLng(0, 0, 0, 0);

            RectLatLng fotoRect = gMapControl1.ViewArea; // запоминаем текущие параметры текущего Zoom-а для деления на кадры
            double half_lng_fotoRect = fotoRect.WidthLng / 2.0;
            double half_lat_fotoRect = fotoRect.HeightLat / 2.0;

            // подумать, что делать с последними столбцом и строкой кадров, т.к. они займут только часть кадра
            /*int countKadrLng = (int)(Math.Round(deltaLng / fotoRect.WidthLng); // кол-во столбцов
            int countKadrLat = (int)Math.Round(deltaLat / fotoRect.HeightLat); // кол-во строк*/

            double countKadrLngDbl = deltaLng / fotoRect.WidthLng;
            int countKadrLng = (int)Math.Round(countKadrLngDbl); // кол-во столбцов
            if (countKadrLngDbl > countKadrLng) countKadrLng++; // превращаем 3.2 в 4

            double countKadrLatDbl = deltaLat / fotoRect.HeightLat;
            int countKadrLat = (int)Math.Round(countKadrLatDbl); // кол-во строк
            if (countKadrLatDbl > countKadrLat) countKadrLat++;

            int kadrWidth = gMapControl1.Width;
            int kadrHeight = gMapControl1.Height;

            Bitmap fullImage = new Bitmap(kadrWidth * countKadrLng, kadrHeight * countKadrLat);
            Graphics g = Graphics.FromImage(fullImage);

            for (int row = 0; row < countKadrLat; row++ )
            {
                for (int col = 0; col < countKadrLng; col++)
                { 
                    PointLatLng pCenter = new PointLatLng(pointLeftTop.Lat - fotoRect.HeightLat * row - half_lat_fotoRect, 
                                                          pointLeftTop.Lng + fotoRect.WidthLng * col + half_lng_fotoRect);

                    gMapControl1.Position = pCenter;
                    Thread.Sleep(2000);

                    string fileName = String.Concat("c:\\3\\", (row+1).ToString(), (col+1).ToString(), ".tiff");
                    Image kadrImage = gMapControl1.ToImage();
                    //gMapControl1.ToImage().Save(fileName, System.Drawing.Imaging.ImageFormat.Tiff);

                    //Bitmap bm = new Bitmap(100, 100);//ваша нвоая картинка
                    //Bitmap bm1 = new Bitmap(30, 30);//ваша маленькая картинка

                    //Graphics g = Graphics.FromImage(bm);
                    //g.DrawImage(kadrImage, kadrWidth * col, kadrHeight * row, kadrImage.Width, kadrImage.Height);
                    g.DrawImage(kadrImage, kadrWidth * col, kadrHeight * row);
                    //g.Dispose();

                    // сделать задержку или по нажатию(предпочтительней)
                } // for (int col = 0; col < countKadrLng; col++)

            } // for (int row = 0; row < countKadrLat; row++ )

            // отображение легенды карты ВЫНЕСТИ В ОТД КЛАСС/МЕТОД
            // объединяем таблицы слоев      
            DataTable tableLAYERname = new DataTable();
            DataTable tableLEGENDmain = new DataTable();
            //IEnumerable<DataRow> layerLEGENDrows = new DataTable().AsEnumerable(); // пустая начальная таблица

            if (mapSrcDictionary.Count > 0)
            {
                string idmapsrcENUM = "(";
                int i = 0;
                foreach (var mapSrcInDict in mapSrcDictionary)
                {
                    //idmapsrcENUM = String.Concat(idmapsrcENUM, mapSrcInDict.Value.IdMapSrc.ToString());
                    idmapsrcENUM += mapSrcInDict.Value.IdMapSrc.ToString();
                    i++;
                    if (i < mapSrcDictionary.Count) idmapsrcENUM += ",";

                    // для легенды
                    //layerLEGENDrows = layerLEGENDrows.Union(mapSrcInDict.Value.tableLEGEND.AsEnumerable());
                }
                idmapsrcENUM += ")";

                // попробовать через коллекции для сравнения
                string queryStringLAYERname =
                    String.Concat("SELECT DISTINCT tblPO.idpnrmGROUPLAYER, tblPGL.pnrmGROUPLAYERcapt ",
                    "FROM objectsoke.dbo.tblPanoramaObject tblPO ",
                    "INNER JOIN objectsoke.dbo.tblPanoramaGroupLayer tblPGL ON tblPO.idpnrmGROUPLAYER = tblPGL.idpnrmGROUPLAYER ",
                    "WHERE tblPO.idmapsrc IN ", idmapsrcENUM,
                    " ORDER BY tblPGL.pnrmGROUPLAYERcapt ASC");
                MC_SQLDataProvider.SelectDataFromSQL(tableLAYERname, dbconnectionString, queryStringLAYERname);

                //layerLEGENDrows = layerLEGENDrows.Distinct(new DataRowComparer<DataRow>());
                string queryString1 = String.Concat(
                "SELECT * FROM objectsoke.dbo.tblPanoramaLegend",
                //" WHERE idpnrmGROUPLAYER IN ", idmapsrcENUM,                    
                " ORDER BY idpnrmGROUPLAYER ASC");
                MC_SQLDataProvider.SelectDataFromSQL(tableLEGENDmain, dbconnectionString, queryString1);
            } // if (mapSrcDictionary.Count > 0)

            // добавляем служебный слой XML
            //checkedListBoxControl1.Items.Add(43, "ЗАГРУЖЕНО", layerVisibleDictionary[43], true); // ИСПРАВИТЬ
            //-----------------------------

            int legendRowFirstColumnWidth = 30;
            int legendRowHeight = 20;

            // подсчитываем количество видимых слоев
            int countVisibleLayers = 0;
            foreach (var layerTemp in layerVisibleDictionary)
            {
                if (layerTemp.Value == CheckState.Checked) countVisibleLayers++;
            }

            //Image legendImg = new Bitmap(legendRowFirstColumnWidth * 6, legendRowHeight * countVisibleLayers); // динамически считать длину в зависимости от длины текста! ВКЛЮЧИТЬ, УБИРАЛ ВРЕМЕННО ДЛЯ БАЛ
            Image legendImg = new Bitmap(legendRowFirstColumnWidth * 6, legendRowHeight * 10); // динамически считать длину в зависимости от длины текста! УБРАТЬ см п.1 выше
            Graphics gLegend = Graphics.FromImage(legendImg);
            Rectangle gLegendRect = new Rectangle(0, 0, legendImg.Width - 1, legendImg.Height - 1);
            gLegend.FillRectangle(new SolidBrush(Color.White), gLegendRect);
            gLegend.DrawRectangle(new Pen(Color.Black), gLegendRect);

            int y = 0;
        
            foreach (DataRow rowlayer in tableLAYERname.Rows)            
            {                
                int layerid = (int)rowlayer["idpnrmGROUPLAYER"];

                if (layerVisibleDictionary[layerid] == CheckState.Checked)
                { 

                    string layerid_str = layerid.ToString();
                    string layercapt = rowlayer["pnrmGROUPLAYERcapt"].ToString();
                
                    //DataRow[] layerLEGENDrows = tableLEGENDmain.Select(String.Concat("idpnrmGROUPLAYER = ", layerid_str));
                    DataRow[] layerLEGENDrows = tableLEGENDmain.Select(String.Concat("idpnrmGROUPLAYER = ", layerid_str,
                        " AND idlgnddest = 1"));

                    foreach (DataRow layerLEGENDrow in layerLEGENDrows)
                    {
                        int idpnrmLOCALtype = (int)layerLEGENDrow["idpnrmLOCALtype"];

                        // переделать на заполнение из БД!!! + названия вкладок (стандарт, спутник, гибрид)
                        string pnrmLOCALtypecapt = string.Empty;
                        if (idpnrmLOCALtype == 2) pnrmLOCALtypecapt = "линейный";
                        else if (idpnrmLOCALtype == 3) pnrmLOCALtypecapt = "площадной";
                        else if (idpnrmLOCALtype == 4) pnrmLOCALtypecapt = "подпись";
                        else if (idpnrmLOCALtype == 5) pnrmLOCALtypecapt = "точечный";

                        int idlgnddest = (int)layerLEGENDrow["idlgnddest"];

                        /*Image legendRowImg = new Bitmap(50, 20);
                        Graphics gLegend = Graphics.FromImage(legendRowImg);*/

                        // формируем стиль контура
                        int line_width = (int)layerLEGENDrow["PENthickness"];
                        Pen layerPen = new Pen(Color.FromArgb((int)layerLEGENDrow["PENcolor"]), line_width);
                        int pen_style = (int)layerLEGENDrow["PENstyle"];
                        if (pen_style == 0) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;// сделать через массив!!!
                        else if (pen_style == 1) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDot;
                        else if (pen_style == 2) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDotDot;
                        else if (pen_style == 3) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                        else if (pen_style == 4) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;

                        // формируем стиль заливки
                        Brush layerBrush = new SolidBrush(Color.FromArgb((int)layerLEGENDrow["FILLSOLIDalpha"], Color.FromArgb((int)layerLEGENDrow["FILLSOLIDcolor"])));

                        // стиль точечного объекта
                        int point_type = (int)layerLEGENDrow["POINTtype"];

                        // отрисовка легенды в зависимости от типа объекта
                        int dx = 5; // смещение внутри строки (margin)
                        switch (idpnrmLOCALtype)
                        {
                            case 2: // линейный 
                                gLegend.DrawLine(layerPen, new Point(dx, y + (legendRowHeight - line_width) / 2),
                                                            new Point(legendRowFirstColumnWidth - dx, y + (legendRowHeight - line_width) / 2));
                                break;

                            case 3: // площадной
                                Rectangle pointrect = new Rectangle(dx, y + dx, legendRowFirstColumnWidth - dx - dx, y + legendRowHeight - dx - dx);
                                gLegend.FillRectangle(layerBrush, pointrect);
                                gLegend.DrawRectangle(layerPen, pointrect);
                                break;

                            case 4: // подпись
                                break;

                            case 5: // точечный
                                int dx2 = 8;
                                //dx = 2;
                                pointrect = new Rectangle((legendRowFirstColumnWidth - dx2) / 2, y + (legendRowHeight - dx2) / 2, dx2, y + dx2);
                                switch (point_type)
                                {
                                    case 0: // квадрат
                                        gLegend.FillRectangle(layerBrush, pointrect);
                                        gLegend.DrawRectangle(layerPen, pointrect);
                                        break;
                                    case 1: // круг
                                        gLegend.FillEllipse(layerBrush, pointrect);
                                        gLegend.DrawEllipse(layerPen, pointrect);
                                        break;
                                }
                                break;
                        }

                        gLegend.DrawString(layercapt, new Font("Arial", 8), new SolidBrush(Color.Black), legendRowFirstColumnWidth + dx, y + dx);

                        y += legendRowHeight;

                        layerPen.Dispose();
                        layerBrush.Dispose();
                    } // foreach (DataRow layerLEGENDrow in layerLEGENDrows)

                    //------------------------------------------------                   

                    /*if (idlgnddest == 1)
                        formShowLegend.dataGridView1.Rows.Add(layercapt, pnrmLOCALtypecapt, legendRowImg);
                    else if (idlgnddest == 2)
                        formShowLegend.dataGridView2.Rows.Add(layercapt, pnrmLOCALtypecapt, legendRowImg);
                    else if (idlgnddest == 3)
                        formShowLegend.dataGridView3.Rows.Add(layercapt, pnrmLOCALtypecapt, legendRowImg);*/
                } // if (layerVisibleDictionary[layerid] == CheckState.Checked)


            } // foreach (DataRow rowlayer in tableLAYERname.Rows)

            //g.DrawImage(legendImg, 0, 0);ВКЛЮЧИТЬ, УБИРАЛ ВРЕМЕННО ДЛЯ БАЛ

            gLegend.Dispose();

            // завершение создания легенды карты
            //--------------------------

            fullImage.Save("c:\\3\\fullImage.tiff", System.Drawing.Imaging.ImageFormat.Tiff);

            /*PointLatLng startPoint = pointLeftTop;            
            PointLatLng pCenter = new PointLatLng(startPoint.Lat - deltaLat / 4.0, startPoint.Lng + deltaLng / 4.0);                        
            gMapControl1.Position = pCenter;            
            gMapControl1.ToImage().Save("e:\\3\\01.tiff", System.Drawing.Imaging.ImageFormat.Tiff);
            
            startPoint.Lat = pointLeftTop.Lat;
            startPoint.Lng = pointLeftTop.Lng + deltaLng / 2.0;
            pCenter = new PointLatLng(startPoint.Lat - deltaLat / 4.0, startPoint.Lng + deltaLng / 4.0);
            gMapControl1.Position = pCenter;
            gMapControl1.ToImage().Save("e:\\3\\02.tiff", System.Drawing.Imaging.ImageFormat.Tiff);

            startPoint.Lat = pointLeftTop.Lat - deltaLat / 2.0;
            startPoint.Lng = pointLeftTop.Lng;
            pCenter = new PointLatLng(startPoint.Lat - deltaLat / 4.0, startPoint.Lng + deltaLng / 4.0);
            gMapControl1.Position = pCenter;
            gMapControl1.ToImage().Save("e:\\3\\03.tiff", System.Drawing.Imaging.ImageFormat.Tiff);

            startPoint.Lat = pointLeftTop.Lat - deltaLat / 2.0;
            startPoint.Lng = pointLeftTop.Lng + deltaLng / 2.0;
            pCenter = new PointLatLng(startPoint.Lat - deltaLat / 4.0, startPoint.Lng + deltaLng / 4.0);
            gMapControl1.Position = pCenter;
            gMapControl1.ToImage().Save("e:\\3\\04.tiff", System.Drawing.Imaging.ImageFormat.Tiff);*/

            // восстанавливаем предыдущие параметры
            gMapControl1.Position = prevPosition;
            gMapControl1.Zoom = prevZoom;
            gMapControl1.MapProvider.Copyright = prevCopyright;
            gMapControl1.ShowCenter = true;

            g.Dispose();
            fullImage.Dispose();

            //gMapControl1.Show();

            MessageBox.Show("Сохранение области выполнено");
            //------------------------------------------------
        }

        // тест
        private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
        {
            //gMapControl1.ShowExportDialog();            
            //gMapControl1.SetZoomToFitRect
            //gMapControl1.sele

            // отображение легенды карты ВЫНЕСТИ В ОТД КЛАСС/МЕТОД
            // объединяем таблицы слоев      
            DataTable table_LAYER_LOCALS_name = new DataTable();
            
            //IEnumerable<DataRow> layerLEGENDrows = new DataTable().AsEnumerable(); // пустая начальная таблица

            if (mapSrcDictionary.Count > 0)
            {
                string idmapsrcENUM = "(";
                int i = 0;
                foreach (var mapSrcInDict in mapSrcDictionary)
                {
                    //idmapsrcENUM = String.Concat(idmapsrcENUM, mapSrcInDict.Value.IdMapSrc.ToString());
                    idmapsrcENUM += mapSrcInDict.Value.IdMapSrc.ToString();
                    i++;
                    if (i < mapSrcDictionary.Count) idmapsrcENUM += ",";

                    // для легенды
                    //layerLEGENDrows = layerLEGENDrows.Union(mapSrcInDict.Value.tableLEGEND.AsEnumerable());
                }
                idmapsrcENUM += ")";

                // попробовать через коллекции для сравнения
                string queryStringLAYERname =
                    String.Concat("SELECT DISTINCT tblPO.idpnrmGROUPLAYER, tblPGL.pnrmGROUPLAYERcapt, idpnrmLOCALtype ",
                    "FROM objectsoke.dbo.tblPanoramaObject tblPO ",
                    "INNER JOIN objectsoke.dbo.tblPanoramaGroupLayer tblPGL ON tblPO.idpnrmGROUPLAYER = tblPGL.idpnrmGROUPLAYER ",
                    "WHERE tblPO.idmapsrc IN ", idmapsrcENUM,
                    " ORDER BY tblPGL.pnrmGROUPLAYERcapt ASC");
                MC_SQLDataProvider.SelectDataFromSQL(table_LAYER_LOCALS_name, dbconnectionString, queryStringLAYERname);
                                
            } // if (mapSrcDictionary.Count > 0)

            // добавляем служебный слой XML
            //checkedListBoxControl1.Items.Add(43, "ЗАГРУЖЕНО", layerVisibleDictionary[43], true); // ИСПРАВИТЬ
            //-----------------------------

            Image legendImage = MapLegend.GetMapLegendImage(mapSrcDictionary, table_LAYER_LOCALS_name, layerVisibleDictionary, 1300, 1300);

            legendImage.Save("c:\\3\\legImage.tiff", System.Drawing.Imaging.ImageFormat.Tiff);

            legendImage.Dispose();
        }

    }
}