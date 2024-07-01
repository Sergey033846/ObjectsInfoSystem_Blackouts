using rtp3esh_bd;

using DevExpress.Spreadsheet;
using DevExpress.XtraBars;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Columns;
using DevExpress.XtraTreeList.Nodes;
using DevExpress.XtraTreeList.Nodes.Operations;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using ObjectsInfoSystem.MyClasses;

namespace ObjectsInfoSystem
{
    public partial class FormBlackouts : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        static bool mailBlackoutsSent = false;
        private static void MailSendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            // Get the unique identifier for this asynchronous operation.
            //String token = (string)e.UserState;
            MailMessage msg = (MailMessage)e.UserState;

            /*if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message sent.");
            }*/
            mailBlackoutsSent = true;
            msg.Dispose();
            MessageBox.Show(text: "Файл отключений успешно отправлен!", caption: "Информация");
        }

        public class TreeListKey
        {
            public int ID_DB;
            public int GUID;
            public TreeListKey(int _ID_DB, int _GUID)
            {
                ID_DB = _ID_DB;
                GUID = _GUID;            
            }
        }
        public class TreeListTag
        {
            public int ID_DB;
            public int GUID;
            public string objectType;
            public DateTime blackOutDateFrom;
            public DateTime blackOutDateTo;
            public DateTime blackOutTimeFrom;
            public DateTime blackOutTimeTo;
            public TreeListTag(int _ID_DB, int _GUID, string _objectType)
            {
                ID_DB = _ID_DB;
                GUID = _GUID;
                objectType = _objectType;
            }
        }

        public class MyOperation : TreeListOperation
        {
            object nodeValue;
            TreeListColumn nodeColumn;
            public MyOperation(TreeListColumn NodeColumn, object NodeValue) : base()
            {
                this.nodeColumn = NodeColumn;
                this.nodeValue = NodeValue;
            }

            private void SetNodeCheckStatus2(TreeListNode node)
            {
                TreeListColumn nodeColumn1 = node.TreeList.Columns.ColumnByFieldName("Дата начала");
                TreeListColumn nodeColumn2 = node.TreeList.Columns.ColumnByFieldName("Время начала");
                TreeListColumn nodeColumn3 = node.TreeList.Columns.ColumnByFieldName("Дата окончания");
                TreeListColumn nodeColumn4 = node.TreeList.Columns.ColumnByFieldName("Время окончания");
                string s1 = null;
                string s2 = null;
                string s3 = null;
                string s4 = null;
                if (node.GetValue(nodeColumn1) != null) s1 = node.GetValue(nodeColumn1).ToString();
                if (node.GetValue(nodeColumn2) != null) s2 = node.GetValue(nodeColumn2).ToString();
                if (node.GetValue(nodeColumn3) != null) s3 = node.GetValue(nodeColumn3).ToString();
                if (node.GetValue(nodeColumn4) != null) s4 = node.GetValue(nodeColumn4).ToString();

                node.Checked = !String.IsNullOrEmpty(s1) && !String.IsNullOrEmpty(s2) && !String.IsNullOrEmpty(s3) && !String.IsNullOrEmpty(s4);
            }

            public override void Execute(TreeListNode node)
            {
                node.SetValue(nodeColumn, nodeValue);
                SetNodeCheckStatus(node);                
                foreach (TreeListNode childNode in node.Nodes)
                {
                    childNode.SetValue(nodeColumn, nodeValue);
                    SetNodeCheckStatus(childNode);
                }
            }
        }

        // пользовательские переменные -------------
        int permissionAdmin; // признак прав доступа администратора
                
        public string dbconnectionString; // строка подключения к бд
        
        public TreeNode nodemapsrcloaded; // ссылка на текущий загруженный узел дерева Филиал/Карта
        
        private IDictionary<int, MC_MapSource> mapSrcDictionary;
        private IDictionary<int, CheckState> layerVisibleDictionary;
                
        //---------------------------
        
        public FormBlackouts(string constr, int permissionAdmin_parent)
        {
            InitializeComponent();

            mapSrcDictionary = new Dictionary<int, MC_MapSource>();
            layerVisibleDictionary = new Dictionary<int, CheckState>();

            dbconnectionString = constr;

            permissionAdmin = permissionAdmin_parent;

            // настройка интерфейса в зависимости от прав доступа
            
            //----------------------------------------            
        }

        private void FormBlackouts_Load(object sender, EventArgs e)
        {
            //splashScreenManager1.ShowWaitForm();

            RTP3ESHEntities rtp3eshEntities = new RTP3ESHEntities();
                        
            /*string queryString = 
                String.Concat(
                    "select tblMapSource.idmapsrc, tblMapSource.captionmapsrc, tblMapSource.idcontractor, tblMapSource.idfilial, ",
                    "tblMapSource.commentmapsrc, tblMapSource.datecreated, tblMapSource.dateloaded, tblFilial.captionfilial, tblContractor.captioncontr ",
                    "from((dbo.tblMapSource tblMapSource ",
                    "inner join dbo.tblFilial tblFilial on(tblFilial.idfilial = tblMapSource.idfilial)) ",
                    "inner join dbo.tblContractor tblContractor on(tblContractor.idcontractor = tblMapSource.idcontractor))"
                    );
            DataTable table = new DataTable();            
            MC_SQLDataProvider.SelectDataFromSQL(table, dbconnectionString, queryString);*/

            /*string queryStringLAYERname =
                //String.Concat("SELECT idpnrmGROUPLAYER,pnrmLAYERcaptsrc,pnrmGROUPLAYERcapt ",
                String.Concat("SELECT idpnrmGROUPLAYER ",
                "FROM objectsoke.dbo.tblPanoramaGroupLayer ",
                " ORDER BY idpnrmGROUPLAYER ASC");
            DataTable tableGROUPLAYERname = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableGROUPLAYERname, dbconnectionString, queryStringLAYERname);*/

            // заполняем дерево объектов РТП-3
            /*DataSetPnrmMapSrc DataSetLoad = new DataSetPnrmMapSrc();            
            DataSetPnrmMapSrcTableAdapters.tblFilialTableAdapter tblFilialTableAdapter = new DataSetPnrmMapSrcTableAdapters.tblFilialTableAdapter();
            tblFilialTableAdapter.Fill(DataSetLoad.tblFilial);*/

            // сортировка?
            IQueryable<RTP3_DB> rtp3_DB_List = rtp3eshEntities.RTP3_DB.Select(p => p);
            IQueryable<Filial> rtp3_Filial_List = rtp3eshEntities.Filial.Select(p => p);
            IQueryable<RTP3_Centers> rtp3_Centers_List = rtp3eshEntities.RTP3_Centers.Select(p => p);
            IQueryable<RTP3_Sections> rtp3_Sections_List = rtp3eshEntities.RTP3_Sections.Select(p => p);
            IQueryable<RTP3_Fiders> rtp3_Fiders_List = rtp3eshEntities.RTP3_Fiders.Select(p => p);
            //IQueryable<RTP3_Nodes> rtp3_Nodes_List = rtp3eshEntities.RTP3_Nodes.Select(p => p);
            IQueryable<RTP3_Nodes> rtp3_Nodes_List = rtp3eshEntities.RTP3_Nodes.Where(p => (!String.IsNullOrEmpty(p.Name)) && (p.Transf > 0)).Select(p => p);
            IQueryable<RTP3_LVFiders> rtp3_LVFiders_List = rtp3eshEntities.RTP3_LVFiders.Select(p => p);
            //RTP3_Centers rtp3_Centers = rtp3eshEntities.RTP3_Centers.Where(p => p.id_zayavka == notifyTemp.link_id).Select(p => p).First();

            /*// старый компонент
            treeViewFilials.Nodes.Clear();            
            foreach (RTP3_DB rtp3_DB in rtp3_DB_List)
            {
                string rtp3_DB_caption = rtp3_DB.caption_bd;
                int rtp3_DB_id = rtp3_DB.ID_DB;

                IQueryable<RTP3_Centers> rtp3_DB_Centers_List = rtp3_Centers_List.Where(p => p.ID_DB == rtp3_DB_id).Select(p => p);

                if (rtp3_DB_Centers_List.Count() > 0)
                {
                    TreeNode[] nodes_Centers = new TreeNode[rtp3_DB_Centers_List.Count()];

                    int j = 0;
                    foreach (RTP3_Centers rtp3_DB_Center in rtp3_DB_Centers_List)
                    {
                        string captionnode_Center = rtp3_DB_Center.Name;
                        nodes_Centers[j] = new TreeNode(captionnode_Center);
                        //nodesmapsrc[j].Tag = mapsrcrows[j];
                        j++;
                    }

                    TreeNode node_rtp3_DB = new TreeNode(rtp3_DB_caption, nodes_Centers);
                    
                    treeViewFilials.Nodes.Add(node_rtp3_DB);
                }
            } 
            treeViewFilials.Sort(); // проверить*/

            // создаем соединение и загружаем данные
            string queryString =
            String.Concat(
                "SELECT ",
                "RTP3_Centers.ID_DB, RTP3_Centers.GUID AS Centers_GUID, RTP3_Centers.OwnerGUID AS Centers_OwnerGUID, RTP3_Centers.Name AS Centers_Name,",
                "RTP3_Sections.GUID AS Sections_GUID, RTP3_Sections.OwnerGUID AS Sections_OwnerGUID, RTP3_Sections.Name AS Sections_Name, RTP3_Sections.Unom,",
                "RTP3_Fiders.GUID AS Fiders_GUID, RTP3_Fiders.OwnerGUID AS Fiders_OwnerGUID, RTP3_Fiders.Name AS Fiders_Name,",
                "RTP3_Nodes.GUID AS Nodes_GUID, RTP3_Nodes.OwnerGUID AS Nodes_OwnerGUID, RTP3_Nodes.Name AS Nodes_Name, RTP3_Nodes.Transf AS Nodes_Transf, RTP3_Nodes.onBalance AS Nodes_onBalance,",
                "RTP3_LVFiders.GUID AS LVFiders_GUID, RTP3_LVFiders.OwnerGUID AS LVFiders_OwnerGUID, RTP3_LVFiders.Name AS LVFiders_Name",
                " FROM RTP3_Centers",
                " LEFT JOIN RTP3_Sections ON RTP3_Sections.ID_DB = RTP3_Centers.ID_DB AND RTP3_Sections.OwnerGUID = RTP3_Centers.GUID",
                " LEFT JOIN RTP3_Fiders ON RTP3_Fiders.ID_DB = RTP3_Sections.ID_DB AND RTP3_Fiders.OwnerGUID = RTP3_Sections.GUID",
                " LEFT JOIN RTP3_Nodes ON RTP3_Nodes.ID_DB = RTP3_Fiders.ID_DB AND RTP3_Nodes.OwnerGUID = RTP3_Fiders.GUID",
                " LEFT JOIN RTP3_LVFiders ON RTP3_LVFiders.ID_DB = RTP3_Nodes.ID_DB AND RTP3_LVFiders.OwnerGUID = RTP3_Nodes.GUID",
                " WHERE RTP3_Nodes.Name IS NOT NULL AND RTP3_Nodes.Transf > 0", /* 0 = опора, -1 = начало фидера*/
                " ORDER BY ID_DB, Centers_GUID, Sections_GUID, Fiders_GUID, Nodes_GUID" /*, LVFiders_GUID*/
                );
            DataTable dt = new DataTable();
            //MC_SQLDataProvider.SelectDataFromSQL(dt, dbconnectionString, queryString);

            //-----------------------
            // Создаем колонки
            treeList1.BeginUpdate();
            /*TreeListColumn col1 = treeList1.Columns.Add();
            col1.Caption = "Customer";
            col1.VisibleIndex = 0;*/
            /*TreeListColumn col2 = treeList1.Columns.Add();
            col2.Caption = "Location";
            col2.VisibleIndex = 1;
            TreeListColumn col3 = treeList1.Columns.Add();
            col3.Caption = "Phone";
            col3.VisibleIndex = 2;*/
            treeList1.EndUpdate();
            
            // новый компонент
            treeList1.BeginUnboundLoad();
            treeList1.Nodes.Clear();

            /*foreach (RTP3_DB rtp3_DB in rtp3_DB_List)
            {
                TreeListTag nodeDB_tag = new TreeListTag(rtp3_DB.ID_DB, -1, "DB");
                TreeListNode nodeDB = treeList1.AppendNode(new object[] { rtp3_DB.caption_bd }, null, nodeDB_tag);
            }

            foreach (DataRow dr in dt.Rows)
            {
                string nodeType = "DB";
                int idDB = Convert.ToInt32(dr["ID_DB"]);
                TreeListNode findNodeDB = treeList1.FindNode((node) => {
                    return (node.Tag as TreeListTag).ID_DB == idDB && (node.Tag as TreeListTag).objectType == nodeType;
                });

                nodeType = "Center";
                int guidCenter = Convert.ToInt32(dr["Centers_GUID"]);
                TreeListNode findNodeCenter = treeList1.FindNode((node) => {
                    return (node.Tag as TreeListTag).ID_DB == idDB && (node.Tag as TreeListTag).objectType == nodeType && (node.Tag as TreeListTag).GUID == guidCenter;
                });
                if (findNodeCenter == null)
                {
                    TreeListTag nodeTag = new TreeListTag(idDB, guidCenter, nodeType);
                    findNodeCenter = treeList1.AppendNode(new object[] { dr["Centers_Name"].ToString() }, findNodeDB, nodeTag);
                }

                if (!String.IsNullOrWhiteSpace(dr["Sections_GUID"].ToString()))
                {
                    nodeType = "Section";
                    int guidSection = Convert.ToInt32(dr["Sections_GUID"]);
                    TreeListNode findNodeSection = treeList1.FindNode((node) => {
                        return (node.Tag as TreeListTag).ID_DB == idDB && (node.Tag as TreeListTag).objectType == nodeType && (node.Tag as TreeListTag).GUID == guidSection;
                    });
                    if (findNodeSection == null)
                    {
                        TreeListTag nodeTag = new TreeListTag(idDB, guidSection, nodeType);
                        findNodeSection = treeList1.AppendNode(new object[] { dr["Sections_Name"].ToString() }, findNodeCenter, nodeTag);
                    }

                    //--------------------------------------------

                    if (!String.IsNullOrWhiteSpace(dr["Fiders_GUID"].ToString()))
                    {
                        nodeType = "Fider";
                        int guidFider = Convert.ToInt32(dr["Fiders_GUID"]);
                        TreeListNode findNodeFider = treeList1.FindNode((node) => {
                            return (node.Tag as TreeListTag).ID_DB == idDB && (node.Tag as TreeListTag).objectType == nodeType && (node.Tag as TreeListTag).GUID == guidFider;
                        });
                        if (findNodeFider == null)
                        {
                            TreeListTag nodeTag = new TreeListTag(idDB, guidFider, nodeType);
                            findNodeFider = treeList1.AppendNode(new object[] { dr["Fiders_Name"].ToString() }, findNodeSection, nodeTag);
                        }

                        //--------------------------------------------

                        if (!String.IsNullOrWhiteSpace(dr["Nodes_GUID"].ToString()))
                        {
                            nodeType = "Node";
                            int guidNode = Convert.ToInt32(dr["Nodes_GUID"]);
                            TreeListNode findNodeNode = treeList1.FindNode((node) => {
                                return (node.Tag as TreeListTag).ID_DB == idDB && (node.Tag as TreeListTag).objectType == nodeType && (node.Tag as TreeListTag).GUID == guidNode;
                            });
                            if (findNodeNode == null)
                            {
                                TreeListTag nodeTag = new TreeListTag(idDB, guidNode, nodeType);
                                findNodeNode = treeList1.AppendNode(new object[] { dr["Nodes_Name"].ToString() }, findNodeFider, nodeTag);
                            }

                            //--------------------------------------------

                            if (!String.IsNullOrWhiteSpace(dr["LVFiders_GUID"].ToString()))
                            {
                                nodeType = "LVFider";
                                int guidLVFider = Convert.ToInt32(dr["LVFiders_GUID"]);                                
                                TreeListTag nodeTag = new TreeListTag(idDB, guidLVFider, nodeType);
                                treeList1.AppendNode(new object[] { dr["LVFiders_Name"].ToString() }, findNodeNode, nodeTag);                                
                            }
                        }
                    }
                }                  
            } // foreach (DataRow dr in dt.Rows)*/

            /*// для теста быстродействия
            foreach (DataRow dr in dt.Rows)
            {
                string nodeType = "DB";
                int idDB = Convert.ToInt32(dr["ID_DB"]);
                TreeListNode findNodeDB = treeList1.FindNode((node) => {
                    return (node.Tag as TreeListTag).ID_DB == idDB && (node.Tag as TreeListTag).objectType == nodeType;
                });

                nodeType = "Center";
                int guidCenter = Convert.ToInt32(dr["Centers_GUID"]);                
                TreeListTag nodeTag = new TreeListTag(idDB, guidCenter, nodeType);
                treeList1.AppendNode(new object[] { dr["Centers_Name"].ToString() }, findNodeDB, nodeTag);

                if (!String.IsNullOrWhiteSpace(dr["Sections_GUID"].ToString()))
                {
                    nodeType = "Section";
                    int guidSection = Convert.ToInt32(dr["Sections_GUID"]);
                    nodeTag = new TreeListTag(idDB, guidSection, nodeType);
                    treeList1.AppendNode(new object[] { dr["Sections_Name"].ToString() }, findNodeDB, nodeTag);

                    if (!String.IsNullOrWhiteSpace(dr["Fiders_GUID"].ToString()))
                    {
                        nodeType = "Fider";
                        int guidFider = Convert.ToInt32(dr["Fiders_GUID"]);
                        nodeTag = new TreeListTag(idDB, guidFider, nodeType);
                        treeList1.AppendNode(new object[] { dr["Fiders_Name"].ToString() }, findNodeDB, nodeTag);

                        if (!String.IsNullOrWhiteSpace(dr["Nodes_GUID"].ToString()))
                        {
                            nodeType = "Node";
                            int guidNode = Convert.ToInt32(dr["Nodes_GUID"]);
                            nodeTag = new TreeListTag(idDB, guidNode, nodeType);
                            treeList1.AppendNode(new object[] { dr["Nodes_Name"].ToString() }, findNodeDB, nodeTag);

                            if (!String.IsNullOrWhiteSpace(dr["LVFiders_GUID"].ToString()))
                            {
                                nodeType = "LVFider";
                                int guidLVFider = Convert.ToInt32(dr["LVFiders_GUID"]);
                                nodeTag = new TreeListTag(idDB, guidLVFider, nodeType);
                                treeList1.AppendNode(new object[] { dr["LVFiders_Name"].ToString() }, findNodeDB, nodeTag);
                            }
                        }
                    }
                }
            } // foreach (DataRow dr in dt.Rows)*/

            foreach (RTP3_DB rtp3_DB in rtp3_DB_List)
            {                
                TreeListNode rootNode = treeList1.AppendNode(new object[] { rtp3_DB.caption_bd }, null);

                IQueryable<RTP3_Centers> child_Centers_List = rtp3_Centers_List.Where(p => p.ID_DB == rtp3_DB.ID_DB).Select(p => p);
                foreach (RTP3_Centers child_Center in child_Centers_List)
                {
                    //TreeListNode childNode_Center = treeList1.AppendNode(new object[] { child_Center.Name }, rootNode);
                    //TreeListTag childNode_Center_tag = new TreeListTag(child_Center.ID_DB, child_Center.GUID, "Center");
                    //TreeListNode childNode_Center = treeList1.AppendNode(new object[] { child_Center.Name }, rootNode, childNode_Center_tag);

                    IQueryable<RTP3_Sections> child_Sections_List = rtp3_Sections_List.Where(p => p.ID_DB == child_Center.ID_DB && p.OwnerGUID == child_Center.GUID).Select(p => p);
                    if (child_Sections_List.Count() > 0)
                    {
                        TreeListTag childNode_Center_tag = new TreeListTag(child_Center.ID_DB, child_Center.GUID, "Center");
                        TreeListNode childNode_Center = treeList1.AppendNode(new object[] { child_Center.Name }, rootNode, childNode_Center_tag);
                        foreach (RTP3_Sections child_Section in child_Sections_List)
                        {
                            //TreeListNode childNode_Section = treeList1.AppendNode(new object[] { child_Section.Name }, childNode_Center);
                            //TreeListTag childNode_Section_tag = new TreeListTag(child_Section.ID_DB, child_Section.GUID, "Section");
                            //TreeListNode childNode_Section = treeList1.AppendNode(new object[] { child_Section.Name }, childNode_Center, childNode_Section_tag);

                            IQueryable<RTP3_Fiders> child_Fiders_List = rtp3_Fiders_List.Where(p => p.ID_DB == child_Section.ID_DB && p.OwnerGUID == child_Section.GUID).Select(p => p);
                            if (child_Fiders_List.Count() > 0)
                            {
                                TreeListTag childNode_Section_tag = new TreeListTag(child_Section.ID_DB, child_Section.GUID, "Section");
                                TreeListNode childNode_Section = treeList1.AppendNode(new object[] { child_Section.Name }, childNode_Center, childNode_Section_tag);
                                foreach (RTP3_Fiders child_Fider in child_Fiders_List)
                                {
                                    //TreeListNode childNodeFider = treeList1.AppendNode(new object[] { child_Fiders.Name }, childNode_Section);
                                    //TreeListTag childNode_Fider_tag = new TreeListTag(child_Fider.ID_DB, child_Fider.GUID, "Fider");
                                    //TreeListNode childNode_Fider = treeList1.AppendNode(new object[] { child_Fider.Name }, childNode_Center, childNode_Fider_tag);
                                    //TreeListNode childNode_Fider = treeList1.AppendNode(new object[] { child_Fider.Name }, childNode_Section, childNode_Fider_tag);

                                    IQueryable<RTP3_Nodes> child_Nodes_List =
                                        rtp3_Nodes_List.Where(p => p.ID_DB == child_Fider.ID_DB && p.OwnerGUID == child_Fider.GUID
                                        && !String.IsNullOrEmpty(p.Name) && !p.Name.Contains("оп.")).Select(p => p);
                                    if (child_Nodes_List.Count() > 0)
                                    {
                                        TreeListTag childNode_Fider_tag = new TreeListTag(child_Fider.ID_DB, child_Fider.GUID, "Fider");
                                        TreeListNode childNode_Fider = treeList1.AppendNode(new object[] { child_Fider.Name }, childNode_Section, childNode_Fider_tag);
                                        foreach (RTP3_Nodes child_Node in child_Nodes_List)
                                        {
                                            TreeListTag childNode_Node_tag = new TreeListTag(child_Node.ID_DB, child_Node.GUID, "Node");
                                            TreeListNode childNode_Node = treeList1.AppendNode(new object[] { child_Node.Info }, childNode_Fider, childNode_Node_tag);

                                            IQueryable<RTP3_LVFiders> child_LVFiders_List =
                                                rtp3_LVFiders_List.Where(p => p.ID_DB == child_Node.ID_DB && p.OwnerGUID == child_Node.GUID).Select(p => p);
                                            foreach (RTP3_LVFiders child_LVFider in child_LVFiders_List)
                                            {
                                                TreeListTag childNode_LVFider_tag = new TreeListTag(child_LVFider.ID_DB, child_LVFider.GUID, "LVFider");
                                                TreeListNode childNode_LVFider = treeList1.AppendNode(new object[] { child_LVFider.Name }, childNode_Node, childNode_LVFider_tag);
                                            }
                                        }

                                        if (!childNode_Fider.HasChildren) treeList1.DeleteNode(childNode_Fider);

                                    } // if (child_Nodes_List.Count() > 0)
                                }

                                if (!childNode_Section.HasChildren) treeList1.DeleteNode(childNode_Section);

                            } // if (child_Fiders_List.Count() > 0)                            
                        } // foreach (RTP3_Sections child_Section in child_Sections_List)

                        if (!childNode_Center.HasChildren) treeList1.DeleteNode(childNode_Center);

                    } // if (child_Sections_List.Count() > 0)                    
                } // foreach (RTP3_Centers child_Center in child_Centers_List)
            } // foreach (RTP3_DB rtp3_DB in rtp3_DB_List)
            
            treeList1.BestFitColumns();            
            treeList1.EndUnboundLoad();

            barButtonNodesCount.Caption = String.Format("Элементов сети: {0}", treeList1.AllNodesCount.ToString());

            /*treeViewFilials.Nodes.Clear();
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
            treeViewFilials.Sort();*/

            //splashScreenManager1.CloseWaitForm();
        } // private void FormBlackouts_Load(object sender, EventArgs e)

        static void SetNodeCheckStatus(TreeListNode node)
        {
            TreeListColumn nodeColumn1 = node.TreeList.Columns.ColumnByFieldName("Дата начала");
            TreeListColumn nodeColumn2 = node.TreeList.Columns.ColumnByFieldName("Время начала");
            TreeListColumn nodeColumn3 = node.TreeList.Columns.ColumnByFieldName("Дата окончания");
            TreeListColumn nodeColumn4 = node.TreeList.Columns.ColumnByFieldName("Время окончания");
            DateTime? d1 = null; string s1 = null;
            string s2 = null;
            DateTime? d3 = null; string s3 = null;
            string s4 = null;
            if (node.GetValue(nodeColumn1) != null)
            {
                d1 = Convert.ToDateTime(node.GetValue(nodeColumn1));
                s1 = node.GetValue(nodeColumn1).ToString();
            }
            if (node.GetValue(nodeColumn2) != null) s2 = node.GetValue(nodeColumn2).ToString();
            if (node.GetValue(nodeColumn3) != null)
            {
                d3 = Convert.ToDateTime(node.GetValue(nodeColumn3));
                s3 = node.GetValue(nodeColumn3).ToString();
            }
            if (node.GetValue(nodeColumn4) != null) s4 = node.GetValue(nodeColumn4).ToString();

            bool flagDatesChecking = d1 <= d3;

            node.Checked = !String.IsNullOrEmpty(s1) && !String.IsNullOrEmpty(s2) && !String.IsNullOrEmpty(s3) && !String.IsNullOrEmpty(s4) && flagDatesChecking;
        }

        private bool myIsAllNodesCellsFilled(TreeListNode node)
        {
            TreeListColumn nodeColumn1 = node.TreeList.Columns.ColumnByFieldName("Дата начала");
            TreeListColumn nodeColumn2 = node.TreeList.Columns.ColumnByFieldName("Время начала");
            TreeListColumn nodeColumn3 = node.TreeList.Columns.ColumnByFieldName("Дата окончания");
            TreeListColumn nodeColumn4 = node.TreeList.Columns.ColumnByFieldName("Время окончания");
            DateTime? d1 = null; string s1 = null;
            string s2 = null;
            DateTime? d3 = null; string s3 = null;
            string s4 = null;
            if (node.GetValue(nodeColumn1) != null)
            {
                d1 = Convert.ToDateTime(node.GetValue(nodeColumn1));
                s1 = node.GetValue(nodeColumn1).ToString();                
            }
            if (node.GetValue(nodeColumn2) != null) s2 = node.GetValue(nodeColumn2).ToString();
            if (node.GetValue(nodeColumn3) != null)
            {
                d3 = Convert.ToDateTime(node.GetValue(nodeColumn3));
                s3 = node.GetValue(nodeColumn3).ToString();
            }
            if (node.GetValue(nodeColumn4) != null) s4 = node.GetValue(nodeColumn4).ToString();

            bool flagDatesChecking = d1 <= d3;

            return !String.IsNullOrEmpty(s1) && !String.IsNullOrEmpty(s2) && !String.IsNullOrEmpty(s3) && !String.IsNullOrEmpty(s4) && flagDatesChecking;
        }

        // "спуск" значений ячеек столбцов информации об отключениях дочерним элементам
        private void treeList1_CellValueChanged(object sender, CellValueChangedEventArgs e)
        {
            treeList1.BeginUnboundLoad();

            /*TreeListNode cellValueChangedNode = e.Node;
            
            foreach (TreeListNode childNode in cellValueChangedNode.Nodes)
            {
                childNode.SetValue(e.Column, e.Value);                
            }*/
            //treeList1.NodesIterator.DoOperation(new MyOperation(e.Node));

            /*if (myIsAllNodesCellsFilled(e.Node))
            {*/
                SetNodeCheckStatus(e.Node);                
                treeList1.NodesIterator.DoLocalOperation(new MyOperation(e.Column, e.Value), e.Node.Nodes);
            //}

            treeList1.EndUnboundLoad();
        } // private void treeList1_CellValueChanged(object sender, CellValueChangedEventArgs e)

        // реестр потребителей (просмотр)
        private void barButtonBlackoutConsumers_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (treeList1.GetAllCheckedNodes().Count > 0)
            {
                splashScreenManager1.ShowWaitForm();
                splashScreenManager1.SetWaitFormDescription("получение данных...");

                string queryFilter = String.Empty;
                List<TreeListNode> treeListCheckedNodes = this.treeList1.GetAllCheckedNodes();
                //Dictionary<TreeListKey, TreeListTag> treeListCheckedNodesDictionary = new Dictionary<TreeListKey, TreeListTag>();
                Dictionary<string, TreeListTag> treeListCheckedNodesDictionary = new Dictionary<string, TreeListTag>();
                bool firstNodeinQuery = true;

                TreeListColumn nodeColumnBlackOutDateFrom = this.treeList1.Columns.ColumnByFieldName("Дата начала");
                TreeListColumn nodeColumnBlackOutTimeFrom = this.treeList1.Columns.ColumnByFieldName("Время начала");
                TreeListColumn nodeColumnBlackOutDateTo = this.treeList1.Columns.ColumnByFieldName("Дата окончания");
                TreeListColumn nodeColumnBlackOutTimeTo = this.treeList1.Columns.ColumnByFieldName("Время окончания");

                foreach (TreeListNode checkedNode in treeListCheckedNodes)
                {
                    // учитываем только листовые элементы
                    if (!checkedNode.HasChildren)
                    {
                        TreeListTag nodeTag = checkedNode.Tag as TreeListTag;
                        if (!firstNodeinQuery)
                        {
                            queryFilter = String.Concat(queryFilter, " OR ");
                        }
                        /*if (nodeTag.objectType == "Center") queryFilter = String.Concat(queryFilter, String.Format("(RTP3_Centers.ID_DB = {0} AND RTP3_Centers.GUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "Fider") queryFilter = String.Concat(queryFilter, String.Format("(RTP3_Centers.ID_DB = {0} AND RTP3_Fiders.GUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "Node") queryFilter = String.Concat(queryFilter, String.Format("(RTP3_Centers.ID_DB = {0} AND RTP3_Nodes.GUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "LVFider") queryFilter = String.Concat(queryFilter, String.Format("(RTP3_Centers.ID_DB = {0} AND LVFiders_GUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));*/

                        if (nodeTag.objectType == "Center") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND CenterGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "Section") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND SectionGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "Fider") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND FiderGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "Node") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND NodeGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                        if (nodeTag.objectType == "LVFider") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND LVFiderGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));

                        firstNodeinQuery = false;

                        nodeTag.blackOutDateFrom = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutDateFrom));
                        nodeTag.blackOutDateTo = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutDateTo));
                        nodeTag.blackOutTimeFrom = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutTimeFrom));
                        nodeTag.blackOutTimeTo = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutTimeTo));
                        //treeListCheckedNodesDictionary.Add(new TreeListKey(nodeTag.ID_DB, nodeTag.GUID), nodeTag);
                        treeListCheckedNodesDictionary.Add(String.Concat(nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()), nodeTag);
                    }
                }

                // создаем соединение и загружаем данные
                /*string queryString = 
                String.Concat(
                    "SELECT",
                    " ConsumersInfo.phoneInfo, ConsumersInfo.addressInfo, ConsumersInfo.codeIESBK,",
                    " RTP3_Centers.ID_DB, RTP3_Centers.GUID AS Centers_GUID, RTP3_Centers.OwnerGUID AS Centers_OwnerGUID, RTP3_Centers.Name AS Centers_Name,",
                    " RTP3_Sections.ID_DB AS Sections_ID_DB, RTP3_Sections.GUID AS Sections_GUID, RTP3_Sections.OwnerGUID AS Sections_OwnerGUID, RTP3_Sections.Name AS Sections_Name, RTP3_Sections.Unom,",
                    " RTP3_Fiders.GUID AS Fiders_GUID, RTP3_Fiders.OwnerGUID AS Fiders_OwnerGUID, RTP3_Fiders.Name AS Fiders_Name,",
                    " RTP3_Nodes.GUID AS Nodes_GUID, RTP3_Nodes.OwnerGUID AS Nodes_OwnerGUID, RTP3_Nodes.Name AS Nodes_Name, RTP3_Nodes.Transf AS Nodes_Transf, RTP3_Nodes.onBalance AS Nodes_onBalance,",
                    " mt_LVConsumersA_LVNodes_Fiders.LVFiders_GUID AS LVFiders_GUID, mt_LVConsumersA_LVNodes_Fiders.LVFiders_OwnerGUID AS LVFiders_OwnerGUID, mt_LVConsumersA_LVNodes_Fiders.LVFiders_Name AS LVFiders_Name",

                    " FROM RTP3_Centers",
                    " LEFT JOIN RTP3_Sections ON RTP3_Sections.ID_DB = RTP3_Centers.ID_DB AND RTP3_Sections.OwnerGUID = RTP3_Centers.GUID",
                    " LEFT JOIN RTP3_Fiders ON RTP3_Fiders.ID_DB = RTP3_Sections.ID_DB AND RTP3_Fiders.OwnerGUID = RTP3_Sections.GUID",
                    " LEFT JOIN RTP3_Nodes ON RTP3_Nodes.ID_DB = RTP3_Fiders.ID_DB AND RTP3_Nodes.OwnerGUID = RTP3_Fiders.GUID",
                    " RIGHT JOIN",
                    " (",
                    " SELECT",
                    " mt_LVConsumersA_LVNodes.ID_DB AS LVConsumersA_LVNodes_ID_DB, mt_LVConsumersA_LVNodes.GUID AS LVConsumersA_LVNodes_GUID, mt_LVConsumersA_LVNodes.OwnerGUID AS LVConsumersA_LVNodes_OwnerGUID, mt_LVConsumersA_LVNodes.Name AS LVConsumersA_LVNodes_Name,",
                    " LVFiders_GUID, LVFiders_OwnerGUID, (CASE WHEN LVFiders_Name IS NULL THEN 'подключен к ТП' ELSE LVFiders_Name END) AS LVFiders_Name",
                    " FROM(SELECT ID_DB, GUID, OwnerGUID, Name FROM RTP3_LVConsumersA UNION SELECT ID_DB, GUID, OwnerGUID, Name FROM RTP3_LVNodes) AS mt_LVConsumersA_LVNodes",
                    " LEFT JOIN",
                    " (SELECT ID_DB AS LVFiders_ID_DB, GUID AS LVFiders_GUID, OwnerGUID AS LVFiders_OwnerGUID, Name AS LVFiders_Name FROM RTP3_LVFiders) AS mt_LVFiders",
                    " ON mt_LVConsumersA_LVNodes.ID_DB = mt_LVFiders.LVFiders_ID_DB AND mt_LVConsumersA_LVNodes.OwnerGUID = mt_LVFiders.LVFiders_GUID",
                    " ) AS mt_LVConsumersA_LVNodes_Fiders",
                    " ON(mt_LVConsumersA_LVNodes_Fiders.LVConsumersA_LVNodes_OwnerGUID = RTP3_Nodes.GUID) OR(mt_LVConsumersA_LVNodes_Fiders.LVFiders_OwnerGUID = RTP3_Nodes.GUID)",

                    " RIGHT JOIN ConsumersInfo ON mt_LVConsumersA_LVNodes_Fiders.LVConsumersA_LVNodes_ID_DB = ConsumersInfo.ID_DB AND mt_LVConsumersA_LVNodes_Fiders.LVConsumersA_LVNodes_GUID = ConsumersInfo.GUIDRTP3",
                    " WHERE (ConsumersInfo.phoneInfo IS NOT NULL)", String.Format(" AND {0}", queryFilter)                
                    );*/
                string queryString =
                String.Concat(
                    "SELECT ",
                    "ID_DB, CenterGUID, SectionGUID, FiderGUID, Transforms2_Ident, NodeGUID, LVFiderGUID,",
                    "codeIESBK, codeOKE, addressInfo, phoneInfo, emailInfo, GUIDRTP3, ContractRTP3, houseFIASid ",
                    "FROM ConsumersInfo ",
                    //"WHERE (phoneInfo IS NOT NULL)", String.Format(" AND {0}", queryFilter)
                    "WHERE ", String.Format("{0}", queryFilter)
                    );
                DataTable dt = new DataTable();
                MC_SQLDataProvider.SelectDataFromSQL(dt, dbconnectionString, queryString);

                if (dt.Rows.Count > 0)
                {
                    // формируем отчет
                    splashScreenManager1.SetWaitFormDescription("формирование списка...");

                    FormReportGrid formReport = null;
                    formReport = new FormReportGrid();
                    //formReport.MdiParent = this;
                    formReport.Text = String.Concat("Реестр потребителей");
                    IWorkbook workbook = formReport.spreadsheetControl1.Document;
                    Worksheet worksheet = workbook.Worksheets[0];

                    workbook.History.IsEnabled = false;
                    formReport.spreadsheetControl1.BeginUpdate();

                    int row_start_data = 1;
                    int row_columns_names = 0;

                    // выводим названия столбцов
                    string[] columns_names = {
                "№ п/п", "Телефон", "Дата начала", "Дата окончания", "Время начала", "Время окончания", "Адрес", "Код ИЭСБК", "ФИАС дом GUID",
                "Код ОКЭ"
                };

                    for (int col = 0; col < 10; col++)
                    {
                        worksheet[row_columns_names, col].SetValue(columns_names[col]);
                    }

                    // форматируем заголовок таблицы                                
                    worksheet.Rows[row_columns_names].Alignment.WrapText = true;
                    worksheet.Rows[row_columns_names].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                    worksheet.Rows[row_columns_names].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                    worksheet.Rows[row_columns_names].Font.Name = "Arial";
                    worksheet.Rows[row_columns_names].Font.Size = 8;

                    //worksheet.Range["A1:AE3"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                    // выводим данные
                    int row_index = 0;
                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        string phoneInfo = dt.Rows[row]["phoneInfo"].ToString(); // Телефон

                        /*if (!String.IsNullOrEmpty(phoneInfo))
                        {*/
                            worksheet.Rows[row_index + row_start_data].Font.Name = "Arial";
                            worksheet.Rows[row_index + row_start_data].Font.Size = 8;

                            worksheet[row_index + row_start_data, 0].SetValue((row_index + 1).ToString("D5")); // № п/п

                            worksheet[row_index + row_start_data, 1].SetValue(phoneInfo); // Телефон

                            int id_DB = Convert.ToInt32(dt.Rows[row]["ID_DB"].ToString());
                            int centers_GUID = Convert.ToInt32(dt.Rows[row]["CenterGUID"].ToString());
                            int fiders_GUID = Convert.ToInt32(dt.Rows[row]["FiderGUID"].ToString());
                            int nodes_GUID = Convert.ToInt32(dt.Rows[row]["NodeGUID"].ToString());
                            int? lvFiders_GUID = null;
                            if (!String.IsNullOrEmpty(dt.Rows[row]["LVFiderGUID"].ToString())) lvFiders_GUID = Convert.ToInt32(dt.Rows[row]["LVFiderGUID"]);
                            /*TreeListKey treeListKeyCenters = new TreeListKey(id_DB, centers_GUID);
                            TreeListKey treeListKeyFiders = new TreeListKey(id_DB, fiders_GUID);
                            TreeListKey treeListKeyNodes = new TreeListKey(id_DB, nodes_GUID);
                            TreeListKey treeListKeyLVFiders = new TreeListKey(id_DB, lvFiders_GUID);*/
                            string treeListKeyCenters = String.Concat(id_DB.ToString(), centers_GUID.ToString());
                            string treeListKeyFiders = String.Concat(id_DB.ToString(), fiders_GUID.ToString());
                            string treeListKeyNodes = String.Concat(id_DB.ToString(), nodes_GUID.ToString());
                            string treeListKeyLVFiders = String.Concat(id_DB.ToString(), lvFiders_GUID.ToString());
                            TreeListTag treeListTag = null;
                            if (lvFiders_GUID != null && treeListCheckedNodesDictionary.ContainsKey(treeListKeyLVFiders))
                            {
                                treeListTag = treeListCheckedNodesDictionary[treeListKeyLVFiders];
                            }
                            else if (treeListCheckedNodesDictionary.ContainsKey(treeListKeyNodes)) treeListTag = treeListCheckedNodesDictionary[treeListKeyNodes];
                            else if (treeListCheckedNodesDictionary.ContainsKey(treeListKeyFiders)) treeListTag = treeListCheckedNodesDictionary[treeListKeyFiders];
                            else if (treeListCheckedNodesDictionary.ContainsKey(treeListKeyCenters)) treeListTag = treeListCheckedNodesDictionary[treeListKeyCenters];

                            if (treeListTag != null)
                            {
                                worksheet[row_index + row_start_data, 2].SetValue(treeListTag.blackOutDateFrom.ToString("yyyy-MM-dd")); // Дата начала отключения
                                worksheet[row_index + row_start_data, 3].SetValue(treeListTag.blackOutDateTo.ToString("yyyy-MM-dd")); // Дата окончания отключения
                                worksheet[row_index + row_start_data, 4].SetValue(treeListTag.blackOutTimeFrom.ToString("HH:mm")); // Время начала отключения
                                worksheet[row_index + row_start_data, 5].SetValue(treeListTag.blackOutTimeTo.ToString("HH:mm")); // Время окончания отключения
                            }

                            worksheet[row_index + row_start_data, 6].SetValue(dt.Rows[row]["addressInfo"].ToString()); // Адрес

                            if (!String.IsNullOrEmpty(dt.Rows[row]["codeIESBK"].ToString())) worksheet[row_index + row_start_data, 7].SetValue(dt.Rows[row]["codeIESBK"].ToString()); // Код ИЭСБК   
                            else worksheet[row_index + row_start_data, 7].SetValue(dt.Rows[row]["codeOKE"].ToString()); // Код ИЭСБК   

                            worksheet[row_index + row_start_data, 8].SetValue(dt.Rows[row]["houseFIASid"].ToString()); // ФИАС дом GUID

                            worksheet[row_index + row_start_data, 9].SetValue(dt.Rows[row]["codeOKE"].ToString()); // Код ОКЭ   
                            //worksheet[row_index + row_start_data, 10].SetValue(dt.Rows[row]["codeOKE"].ToString()); // Потребитель   

                            row_index++;
                        //}
                    } // for (int row = 0; row < dt.Rows.Count; row++)

                    // выравнивание столбцов
                    worksheet.Columns[0].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                    splashScreenManager1.SetWaitFormDescription("выравнивание ширины столбцов...");
                    worksheet.Columns.AutoFit(0, 10);

                    formReport.spreadsheetControl1.EndUpdate();

                    splashScreenManager1.CloseWaitForm();

                    formReport.Show();
                } //if (dt.Rows.Count > 0)
                else
                {
                    splashScreenManager1.CloseWaitForm();
                    MessageBox.Show(text: "Список потребителей пуст!", caption: "Ошибка");
                }
            } // if (treeList1.GetAllCheckedNodes().Count > 0)
            else
            {                
                MessageBox.Show(text: "Не выбраны элементы сети!", caption: "Ошибка");
            }

        } // private void barButtonBlackoutConsumers_ItemClick(object sender, ItemClickEventArgs e)

        // реестр потребителей (для ИЭСБК)
        private void barButtonBlackoutConsumers2CSV_ItemClick(object sender, ItemClickEventArgs e)
        {
            FormBlackoutsTypeSelect formBlackoutsTypeSelect = null;
            formBlackoutsTypeSelect = new FormBlackoutsTypeSelect();
            if (formBlackoutsTypeSelect.ShowDialog(this) == DialogResult.OK)
            {
                if (treeList1.GetAllCheckedNodes().Count > 0)
                {
                    splashScreenManager1.ShowWaitForm();

                    RTP3ESHEntities rtp3eshEntities = new RTP3ESHEntities();
                    ProgramOptions programOptions = rtp3eshEntities.ProgramOptions.Select(p => p).First();
                    int blackoutsFileLastNumber = programOptions.blackoutsFileLastNumber;
                    string pathBlackoutsFiles = String.Concat(Application.StartupPath, "\\Blackouts");

                    string filenameBlackoutFile =
                        formBlackoutsTypeSelect.radioButtonPlan.Checked ? String.Format("99{0:yyMMdd}{1:D2}.csv", DateTime.Now, blackoutsFileLastNumber + 1) 
                        : String.Format("АВАРИЯ_{0:yyMMdd}{1:D2}.csv", DateTime.Now, blackoutsFileLastNumber + 1);

                    string fullFilenameBlackoutFile = String.Concat(pathBlackoutsFiles, "\\", filenameBlackoutFile);

                    splashScreenManager1.SetWaitFormDescription("формирование списка потребителей...");

                    string queryFilter = String.Empty;
                    List<TreeListNode> treeListCheckedNodes = this.treeList1.GetAllCheckedNodes();
                    Dictionary<string, TreeListTag> treeListCheckedNodesDictionary = new Dictionary<string, TreeListTag>();
                    bool firstNodeinQuery = true;

                    TreeListColumn nodeColumnBlackOutDateFrom = this.treeList1.Columns.ColumnByFieldName("Дата начала");
                    TreeListColumn nodeColumnBlackOutTimeFrom = this.treeList1.Columns.ColumnByFieldName("Время начала");
                    TreeListColumn nodeColumnBlackOutDateTo = this.treeList1.Columns.ColumnByFieldName("Дата окончания");
                    TreeListColumn nodeColumnBlackOutTimeTo = this.treeList1.Columns.ColumnByFieldName("Время окончания");

                    foreach (TreeListNode checkedNode in treeListCheckedNodes)
                    {
                        // учитываем только листовые элементы
                        if (!checkedNode.HasChildren)
                        {
                            TreeListTag nodeTag = checkedNode.Tag as TreeListTag;
                            if (!firstNodeinQuery)
                            {
                                queryFilter = String.Concat(queryFilter, " OR ");
                            }

                            if (nodeTag.objectType == "Center") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND CenterGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                            if (nodeTag.objectType == "Section") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND SectionGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                            if (nodeTag.objectType == "Fider") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND FiderGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                            if (nodeTag.objectType == "Node") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND NodeGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));
                            if (nodeTag.objectType == "LVFider") queryFilter = String.Concat(queryFilter, String.Format("(ID_DB = {0} AND LVFiderGUID = {1})", nodeTag.ID_DB.ToString(), nodeTag.GUID.ToString()));

                            firstNodeinQuery = false;

                            nodeTag.blackOutDateFrom = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutDateFrom));
                            nodeTag.blackOutDateTo = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutDateTo));
                            nodeTag.blackOutTimeFrom = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutTimeFrom));
                            nodeTag.blackOutTimeTo = Convert.ToDateTime(checkedNode.GetValue(nodeColumnBlackOutTimeTo));
                            //treeListCheckedNodesDictionary.Add(new TreeListKey(nodeTag.ID_DB, nodeTag.GUID), nodeTag);
                            treeListCheckedNodesDictionary.Add(String.Concat(nodeTag.ID_DB.ToString(), "-", nodeTag.GUID.ToString()), nodeTag);
                        }
                    }

                    // создаем соединение и загружаем данные
                    string queryString =
                    String.Concat(
                        "SELECT ",
                        "ID_DB, CenterGUID, SectionGUID, FiderGUID, Transforms2_Ident, NodeGUID, LVFiderGUID,",
                        "codeIESBK, codeOKE, addressInfo, phoneInfo, emailInfo, GUIDRTP3, ContractRTP3, houseFIASid ",
                        "FROM ConsumersInfo ",
                        //"WHERE (phoneInfo IS NOT NULL)", String.Format(" AND {0}", queryFilter)
                        "WHERE ", String.Format("{0}", queryFilter)

                        /*" WHERE (RTP3_LVConsumersInfo.Contract IS NOT NULL)", String.Format(" AND {0}", queryFilter),
                        " ORDER BY LVConsumersInfo_Customer"*/
                        );
                    DataTable dt = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(dt, dbconnectionString, queryString);

                    if (dt.Rows.Count > 0)
                    {
                        // сохраняем файл
                        splashScreenManager1.SetWaitFormDescription("формирование файла...");

                        //using (StreamWriter sw = new StreamWriter(saveFileDialogCSV.FileName, false, System.Text.Encoding.Default))
                        using (StreamWriter sw = new StreamWriter(fullFilenameBlackoutFile, false, System.Text.Encoding.UTF8))
                        {
                            sw.WriteLine(String.Format("#Телефоны отключенных потребителей от {0:G}", DateTime.Now));

                            int row_index = 0;
                            for (int row = 0; row < dt.Rows.Count; row++)
                            {
                                string phoneInfo = dt.Rows[row]["phoneInfo"].ToString(); // Телефон
                                string codeIESBK = dt.Rows[row]["codeIESBK"].ToString();
                                string codeOKE = dt.Rows[row]["codeOKE"].ToString();

                                string stringBlackOut = "";
                                if (!String.IsNullOrEmpty(phoneInfo) || (String.IsNullOrEmpty(phoneInfo) && String.IsNullOrEmpty(codeIESBK) && !codeOKE.Contains("#N/A"))) //if (!String.IsNullOrEmpty(phoneInfo))
                                {
                                    stringBlackOut += (row_index + 1).ToString("D5") + ";";
                                    stringBlackOut += phoneInfo + ";";

                                    int id_DB = Convert.ToInt32(dt.Rows[row]["ID_DB"].ToString());
                                    int centers_GUID = Convert.ToInt32(dt.Rows[row]["CenterGUID"].ToString());
                                    int fiders_GUID = Convert.ToInt32(dt.Rows[row]["FiderGUID"].ToString());
                                    int nodes_GUID = Convert.ToInt32(dt.Rows[row]["NodeGUID"].ToString());
                                    int? lvFiders_GUID = null;
                                    if (!String.IsNullOrEmpty(dt.Rows[row]["LVFiderGUID"].ToString())) lvFiders_GUID = Convert.ToInt32(dt.Rows[row]["LVFiderGUID"]);

                                    string treeListKeyCenters = String.Concat(id_DB.ToString(), "-", centers_GUID.ToString());
                                    string treeListKeyFiders = String.Concat(id_DB.ToString(), "-", fiders_GUID.ToString());
                                    string treeListKeyNodes = String.Concat(id_DB.ToString(), "-", nodes_GUID.ToString());
                                    string treeListKeyLVFiders = String.Concat(id_DB.ToString(), "-", lvFiders_GUID.ToString());
                                    TreeListTag treeListTag = null;
                                    if (lvFiders_GUID != null && treeListCheckedNodesDictionary.ContainsKey(treeListKeyLVFiders))
                                    {
                                        treeListTag = treeListCheckedNodesDictionary[treeListKeyLVFiders];
                                    }
                                    else if (treeListCheckedNodesDictionary.ContainsKey(treeListKeyNodes)) treeListTag = treeListCheckedNodesDictionary[treeListKeyNodes];
                                    else if (treeListCheckedNodesDictionary.ContainsKey(treeListKeyFiders)) treeListTag = treeListCheckedNodesDictionary[treeListKeyFiders];
                                    else if (treeListCheckedNodesDictionary.ContainsKey(treeListKeyCenters)) treeListTag = treeListCheckedNodesDictionary[treeListKeyCenters];

                                    if (treeListTag != null)
                                    {
                                        stringBlackOut += treeListTag.blackOutDateFrom.ToString("yyyy-MM-dd") + ";"; // Дата начала отключения
                                        stringBlackOut += treeListTag.blackOutDateTo.ToString("yyyy-MM-dd") + ";"; // Дата окончания отключения
                                        stringBlackOut += treeListTag.blackOutTimeFrom.ToString("HH:mm") + ";"; // Время начала отключения
                                        stringBlackOut += treeListTag.blackOutTimeTo.ToString("HH:mm") + ";"; // Время окончания отключения
                                    }
                                    stringBlackOut += dt.Rows[row]["addressInfo"].ToString() + ";"; // Адрес
                                                                                                    //stringBlackOut += dt.Rows[row]["codeIESBK"].ToString() + "; "; // Код ИЭСБК
                                    if (!String.IsNullOrEmpty(codeIESBK)) stringBlackOut += codeIESBK + ";"; // Код ИЭСБК   
                                    else stringBlackOut += dt.Rows[row]["codeOKE"].ToString() + ";";

                                    stringBlackOut += ";"; // "пустое" поле для ИЭСБК (нужно им)

                                    stringBlackOut += dt.Rows[row]["houseFIASid"].ToString() + ";"; // ФИАС дом GUID

                                    sw.WriteLine(stringBlackOut);

                                    row_index++;
                                }

                            } // for (int row = 0; row < dt.Rows.Count; row++)

                            sw.Write(String.Format("@ИТОГО;{0}", row_index));
                        }

                        programOptions.blackoutsFileLastNumber += 1;
                        rtp3eshEntities.SaveChanges(); // убирать для отладки
                                                       //MessageBox.Show(text: "Файл успешно сохранен!", caption: "Информация");            

                        // отправка на email
                        splashScreenManager1.SetWaitFormDescription("отправка файла...");
                        MailMessage mail = new MailMessage();
                        //mail.To.Add(new MailAddress("it@oblkomenergo.ru", "it@oblkomenergo.ru"));
                        //mail.To.Add(new MailAddress("csi638@yandex.ru", "csi638@yandex.ru"));
                        mail.To.Add(new MailAddress("tkacheva_lv@es.irkutskenergo.ru", "tkacheva_lv@es.irkutskenergo.ru"));
                        mail.To.Add(new MailAddress("asalkhanova_ta@es.irkutskenergo.ru", "asalkhanova_ta@es.irkutskenergo.ru"));
                        mail.To.Add(new MailAddress("InfoCenter@es.irkutskenergo.ru", "InfoCenter@es.irkutskenergo.ru"));
                        mail.CC.Add(new MailAddress("admin@oke38.ru", "admin@oke38.ru"));
                        mail.CC.Add(new MailAddress("planoke@mail.ru", "planoke@mail.ru"));
                        mail.CC.Add(new MailAddress("davidov@oblkomenergo.ru", "davidov@oblkomenergo.ru"));
                        mail.CC.Add(new MailAddress("krupnov@oblkomenergo.ru", "krupnov@oblkomenergo.ru"));
                        mail.IsBodyHtml = true;
                        mail.BodyEncoding = Encoding.UTF8;
                        mail.SubjectEncoding = Encoding.GetEncoding(1251);

                        mail.Subject =
                            formBlackoutsTypeSelect.radioButtonPlan.Checked ? 
                            String.Format("Информация об отключениях ({0}, рейс {1})", DateTime.Now.ToShortDateString(), programOptions.blackoutsFileLastNumber)
                            : String.Format("Информация об отключениях АВАРИЯ ({0}, рейс {1})", DateTime.Now.ToShortDateString(), programOptions.blackoutsFileLastNumber);

                        //string fileBlackouts = "C:\\Users\\Serg\\Desktop\\Отключения\\9921051600.csv";            
                        Attachment attachmentBlackouts = new Attachment(fullFilenameBlackoutFile);
                        mail.Attachments.Add(attachmentBlackouts);
                        mail.Body = "<p>Автоматическое уведомление!</p><p>Не отвечайте на данное письмо!</p>";
                        mail.From = new MailAddress("robot@oke38.ru", "ОГУЭП Облкоммунэнерго", Encoding.GetEncoding(1251));
                        SmtpClient smtp = new SmtpClient("mail.nic.ru", 2525);
                        smtp.Credentials = new NetworkCredential("robot@oke38.ru", "77pHmX5TgPZ3c");
                        smtp.SendCompleted += new SendCompletedEventHandler(MailSendCompletedCallback);
                        try
                        {
                            smtp.SendAsync(mail, mail);
                        }
                        catch (SmtpFailedRecipientException) { }
                        finally
                        {

                        }

                        // возвращаем исходные значения состояний
                        mailBlackoutsSent = false;

                        splashScreenManager1.CloseWaitForm();
                    } //if (dt.Rows.Count > 0)
                    else
                    {
                        splashScreenManager1.CloseWaitForm();
                        MessageBox.Show(text: "Список потребителей пуст!", caption: "Ошибка");
                    }
                } // if (treeList1.GetAllCheckedNodes().Count > 0)
                else
                {
                    MessageBox.Show(text: "Не выбраны элементы сети!", caption: "Ошибка");
                }
            } // if (form1.ShowDialog(this) == DialogResult.OK)
        } // private void barButtonBlackoutConsumers2CSV_ItemClick(object sender, ItemClickEventArgs e)

        // перехватываем самостоятельную установку checkbox-а дерева элементов сети
        private void treeList1_BeforeCheckNode(object sender, CheckNodeEventArgs e)
        {            
            e.CanCheck = myIsAllNodesCellsFilled(e.Node);
            if (!e.CanCheck) MessageBox.Show(text: "Ошибка в параметрах отключения!", caption: "Ошибка");
            /*if (e.State == CheckState.Indeterminate)
                e.State = (e.PrevState == CheckState.Checked ? CheckState.Unchecked : CheckState.Checked);*/
        }

    } // public partial class FormBlackouts : DevExpress.XtraBars.Ribbon.RibbonForm
}