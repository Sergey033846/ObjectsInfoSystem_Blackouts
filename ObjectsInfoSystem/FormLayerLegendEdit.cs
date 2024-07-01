using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using DevExpress.XtraBars.Alerter;

namespace ObjectsInfoSystem
{
    public partial class FormLayerLegendEdit : DevExpress.XtraEditors.XtraForm
    {
        int permissionAdmin; // признак прав доступа администратора

        string fexpr; // условие фильтра для выборки легенды (необходимо хранить для возможности сохранения)
        int idpnrmLOCALtype; // тип текущего объекта

        string dbconnectionString; // строка подключения к БД

        // метод обновления селекторов параметров интерфейса в зависимости от выбранных слоя, типов объекта и подложки
        public void RefreshUIComponentsState()
        {
            if (treeViewlayers.SelectedNode != null)
            {
                TreeNode selectednode = treeViewlayers.SelectedNode;
                
                if (selectednode.Tag as string != "layercaption" && listBoxControl2.SelectedIndex != -1)
                {
                    string selectednodefexpr = selectednode.Tag as string;
                    fexpr = selectednodefexpr + " AND (idlgnddest = " + (listBoxControl2.SelectedIndex+1).ToString() + ")";

                    // выбираем "легенду", соответствующую выбранным параметрам
                    string queryString1 =
                        "SELECT * " +
                        "FROM [objectsoke].[dbo].[tblPanoramaLegend] " +
                        "WHERE " + fexpr;
                    DataTable tableLEGEND = new DataTable();
                    MC_SQLDataProvider.SelectDataFromSQL(tableLEGEND, dbconnectionString, queryString1);

                    // скрытие/отображение вкладок в зависимости от типа объектов слоя, обновление селекторов
                    idpnrmLOCALtype = Convert.ToInt32(tableLEGEND.Rows[0]["idpnrmLOCALtype"]);

                    switch (idpnrmLOCALtype)
                    {
                        case 2: // линейный
                            xtraTabPageLine.PageVisible = true;
                            xtraTabPageFill.PageVisible = false;
                            xtraTabPagePoint.PageVisible = false;
                            xtraTabPageText.PageVisible = false;

                            // перо
                            colorEdit1.Color = Color.FromArgb(Convert.ToInt32(tableLEGEND.Rows[0]["PENcolor"]));
                            styleEdit1.SelectedIndex = Convert.ToInt32(tableLEGEND.Rows[0]["PENstyle"]);
                            thicknessEdit1.Value = Convert.ToInt32(tableLEGEND.Rows[0]["PENthickness"]);
                            break;

                        case 3: // площадной
                            xtraTabPageLine.PageVisible = true;
                            xtraTabPageFill.PageVisible = true;
                            xtraTabPagePoint.PageVisible = false;
                            xtraTabPageText.PageVisible = false;

                            // перо
                            colorEdit1.Color = Color.FromArgb(Convert.ToInt32(tableLEGEND.Rows[0]["PENcolor"]));
                            styleEdit1.SelectedIndex = Convert.ToInt32(tableLEGEND.Rows[0]["PENstyle"]);
                            thicknessEdit1.Value = Convert.ToInt32(tableLEGEND.Rows[0]["PENthickness"]);

                            // кисть
                            colorEdit2.Color = Color.FromArgb(Convert.ToInt32(tableLEGEND.Rows[0]["FILLSOLIDcolor"]));
                            alphaEdit1.Value = Convert.ToInt32(tableLEGEND.Rows[0]["FILLSOLIDalpha"]); 
                            break;

                        case 4: // подпись
                            xtraTabPageLine.PageVisible = false;
                            xtraTabPageFill.PageVisible = false;
                            xtraTabPagePoint.PageVisible = false;
                            xtraTabPageText.PageVisible = true;
                            break;

                        case 5: // точечный
                            xtraTabPageLine.PageVisible = true;
                            xtraTabPageFill.PageVisible = true;
                            xtraTabPagePoint.PageVisible = true;
                            xtraTabPageText.PageVisible = false;

                            // перо
                            colorEdit1.Color = Color.FromArgb(Convert.ToInt32(tableLEGEND.Rows[0]["PENcolor"]));
                            styleEdit1.SelectedIndex = Convert.ToInt32(tableLEGEND.Rows[0]["PENstyle"]);
                            thicknessEdit1.Value = Convert.ToInt32(tableLEGEND.Rows[0]["PENthickness"]);

                            // кисть
                            colorEdit2.Color = Color.FromArgb(Convert.ToInt32(tableLEGEND.Rows[0]["FILLSOLIDcolor"]));
                            alphaEdit1.Value = Convert.ToInt32(tableLEGEND.Rows[0]["FILLSOLIDalpha"]);

                            // фигура, сторона, радиус
                            pointtypeEdit.SelectedIndex = Convert.ToInt32(tableLEGEND.Rows[0]["POINTtype"]);
                            radiusEdit1.Value = Convert.ToInt32(tableLEGEND.Rows[0]["POINTradius"]);
                            break;
                    }                                         
                    //----------------------

                    tableLEGEND.Dispose();

                    if (permissionAdmin == 100 || permissionAdmin ==2) buttonSaveLegend.Enabled = true;
                } // if (selectednode.Tag as string != "layercaption" && listBoxControl2.SelectedIndex != -1)
                else
                {
                    fexpr = null;
                    buttonSaveLegend.Enabled = false;
                }
            }
        }

        public FormLayerLegendEdit(string constr, int permissionAdmin_parent)
        {
            InitializeComponent();

            fexpr = null;
            dbconnectionString = constr;
            permissionAdmin = permissionAdmin_parent;

            // загрузка слоев, типов объектов
            string queryString1 =
                "SELECT * " +
                "FROM [objectsoke].[dbo].[tblPanoramaGroupLayer] tblPL " +                
                "ORDER BY tblPL.idpnrmGROUPLAYER ASC";
            DataTable tableLAYERS = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableLAYERS, dbconnectionString, queryString1);

            queryString1 =
                "SELECT DISTINCT tblPL.idpnrmGROUPLAYER,tblPL.idpnrmLOCALtype," +
                "tblGL.pnrmGROUPLAYERcapt,tblLT.pnrmLOCALtypecapt " +
                "FROM[objectsoke].[dbo].[tblPanoramaLegend] tblPL " +
                "INNER JOIN tblPanoramaGroupLayer tblGL ON tblPL.idpnrmGROUPLAYER = tblGL.idpnrmGROUPLAYER " +
                "INNER JOIN tblPanoramaLOCALType tblLT ON tblPL.idpnrmLOCALtype = tblLT.idpnrmLOCALtype " +
                "ORDER BY tblPL.idpnrmGROUPLAYER ASC, tblPL.idpnrmLOCALtype ASC";
            DataTable tableLAYERSLOCALS = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableLAYERSLOCALS, dbconnectionString, queryString1);

            treeViewlayers.Nodes.Clear();
            foreach (DataRow rowLAYERS in tableLAYERS.Rows)
            {
                string layerid = rowLAYERS["idpnrmGROUPLAYER"].ToString();
                string layercaption = rowLAYERS["pnrmGROUPLAYERcapt"].ToString();

                DataRow[] LAYERlocals = tableLAYERSLOCALS.Select("idpnrmGROUPLAYER = " + layerid + "");

                if (LAYERlocals.Count() > 0)
                {
                    TreeNode[] childnodes = new TreeNode[LAYERlocals.Count()];
                    for (int j = 0; j < LAYERlocals.Count(); j++)
                    {    
                        string captionchildnode = LAYERlocals[j]["pnrmLOCALtypecapt"].ToString();
                        childnodes[j] = new TreeNode(captionchildnode);

                        string childTag = "(idpnrmGROUPLAYER = " + layerid + ") AND (idpnrmLOCALtype = " + LAYERlocals[j]["idpnrmLOCALtype"].ToString() + ")";
                        //childnodes[j].Tag = LAYERlocals[j];
                        childnodes[j].Tag = childTag;
                    }

                    TreeNode nodelayer = new TreeNode(layercaption, childnodes);
                    nodelayer.Tag = "layercaption";

                    treeViewlayers.Nodes.Add(nodelayer);
                }
            } // foreach (DataRow rowLAYERS in tableLAYERS.Rows)

            tableLAYERS.Dispose();
            tableLAYERSLOCALS.Dispose();
            //--------------------------------------

            // загрузка подложек
            listBoxControl2.Items.Clear();

            queryString1 =
                "SELECT * " +
                "FROM [objectsoke].[dbo].[tblLegendDestination] " +
                "ORDER BY idlgnddest ASC";
            DataTable tableLEGENDdest = new DataTable();
            MC_SQLDataProvider.SelectDataFromSQL(tableLEGENDdest, dbconnectionString, queryString1);

            foreach (DataRow rowLEGENDdest in tableLEGENDdest.Rows)
            {
                listBoxControl2.Items.Add(rowLEGENDdest["lgnddestcapt"].ToString());
            }

            listBoxControl2.SelectedIndex = -1;

            tableLEGENDdest.Dispose();
            //--------------------------------------

            pictureBox1.Image = new Bitmap(pictureBox1.Width, pictureBox1.Height);            
        }

        // метод отрисовки окна предварительного просмотра
        // local = 0 (линия, контур), 1 (заливка), 2 (текст)
        private void DrawStylePreview()
        {            
            Graphics g = Graphics.FromImage(pictureBox1.Image);
            
            // формируем стиль контура
            int line_width = Convert.ToInt32(thicknessEdit1.Value);
            Pen layerPen = new Pen(colorEdit1.Color, line_width);
                                
            if (styleEdit1.SelectedIndex == 0) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
            else if (styleEdit1.SelectedIndex == 1) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDot;
            else if (styleEdit1.SelectedIndex == 2) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDotDot;
            else if (styleEdit1.SelectedIndex == 3) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
            else if (styleEdit1.SelectedIndex == 4) layerPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            
            // очитска холста
            g.FillRectangle(new SolidBrush(Color.White), pictureBox1.ClientRectangle);

            if (xtraTabPagePoint.PageVisible)
            {
                int dx = (int)radiusEdit1.Value;

                Brush layerBrush = new SolidBrush(Color.FromArgb(Convert.ToInt32(alphaEdit1.Value), colorEdit2.Color));

                // очитска холста
                g.FillRectangle(new SolidBrush(Color.White), pictureBox1.ClientRectangle);

                Rectangle pointrect = new Rectangle((pictureBox1.Width - dx) / 2, (pictureBox1.Height - dx) / 2, dx, dx);

                switch (pointtypeEdit.SelectedIndex)
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
            }
            else if (xtraTabPageFill.PageVisible)
            {
                int dx = 10;
                                
                Brush layerBrush = new SolidBrush(Color.FromArgb(Convert.ToInt32(alphaEdit1.Value), colorEdit2.Color));
                //Brush layerBrush = new HatchBrush((HatchStyle)9, Color.Yellow, colorEdit2.Color);                    

                // очитска холста
                g.FillRectangle(new SolidBrush(Color.White), pictureBox1.ClientRectangle);

                //g.DrawString("ТЕСТ", new Font("Arial", 8, FontStyle.Bold), new SolidBrush(Color.Black), 10, 10); // для теста прозрачности
                g.FillRectangle(layerBrush, new Rectangle(dx, dx, pictureBox1.Width - dx * 2, pictureBox1.Height - dx * 2));
                g.DrawRectangle(layerPen, new Rectangle(dx, dx, pictureBox1.Width - dx * 2, pictureBox1.Height - dx * 2));
            }
            else
            {
                int dx = 5;
                g.DrawLine(layerPen, new Point(dx, (pictureBox1.Height - line_width) / 2), new Point(pictureBox1.Width - dx, (pictureBox1.Height - line_width) / 2));
            }
            
            layerPen.Dispose();
            g.Dispose();

            pictureBox1.Invalidate();                    
        }

        private void thicknessEdit1_ValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void styleEdit1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void colorEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void colorEdit2_EditValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void alphaEdit1_ValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void thicknessEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void alphaEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void listBoxControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshUIComponentsState();
        }

        private void treeViewlayers_AfterSelect(object sender, TreeViewEventArgs e)
        {
            RefreshUIComponentsState();
        }

        // сохранить текущую легенду
        private void buttonSaveLegend_Click(object sender, EventArgs e)
        {
            string queryString1 = "";

            if (idpnrmLOCALtype == 2 || idpnrmLOCALtype == 3 || idpnrmLOCALtype == 5)
            {
                queryString1 =
                "UPDATE [objectsoke].[dbo].[tblPanoramaLegend] " +
                "SET PENcolor = " + colorEdit1.Color.ToArgb().ToString() + ", PENthickness = " + thicknessEdit1.Value.ToString() + ", PENstyle = " + styleEdit1.SelectedIndex.ToString() +
                " WHERE " + fexpr;

                MC_SQLDataProvider.UpdateSQLQuery(dbconnectionString, queryString1);
            }

            if (idpnrmLOCALtype == 3 || idpnrmLOCALtype == 5)
            {
                queryString1 =
                "UPDATE [objectsoke].[dbo].[tblPanoramaLegend] " +
                "SET FILLSOLIDcolor = " + colorEdit2.Color.ToArgb().ToString() + ", FILLSOLIDalpha = " + alphaEdit1.Value.ToString() +
                " WHERE " + fexpr;

                MC_SQLDataProvider.UpdateSQLQuery(dbconnectionString, queryString1);
            }

            if (idpnrmLOCALtype == 5)
            {
                queryString1 =
                "UPDATE [objectsoke].[dbo].[tblPanoramaLegend] " +
                "SET POINTradius = " + radiusEdit1.Value.ToString() + ", POINTtype = " + pointtypeEdit.SelectedIndex.ToString() +
                " WHERE " + fexpr;

                MC_SQLDataProvider.UpdateSQLQuery(dbconnectionString, queryString1);
            }

            AlertInfo info = new AlertInfo("Легенда карты", "Данные успешно сохранены");
            alertControl1.Show(this, info);
        }

        private void radiusEdit1_EditValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void radiusEdit1_ValueChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }

        private void pointtypeEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            DrawStylePreview();
        }
    }
}