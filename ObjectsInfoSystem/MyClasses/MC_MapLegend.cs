using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ObjectsInfoSystem.MyClasses.gmap
{
    class MC_MapLegend
    {
        private MC_SQLDataProvider sqlDataProvider;

        private DataTable tableLEGEND;
                
        public MC_MapLegend(MC_SQLDataProvider SQLDataProvider)
        {
            sqlDataProvider = SQLDataProvider;

            tableLEGEND = sqlDataProvider.SelectDataFromSQL("SELECT * FROM tblPanoramaLegend ORDER BY idpnrmGROUPLAYER ASC");
        }

        // формируем карандаш по исходным параметрам
        private void GetLayerLocalStyles(int idPnrmGroupLayer, int idPnrmLocalType, int idLegendDestinationSurface, // входные параметры
                                         out Pen penLayerLocal, out Brush brushLayerLocal, out int pointType) // выходные значения
        {
            // значения по умолчанию
            penLayerLocal = new Pen(Color.Black, 1);
            brushLayerLocal = new SolidBrush(Color.Black);
            pointType = 0;
            
            DataRow[] rowsLayerLocal = tableLEGEND.Select(String.Concat("idpnrmGROUPLAYER = ", idPnrmGroupLayer.ToString(), 
                " AND idpnrmLOCALtype = ", idPnrmLocalType.ToString(), " AND idlgnddest = ", idLegendDestinationSurface.ToString()));

            if (rowsLayerLocal.Count() == 1)
            {
                DataRow rowLayerLocal = rowsLayerLocal[0];

                // формируем карандаш
                penLayerLocal.Dispose();
                penLayerLocal = new Pen(Color.FromArgb((int)rowLayerLocal["PENcolor"]), (int)rowLayerLocal["PENthickness"]);

                int pen_style = (int)rowLayerLocal["PENstyle"];
                if (pen_style == 0) penLayerLocal.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                else if (pen_style == 1) penLayerLocal.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDot;
                else if (pen_style == 2) penLayerLocal.DashStyle = System.Drawing.Drawing2D.DashStyle.DashDotDot;
                else if (pen_style == 3) penLayerLocal.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                else if (pen_style == 4) penLayerLocal.DashStyle = System.Drawing.Drawing2D.DashStyle.Solid;
                //-------------------

                // формируем кисть
                brushLayerLocal.Dispose();
                brushLayerLocal = new SolidBrush(Color.FromArgb((int)rowLayerLocal["FILLSOLIDalpha"], Color.FromArgb((int)rowLayerLocal["FILLSOLIDcolor"])));

                pointType = (int)rowLayerLocal["POINTtype"];

            } // if (rowsLayerLocal.Count() == 1)
        }

        // метод формирования изображения легенды карты для вывода в изображение        
        // входные параметры: словарь карт, таблица слоев, словарь видимости слоев, размер холста карты
        public Image GetMapLegendImage(IDictionary<int, MC_MapSource> mapSrcDictionary, DataTable table_LAYER_LOCALS_name, IDictionary<int, CheckState> layerVisibleDictionary,
                                       int canvasWidth, int canvasHeight)
        {            
            // параметры по умолчанию
            int legendRowFirstColumnWidth = 30 * 2;
            int legendRowHeight = 30;

            Font legendFont = new Font("Arial", 8);
            Brush legendFontBrush = new SolidBrush(Color.Black);

            /*Font legendCaptionFont = new Font("Arial", 8);
            Brush legendCaptionFontBrush = Brushes.Blue;
            string legendCaption = "УСЛОВНЫЕ ОБОЗНАЧЕНИЯ";*/

            int dx = 5; // смещение внутри строки объекта легенды (margin)

            // подсчитываем количество видимых слоев и максимальную длину названия в графических единицах измерения
            //int countVisibleLayers = 0;
            int countVisibleLocals = 0;
            float maxFontWidth = 0;

            Image legendImage = new Bitmap(legendRowFirstColumnWidth, legendRowHeight * countVisibleLocals + 1); // временный объект
            Graphics gLegend = Graphics.FromImage(legendImage);
            
            // дополнительно подсчитываем кол-во локаций в каждом видимом слое, чтобы рассчитать общее кол-во строк
            foreach (var layerTemp in layerVisibleDictionary)
            {
                if (layerTemp.Value == CheckState.Checked)
                {
                    //countVisibleLayers++;

                    // название слоя добавить в словарь!
                    DataRow[] rowsLayerLocal = table_LAYER_LOCALS_name.Select(String.Concat("idpnrmGROUPLAYER = ", layerTemp.Key));
                    countVisibleLocals += rowsLayerLocal.Count();

                    string layerCaption = rowsLayerLocal[0]["pnrmGROUPLAYERcapt"].ToString();
                    SizeF currentSizeF = gLegend.MeasureString(layerCaption, legendFont);

                    if (currentSizeF.Width > maxFontWidth) maxFontWidth = currentSizeF.Width;
                }
            }

            //SizeF legendCaptionSizeF = gLegend.MeasureString(legendCaption, legendCaptionFont);

            // подгоняем под размеры холста карты
            legendRowHeight = (int)Math.Round(canvasHeight * 0.4 / countVisibleLocals);

            legendImage.Dispose();
            /*legendImage = new Bitmap((int)Math.Round(legendRowFirstColumnWidth + maxFontWidth + 2*dx),
                                     (int)Math.Round(legendRowHeight * countVisibleLocals + legendCaptionSizeF.Height +2*dx + 1));*/

            legendImage = new Bitmap((int)Math.Round(legendRowFirstColumnWidth + maxFontWidth + 2 * dx),
                                     (int)Math.Round(legendRowHeight * countVisibleLocals + 2 * dx + 1.0));

            gLegend = Graphics.FromImage(legendImage);
            Rectangle gLegendRect = new Rectangle(0, 0, legendImage.Width - 1, legendImage.Height - 1);
            gLegend.FillRectangle(new SolidBrush(Color.White), gLegendRect);
            gLegend.DrawRectangle(new Pen(Color.Black), gLegendRect);
            
            // выводим заголовок легенды
            //gLegend.DrawString(legendCaption, legendCaptionFont, legendCaptionFontBrush, dx + (legendImage.Width - legendCaptionSizeF.Width) / 2.0F, dx);

            // выводим объекты легенды
            //int y = (int)Math.Round(legendCaptionSizeF.Height + 2 * dx);
            int y = 0;
                        
            foreach (DataRow rowlayer in table_LAYER_LOCALS_name.Rows)
            {
                int layerId = (int)rowlayer["idpnrmGROUPLAYER"];

                if (layerVisibleDictionary[layerId] == CheckState.Checked)
                {                    
                    string layerCaption = rowlayer["pnrmGROUPLAYERcapt"].ToString();
                    int idpnrmLOCALtype = (int)rowlayer["idpnrmLOCALtype"];                                        
                    int idlgnddest = 1;
                    
                    Pen layerPen;
                    Brush layerBrush;
                    int point_type;
                    GetLayerLocalStyles(layerId, idpnrmLOCALtype, idlgnddest, out layerPen, out layerBrush, out point_type);
                    
                    // отрисовка легенды в зависимости от типа объекта                        
                    switch (idpnrmLOCALtype)
                    {
                        case 2: // линейный 
                            gLegend.DrawLine(layerPen, new Point(dx, y + (legendRowHeight - (int)layerPen.Width) / 2),
                                                       new Point(legendRowFirstColumnWidth - dx, y + (legendRowHeight - (int)layerPen.Width) / 2));                            
                            break;

                        case 3: // площадной
                            Rectangle pointrect = new Rectangle(2*dx, y + dx, legendRowFirstColumnWidth - 4 * dx, legendRowHeight - 2 * dx);
                            gLegend.FillRectangle(layerBrush, pointrect);
                            gLegend.DrawRectangle(layerPen, pointrect);
                            break;

                        case 4: // подпись
                            break;

                        case 5: // точечный
                            int dx2 = 8; // другой размер отступа
                            
                            pointrect = new Rectangle((legendRowFirstColumnWidth - dx2) / 2, y + (legendRowHeight - dx2) / 2, dx2, dx2);
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

                    gLegend.DrawString(layerCaption, legendFont, legendFontBrush, legendRowFirstColumnWidth + dx, y + dx);
                    
                    //gLegend.DrawString("1", new Font("Arial", 8), new SolidBrush(Color.Magenta), 15 + dx, y + dx);

                    y += legendRowHeight;

                    layerPen.Dispose();
                    layerBrush.Dispose();                  

                } // if (layerVisibleDictionary[layerid] == CheckState.Checked)
                
            } // foreach (DataRow rowlayer in tableLAYERname.Rows)
                        
            gLegend.Dispose();

            return legendImage;
        }

    }
}
