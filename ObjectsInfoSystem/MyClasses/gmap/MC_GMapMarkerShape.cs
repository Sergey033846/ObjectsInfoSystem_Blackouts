using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using GMap.NET;
using GMap.NET.WindowsForms;

namespace ObjectsInfoSystem.MyClasses.gmap
{
    enum my_MarkerType { mtSQUARE, mtCIRCLE };     // тип маркера = 0 - квадрат, 1 - круг

    // класс "маркер-фигура (квадрат, эллипс)"
    class MC_GMapMarkerShape : GMapMarker
    {
        private Pen prv_markerPen;
        private Brush prv_markerBrush;
        private my_MarkerType prv_markerType;
        private int prv_markerRadius;   // длина стороны квадрата, радиус окружности

        public string Name { get; set; }

        public MC_GMapMarkerShape(string markerName, GMap.NET.PointLatLng p, Pen pen, Brush brush, my_MarkerType mrktype, int radius) : base(p)
        {
            Name = markerName;

            Size = new System.Drawing.Size(radius, radius);
            Offset = new System.Drawing.Point(-Size.Width / 2, -Size.Height / 2);

            prv_markerPen = pen;
            prv_markerBrush = brush;
            prv_markerRadius = radius;
            prv_markerType = mrktype;
        }

        public override void OnRender(Graphics g)
        {
            Rectangle rect = new Rectangle(LocalPosition.X, LocalPosition.Y, prv_markerRadius, prv_markerRadius);

            switch (prv_markerType)
            {
                case my_MarkerType.mtSQUARE:
                    g.FillRectangle(prv_markerBrush, rect);
                    g.DrawRectangle(prv_markerPen, rect);
                    break;

                case my_MarkerType.mtCIRCLE:
                    g.FillEllipse(prv_markerBrush, rect);
                    g.DrawEllipse(prv_markerPen, rect);
                    break;
            }
        }
    } // class GMapMarkerFigure : GMapMarker      
}
