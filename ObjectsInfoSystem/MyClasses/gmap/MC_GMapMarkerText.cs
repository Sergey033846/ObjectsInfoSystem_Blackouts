using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using GMap.NET.WindowsForms;

namespace ObjectsInfoSystem.MyClasses.gmap
{
    // класс "маркер-текст"
    class MC_GMapMarkerText : GMapMarker
    {
        private Brush prv_markerBrush;
        public string DrawText { get; set; }
        public float AngleText { get; set; }
        public GMap.NET.PointLatLng PointXY { get; set; }
        public string Name { get; set; }

        public MC_GMapMarkerText(string markerName, GMap.NET.PointLatLng p, string drawText, float angleText, Brush markerBrush) : base(p)
        {
            Name = markerName;
            
            DrawText = drawText;
            AngleText = angleText;
            PointXY = p;
            prv_markerBrush = markerBrush;            
        }

        public override void OnRender(Graphics g)
        {
            g.TranslateTransform(LocalPosition.X, LocalPosition.Y);
            g.RotateTransform(AngleText);

            g.DrawString(DrawText, new Font("Arial",8), prv_markerBrush, 0, 0);
            
            g.RotateTransform(-AngleText);
            g.TranslateTransform(-LocalPosition.X, -LocalPosition.Y);            
        }
    } // class MC_GMapMarkerText : GMapMarker
}
