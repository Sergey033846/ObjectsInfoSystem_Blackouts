using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using GMap.NET;
using GMap.NET.WindowsForms;

namespace ObjectsInfoSystem.MyClasses.gmap
{
    // класс "маркер-изображение"
    class MC_GMapMarkerImage : GMapMarker
    {
        private Image image;
        public Image Image
        {
            get
            {
                return image;
            }
            set
            {
                image = value;
                if (image != null)
                {
                    this.Size = new Size(image.Width, image.Height);
                }
            }
        }

        public MC_GMapMarkerImage(GMap.NET.PointLatLng p, Image image) : base(p)
            {
            Size = new System.Drawing.Size(image.Width, image.Height);
            Offset = new System.Drawing.Point(-Size.Width / 2, -Size.Height / 2);
            this.image = image;
        }

        public override void OnRender(Graphics g)
        {
            if (image != null)
            {
                Rectangle rect = new Rectangle(LocalPosition.X, LocalPosition.Y, Size.Width, Size.Height);
                g.DrawImage(image, rect);
            }
        }
    }
}
