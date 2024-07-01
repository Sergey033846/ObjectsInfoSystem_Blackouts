using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using GMap.NET;
using GMap.NET.WindowsForms;

namespace ObjectsInfoSystem.MyClasses.gmap
{
    // класс методов для работы с компонентом GMapControl и его свойствами
    class MC_GMap
    {
        // поиск индекса оверлея по имени (Id)
        public static GMapOverlay GetOverlayIndexByName(GMapControl gmc, string overlay_name) 
        {
            return gmc.Overlays.First(overlay => String.Equals(overlay.Id, overlay_name));            
        }
    }
}
