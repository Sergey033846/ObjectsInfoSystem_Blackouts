//------------------------------------------------------------------------------
// <auto-generated>
//     Этот код создан по шаблону.
//
//     Изменения, вносимые в этот файл вручную, могут привести к непредвиденной работе приложения.
//     Изменения, вносимые в этот файл вручную, будут перезаписаны при повторном создании кода.
// </auto-generated>
//------------------------------------------------------------------------------

namespace rtp3esh_bd
{
    using System;
    using System.Collections.Generic;
    
    public partial class RTP3_Coors
    {
        public int GUID { get; set; }
        public Nullable<int> X { get; set; }
        public Nullable<int> Y { get; set; }
        public Nullable<int> deltaX { get; set; }
        public Nullable<int> deltaY { get; set; }
        public Nullable<int> Rule_ { get; set; }
        public Nullable<int> Additional { get; set; }
        public int ID_DB { get; set; }
    
        public virtual RTP3_GUID RTP3_GUID { get; set; }
    }
}
