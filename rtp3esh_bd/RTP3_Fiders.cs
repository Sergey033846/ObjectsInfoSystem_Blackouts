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
    
    public partial class RTP3_Fiders
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RTP3_Fiders()
        {
            this.RTP3_Lines = new HashSet<RTP3_Lines>();
            this.RTP3_Nodes = new HashSet<RTP3_Nodes>();
        }
    
        public int GUID { get; set; }
        public int OwnerGUID { get; set; }
        public int ID_DB { get; set; }
        public string Name { get; set; }
        public Nullable<short> fCircuit { get; set; }
        public Nullable<System.DateTime> DBB { get; set; }
        public Nullable<System.DateTime> Dkor { get; set; }
    
        public virtual RTP3_GUID RTP3_GUID { get; set; }
        public virtual RTP3_Sections RTP3_Sections { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_Lines> RTP3_Lines { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_Nodes> RTP3_Nodes { get; set; }
    }
}
