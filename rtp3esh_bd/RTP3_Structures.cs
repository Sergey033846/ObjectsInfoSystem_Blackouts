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
    
    public partial class RTP3_Structures
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RTP3_Structures()
        {
            this.RTP3_Lines = new HashSet<RTP3_Lines>();
            this.RTP3_LVLines = new HashSet<RTP3_LVLines>();
        }
    
        public int ID_DB { get; set; }
        public int Ident { get; set; }
        public string Name { get; set; }
        public string ShortName { get; set; }
    
        public virtual RTP3_DB RTP3_DB { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_Lines> RTP3_Lines { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_LVLines> RTP3_LVLines { get; set; }
    }
}
