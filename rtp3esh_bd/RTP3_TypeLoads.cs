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
    
    public partial class RTP3_TypeLoads
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RTP3_TypeLoads()
        {
            this.RTP3_LVFiders = new HashSet<RTP3_LVFiders>();
            this.RTP3_LVConsumers = new HashSet<RTP3_LVConsumers>();
        }
    
        public int ID_DB { get; set; }
        public int Ident { get; set; }
        public decimal CosFi { get; set; }
        public string Name { get; set; }
    
        public virtual RTP3_DB RTP3_DB { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_LVFiders> RTP3_LVFiders { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_LVConsumers> RTP3_LVConsumers { get; set; }
    }
}
