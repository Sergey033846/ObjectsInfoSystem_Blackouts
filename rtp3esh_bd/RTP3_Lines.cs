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
    
    public partial class RTP3_Lines
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public RTP3_Lines()
        {
            this.RTP3_DoubleLines = new HashSet<RTP3_DoubleLines>();
        }
    
        public int GUID { get; set; }
        public int OwnerGUID { get; set; }
        public int ID_DB { get; set; }
        public int Node1 { get; set; }
        public int Node2 { get; set; }
        public string Info { get; set; }
        public int Provod { get; set; }
        public short State_ { get; set; }
        public Nullable<decimal> LineLength { get; set; }
        public Nullable<short> Kpl { get; set; }
        public Nullable<short> onBalance { get; set; }
        public Nullable<decimal> Idop1 { get; set; }
        public Nullable<decimal> Idop2 { get; set; }
        public Nullable<int> Structures { get; set; }
        public Nullable<System.DateTime> DBB { get; set; }
        public Nullable<short> TabOrder { get; set; }
    
        public virtual RTP3_Fiders RTP3_Fiders { get; set; }
        public virtual RTP3_GUID RTP3_GUID { get; set; }
        public virtual RTP3_References RTP3_References { get; set; }
        public virtual RTP3_Owners RTP3_Owners { get; set; }
        public virtual RTP3_Structures RTP3_Structures { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<RTP3_DoubleLines> RTP3_DoubleLines { get; set; }
    }
}
