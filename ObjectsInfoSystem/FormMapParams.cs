using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ObjectsInfoSystem
{
    public partial class FormMapParams : DevExpress.XtraEditors.XtraForm
    {
        public FormMapParams()
        {
            InitializeComponent();
        }

        private void FormMapParams_Shown(object sender, EventArgs e)
        {
            this.dateEditEnd.DateTime = DateTime.Now;
            DateTime datestart = DateTime.Now;
            this.dateEditStart.DateTime = Convert.ToDateTime("01." + datestart.Month + "," + datestart.Year);                     
        }
    }
}