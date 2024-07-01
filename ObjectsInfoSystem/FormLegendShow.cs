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
    public partial class FormLegendShow : DevExpress.XtraEditors.XtraForm
    {        
        public FormLegendShow()
        {
            InitializeComponent();
        }

        private void FormLegendShow_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Visible = false;
            e.Cancel = true;
        }
    }
}