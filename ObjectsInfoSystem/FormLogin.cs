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
    public partial class FormLogin : DevExpress.XtraEditors.XtraForm
    {
        public FormLogin()
        {
            InitializeComponent();
        }

        private void FormLogin_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r') this.DialogResult = DialogResult.OK; // Enter
        }
    }
}