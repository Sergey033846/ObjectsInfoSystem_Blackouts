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
    public partial class FormLoadTypeSelect : DevExpress.XtraEditors.XtraForm
    {
        public FormLoadTypeSelect()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            switch (listBoxLoadType.SelectedValue.ToString())
            {
                case "ПАНОРАМА полная":
                    if (comboBoxEdit7.SelectedIndex == -1 || comboBoxEdit6.SelectedIndex == -1 || textEdit1.Text == "")
                    {
                        MessageBox.Show("Некорректные данные");
                    }
                    else this.DialogResult = DialogResult.OK;
                    break;
                default:
                    this.DialogResult = DialogResult.OK;
                    break;
            }
        }
    }
}