using rtp3esh_bd;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.Entity;

namespace ObjectsInfoSystem
{
    public partial class FormRTP3dbLoadParams : DevExpress.XtraEditors.XtraForm
    {
        public FormRTP3dbLoadParams()
        {
            InitializeComponent();
        }

        private void FormRTP3dbLoadParams_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "rTP3ESHDataSet.RTP3_DB". При необходимости она может быть перемещена или удалена.
            this.RTP3_DBTableAdapter.Fill(this.RTP3ESHDataSet.RTP3_DB);

            // заполняем список таблиц для обновление на основе схемы БД
            SqlConnection SQLconnection = new SqlConnection();
            SQLconnection.ConnectionString = "Data Source=SERVERPFB;Initial Catalog=RTP3ESH;User ID=rtp3esh_admin;Password=ObjectsRtp3Esh2021";
            SqlCommand SQLcommand = new SqlCommand();
            SQLcommand.Connection = SQLconnection;
            SqlDataAdapter SQLdataAdapter = new SqlDataAdapter();
            SQLdataAdapter.SelectCommand = SQLcommand;
            DataTable dt = new DataTable();
            SQLconnection.Open();
            SQLcommand.CommandText = "USE RTP3ESH SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE table_type='BASE TABLE' ORDER BY TABLE_NAME";
            SQLdataAdapter.Fill(dt);
            SQLconnection.Close();
            foreach (DataRow dataRow in dt.Rows)
            {
                string tableNameMy = dataRow["TABLE_NAME"].ToString();
                string tableNameRTP3 = tableNameMy.Replace("RTP3_", String.Empty);
                if (tableNameMy.Contains("RTP3_") && !tableNameMy.Equals("RTP3_DB")) checkedListBoxControlRTP3Tables.Items.Add(tableNameRTP3, tableNameRTP3, CheckState.Unchecked, true);                
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialogRTP3DBFile = new OpenFileDialog();
            openFileDialogRTP3DBFile.Title = "Выберите файл базы данных РТП-3";
            openFileDialogRTP3DBFile.Filter = "База данных РТП-3 (*.gdb)|*.gdb";
            openFileDialogRTP3DBFile.FileName = "";

            if (openFileDialogRTP3DBFile.ShowDialog() == DialogResult.OK)
            {
                textEditRTP3DBFile.Text = openFileDialogRTP3DBFile.FileName;
            }
        }

        // кнопка "загрузить"
        private void buttonOk_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(comboBoxEditRTP3DB.Text) || String.IsNullOrEmpty(textEditRTP3DBFile.Text))
            {
                MessageBox.Show(text: "Проверьте выбранные настройки!", caption: "Ошибка");
            }
            else
            {
                this.DialogResult = DialogResult.OK;
            }
        }

        private void buttonCheckAllRTP3Tables_Click(object sender, EventArgs e)
        {
            checkedListBoxControlRTP3Tables.CheckAll();
        }

        private void buttonUnCheckAllRTP3Tables_Click(object sender, EventArgs e)
        {
            checkedListBoxControlRTP3Tables.UnCheckAll();
        }
    }
}