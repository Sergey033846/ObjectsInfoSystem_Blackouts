namespace ObjectsInfoSystem
{
    partial class FormRTP3dbLoadParams
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormRTP3dbLoadParams));
            this.comboBoxEditRTP3DB = new DevExpress.XtraEditors.LookUpEdit();
            this.RTP3DBBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.RTP3ESHDataSet = new ObjectsInfoSystem.RTP3ESHDataSet();
            this.textEditRTP3DBFile = new DevExpress.XtraEditors.TextEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.buttonRTP3DBFileSelect = new DevExpress.XtraEditors.SimpleButton();
            this.RTP3_DBTableAdapter = new ObjectsInfoSystem.RTP3ESHDataSetTableAdapters.RTP3_DBTableAdapter();
            this.buttonCancel = new DevExpress.XtraEditors.SimpleButton();
            this.buttonOk = new DevExpress.XtraEditors.SimpleButton();
            this.checkedListBoxControlRTP3Tables = new DevExpress.XtraEditors.CheckedListBoxControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.buttonCheckAllRTP3Tables = new DevExpress.XtraEditors.SimpleButton();
            this.buttonUnCheckAllRTP3Tables = new DevExpress.XtraEditors.SimpleButton();
            this.checkEditFillConsumersInfo = new DevExpress.XtraEditors.CheckEdit();
            ((System.ComponentModel.ISupportInitialize)(this.comboBoxEditRTP3DB.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RTP3DBBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.RTP3ESHDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.textEditRTP3DBFile.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkedListBoxControlRTP3Tables)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkEditFillConsumersInfo.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBoxEditRTP3DB
            // 
            this.comboBoxEditRTP3DB.Location = new System.Drawing.Point(8, 31);
            this.comboBoxEditRTP3DB.Name = "comboBoxEditRTP3DB";
            this.comboBoxEditRTP3DB.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.comboBoxEditRTP3DB.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("caption_bd", "Имя базы данных", 63, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.None, DevExpress.Utils.DefaultBoolean.Default)});
            this.comboBoxEditRTP3DB.Properties.DataSource = this.RTP3DBBindingSource;
            this.comboBoxEditRTP3DB.Properties.DisplayMember = "caption_bd";
            this.comboBoxEditRTP3DB.Properties.NullText = "";
            this.comboBoxEditRTP3DB.Properties.PopupSizeable = false;
            this.comboBoxEditRTP3DB.Properties.ValueMember = "ID_DB";
            this.comboBoxEditRTP3DB.Size = new System.Drawing.Size(325, 20);
            this.comboBoxEditRTP3DB.TabIndex = 0;
            // 
            // RTP3DBBindingSource
            // 
            this.RTP3DBBindingSource.DataMember = "RTP3_DB";
            this.RTP3DBBindingSource.DataSource = this.RTP3ESHDataSet;
            // 
            // RTP3ESHDataSet
            // 
            this.RTP3ESHDataSet.DataSetName = "RTP3ESHDataSet";
            this.RTP3ESHDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // textEditRTP3DBFile
            // 
            this.textEditRTP3DBFile.Location = new System.Drawing.Point(8, 89);
            this.textEditRTP3DBFile.Name = "textEditRTP3DBFile";
            this.textEditRTP3DBFile.Properties.ReadOnly = true;
            this.textEditRTP3DBFile.Size = new System.Drawing.Size(278, 20);
            this.textEditRTP3DBFile.TabIndex = 1;
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(8, 12);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(149, 13);
            this.labelControl1.TabIndex = 2;
            this.labelControl1.Text = "Укажите базу данных РТП-3:";
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(8, 70);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(156, 13);
            this.labelControl2.TabIndex = 3;
            this.labelControl2.Text = "Выберите файл базы данных :";
            // 
            // buttonRTP3DBFileSelect
            // 
            this.buttonRTP3DBFileSelect.Location = new System.Drawing.Point(292, 86);
            this.buttonRTP3DBFileSelect.Name = "buttonRTP3DBFileSelect";
            this.buttonRTP3DBFileSelect.Size = new System.Drawing.Size(41, 23);
            this.buttonRTP3DBFileSelect.TabIndex = 5;
            this.buttonRTP3DBFileSelect.Text = "...";
            this.buttonRTP3DBFileSelect.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // RTP3_DBTableAdapter
            // 
            this.RTP3_DBTableAdapter.ClearBeforeFill = true;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(561, 348);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(92, 23);
            this.buttonCancel.TabIndex = 6;
            this.buttonCancel.Text = "Отмена";
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(463, 348);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(92, 23);
            this.buttonOk.TabIndex = 7;
            this.buttonOk.Text = "Загрузить";
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // checkedListBoxControlRTP3Tables
            // 
            this.checkedListBoxControlRTP3Tables.CheckOnClick = true;
            this.checkedListBoxControlRTP3Tables.HighlightedItemStyle = DevExpress.XtraEditors.HighlightStyle.Standard;
            this.checkedListBoxControlRTP3Tables.Location = new System.Drawing.Point(360, 31);
            this.checkedListBoxControlRTP3Tables.LookAndFeel.SkinName = "Office 2007 Black";
            this.checkedListBoxControlRTP3Tables.LookAndFeel.UseDefaultLookAndFeel = false;
            this.checkedListBoxControlRTP3Tables.MultiColumn = true;
            this.checkedListBoxControlRTP3Tables.Name = "checkedListBoxControlRTP3Tables";
            this.checkedListBoxControlRTP3Tables.Size = new System.Drawing.Size(293, 311);
            this.checkedListBoxControlRTP3Tables.TabIndex = 8;
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(360, 12);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(171, 13);
            this.labelControl3.TabIndex = 9;
            this.labelControl3.Text = "Выберите таблицы для загрузки:";
            // 
            // buttonCheckAllRTP3Tables
            // 
            this.buttonCheckAllRTP3Tables.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("buttonCheckAllRTP3Tables.ImageOptions.Image")));
            this.buttonCheckAllRTP3Tables.Location = new System.Drawing.Point(603, 5);
            this.buttonCheckAllRTP3Tables.Name = "buttonCheckAllRTP3Tables";
            this.buttonCheckAllRTP3Tables.Size = new System.Drawing.Size(24, 22);
            this.buttonCheckAllRTP3Tables.TabIndex = 10;
            this.buttonCheckAllRTP3Tables.ToolTip = "выделить все";
            this.buttonCheckAllRTP3Tables.Click += new System.EventHandler(this.buttonCheckAllRTP3Tables_Click);
            // 
            // buttonUnCheckAllRTP3Tables
            // 
            this.buttonUnCheckAllRTP3Tables.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("buttonUnCheckAllRTP3Tables.ImageOptions.Image")));
            this.buttonUnCheckAllRTP3Tables.Location = new System.Drawing.Point(629, 5);
            this.buttonUnCheckAllRTP3Tables.Name = "buttonUnCheckAllRTP3Tables";
            this.buttonUnCheckAllRTP3Tables.Size = new System.Drawing.Size(24, 22);
            this.buttonUnCheckAllRTP3Tables.TabIndex = 11;
            this.buttonUnCheckAllRTP3Tables.ToolTip = "снять выделение";
            this.buttonUnCheckAllRTP3Tables.Click += new System.EventHandler(this.buttonUnCheckAllRTP3Tables_Click);
            // 
            // checkEditFillConsumersInfo
            // 
            this.checkEditFillConsumersInfo.Location = new System.Drawing.Point(8, 127);
            this.checkEditFillConsumersInfo.Name = "checkEditFillConsumersInfo";
            this.checkEditFillConsumersInfo.Properties.Caption = "заполнять таблицу ConsumersInfo";
            this.checkEditFillConsumersInfo.Size = new System.Drawing.Size(325, 19);
            this.checkEditFillConsumersInfo.TabIndex = 12;
            // 
            // FormRTP3dbLoadParams
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(660, 378);
            this.Controls.Add(this.checkEditFillConsumersInfo);
            this.Controls.Add(this.buttonUnCheckAllRTP3Tables);
            this.Controls.Add(this.buttonCheckAllRTP3Tables);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.checkedListBoxControlRTP3Tables);
            this.Controls.Add(this.buttonOk);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonRTP3DBFileSelect);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.textEditRTP3DBFile);
            this.Controls.Add(this.comboBoxEditRTP3DB);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "FormRTP3dbLoadParams";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Параметры загрузки данных из базы данных РТП-3";
            this.Load += new System.EventHandler(this.FormRTP3dbLoadParams_Load);
            ((System.ComponentModel.ISupportInitialize)(this.comboBoxEditRTP3DB.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RTP3DBBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.RTP3ESHDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.textEditRTP3DBFile.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkedListBoxControlRTP3Tables)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.checkEditFillConsumersInfo.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.SimpleButton buttonRTP3DBFileSelect;
        private RTP3ESHDataSet RTP3ESHDataSet;
        private System.Windows.Forms.BindingSource RTP3DBBindingSource;
        private RTP3ESHDataSetTableAdapters.RTP3_DBTableAdapter RTP3_DBTableAdapter;
        private DevExpress.XtraEditors.SimpleButton buttonCancel;
        private DevExpress.XtraEditors.SimpleButton buttonOk;
        public DevExpress.XtraEditors.LookUpEdit comboBoxEditRTP3DB;
        public DevExpress.XtraEditors.TextEdit textEditRTP3DBFile;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        public DevExpress.XtraEditors.CheckedListBoxControl checkedListBoxControlRTP3Tables;
        private DevExpress.XtraEditors.SimpleButton buttonCheckAllRTP3Tables;
        private DevExpress.XtraEditors.SimpleButton buttonUnCheckAllRTP3Tables;
        public DevExpress.XtraEditors.CheckEdit checkEditFillConsumersInfo;
    }
}