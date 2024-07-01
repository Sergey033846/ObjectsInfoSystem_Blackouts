namespace ObjectsInfoSystem
{
    partial class FormSetObjectFromXML
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
            this.memoEditSrcPoints = new DevExpress.XtraEditors.MemoEdit();
            this.memoEditTest = new DevExpress.XtraEditors.MemoEdit();
            this.buttonCreate = new DevExpress.XtraEditors.SimpleButton();
            this.listBoxControlLOCAL = new DevExpress.XtraEditors.ListBoxControl();
            this.listBoxControlLAYER = new DevExpress.XtraEditors.ListBoxControl();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
            ((System.ComponentModel.ISupportInitialize)(this.memoEditSrcPoints.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.memoEditTest.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.listBoxControlLOCAL)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.listBoxControlLAYER)).BeginInit();
            this.SuspendLayout();
            // 
            // memoEditSrcPoints
            // 
            this.memoEditSrcPoints.Location = new System.Drawing.Point(251, 155);
            this.memoEditSrcPoints.Name = "memoEditSrcPoints";
            this.memoEditSrcPoints.Size = new System.Drawing.Size(231, 94);
            this.memoEditSrcPoints.TabIndex = 8;
            // 
            // memoEditTest
            // 
            this.memoEditTest.Location = new System.Drawing.Point(251, 274);
            this.memoEditTest.Name = "memoEditTest";
            this.memoEditTest.Size = new System.Drawing.Size(231, 112);
            this.memoEditTest.TabIndex = 9;
            // 
            // buttonCreate
            // 
            this.buttonCreate.ButtonStyle = DevExpress.XtraEditors.Controls.BorderStyles.Style3D;
            this.buttonCreate.Location = new System.Drawing.Point(383, 392);
            this.buttonCreate.Name = "buttonCreate";
            this.buttonCreate.Size = new System.Drawing.Size(99, 31);
            this.buttonCreate.TabIndex = 10;
            this.buttonCreate.Text = "Создать";
            this.buttonCreate.Click += new System.EventHandler(this.buttonCreate_Click);
            // 
            // listBoxControlLOCAL
            // 
            this.listBoxControlLOCAL.Items.AddRange(new object[] {
            "линейный",
            "площадной",
            "точечный"});
            this.listBoxControlLOCAL.Location = new System.Drawing.Point(251, 31);
            this.listBoxControlLOCAL.LookAndFeel.SkinName = "Office 2013 Dark Gray";
            this.listBoxControlLOCAL.LookAndFeel.UseDefaultLookAndFeel = false;
            this.listBoxControlLOCAL.Name = "listBoxControlLOCAL";
            this.listBoxControlLOCAL.Size = new System.Drawing.Size(231, 70);
            this.listBoxControlLOCAL.TabIndex = 18;
            // 
            // listBoxControlLAYER
            // 
            this.listBoxControlLAYER.Location = new System.Drawing.Point(12, 31);
            this.listBoxControlLAYER.LookAndFeel.SkinName = "Office 2013 Dark Gray";
            this.listBoxControlLAYER.LookAndFeel.UseDefaultLookAndFeel = false;
            this.listBoxControlLAYER.Name = "listBoxControlLAYER";
            this.listBoxControlLAYER.Size = new System.Drawing.Size(231, 392);
            this.listBoxControlLAYER.SortOrder = System.Windows.Forms.SortOrder.Ascending;
            this.listBoxControlLAYER.TabIndex = 19;
            // 
            // labelControl1
            // 
            this.labelControl1.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.labelControl1.Location = new System.Drawing.Point(12, 12);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(130, 13);
            this.labelControl1.TabIndex = 20;
            this.labelControl1.Text = "Шаг 1. Выберите слой:";
            // 
            // labelControl2
            // 
            this.labelControl2.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.labelControl2.Location = new System.Drawing.Point(251, 12);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(178, 13);
            this.labelControl2.TabIndex = 21;
            this.labelControl2.Text = "Шаг 2. Выберите тип объекта:";
            // 
            // labelControl3
            // 
            this.labelControl3.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.labelControl3.Location = new System.Drawing.Point(251, 107);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(174, 13);
            this.labelControl3.TabIndex = 22;
            this.labelControl3.Text = "Шаг3. Введите номера точек:";
            // 
            // labelControl4
            // 
            this.labelControl4.Location = new System.Drawing.Point(251, 255);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(94, 13);
            this.labelControl4.TabIndex = 23;
            this.labelControl4.Text = "Протокол работы:";
            // 
            // labelControl5
            // 
            this.labelControl5.Location = new System.Drawing.Point(251, 121);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(88, 13);
            this.labelControl5.TabIndex = 24;
            this.labelControl5.Text = "разделитель - \",\"";
            // 
            // labelControl6
            // 
            this.labelControl6.Location = new System.Drawing.Point(251, 136);
            this.labelControl6.Name = "labelControl6";
            this.labelControl6.Size = new System.Drawing.Size(70, 13);
            this.labelControl6.TabIndex = 25;
            this.labelControl6.Text = "интервал - \"-\"";
            // 
            // FormSetObjectFromXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(495, 430);
            this.Controls.Add(this.labelControl6);
            this.Controls.Add(this.labelControl5);
            this.Controls.Add(this.labelControl4);
            this.Controls.Add(this.labelControl3);
            this.Controls.Add(this.labelControl2);
            this.Controls.Add(this.labelControl1);
            this.Controls.Add(this.listBoxControlLAYER);
            this.Controls.Add(this.listBoxControlLOCAL);
            this.Controls.Add(this.buttonCreate);
            this.Controls.Add(this.memoEditTest);
            this.Controls.Add(this.memoEditSrcPoints);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormSetObjectFromXML";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Создать объект";
            this.Load += new System.EventHandler(this.FormSetObjectFromXML_Load);
            ((System.ComponentModel.ISupportInitialize)(this.memoEditSrcPoints.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.memoEditTest.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.listBoxControlLOCAL)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.listBoxControlLAYER)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public DevExpress.XtraEditors.MemoEdit memoEditSrcPoints;
        public DevExpress.XtraEditors.MemoEdit memoEditTest;
        private DevExpress.XtraEditors.SimpleButton buttonCreate;
        public DevExpress.XtraEditors.ListBoxControl listBoxControlLOCAL;
        public DevExpress.XtraEditors.ListBoxControl listBoxControlLAYER;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        private DevExpress.XtraEditors.LabelControl labelControl6;
    }
}