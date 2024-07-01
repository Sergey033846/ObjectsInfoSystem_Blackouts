namespace ObjectsInfoSystem
{
    partial class FormBlackoutsTypeSelect
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.radioButtonPlan = new System.Windows.Forms.RadioButton();
            this.radioButtonAvar = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button1.Location = new System.Drawing.Point(32, 119);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(87, 31);
            this.button1.TabIndex = 10;
            this.button1.Text = "Отправить";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Location = new System.Drawing.Point(125, 119);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(87, 31);
            this.button2.TabIndex = 11;
            this.button2.Text = "Отменить";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // radioButtonPlan
            // 
            this.radioButtonPlan.AutoSize = true;
            this.radioButtonPlan.Checked = true;
            this.radioButtonPlan.Location = new System.Drawing.Point(32, 32);
            this.radioButtonPlan.Name = "radioButtonPlan";
            this.radioButtonPlan.Size = new System.Drawing.Size(140, 17);
            this.radioButtonPlan.TabIndex = 12;
            this.radioButtonPlan.TabStop = true;
            this.radioButtonPlan.Text = "Плановое отключение";
            this.radioButtonPlan.UseVisualStyleBackColor = true;
            // 
            // radioButtonAvar
            // 
            this.radioButtonAvar.AutoSize = true;
            this.radioButtonAvar.Location = new System.Drawing.Point(32, 69);
            this.radioButtonAvar.Name = "radioButtonAvar";
            this.radioButtonAvar.Size = new System.Drawing.Size(146, 17);
            this.radioButtonAvar.TabIndex = 13;
            this.radioButtonAvar.Text = "Аварийное отключение";
            this.radioButtonAvar.UseVisualStyleBackColor = true;
            // 
            // FormBlackoutsTypeSelect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(222, 162);
            this.Controls.Add(this.radioButtonAvar);
            this.Controls.Add(this.radioButtonPlan);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormBlackoutsTypeSelect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Укажите вид отключения";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        public System.Windows.Forms.RadioButton radioButtonPlan;
        public System.Windows.Forms.RadioButton radioButtonAvar;
    }
}