namespace Inventory
{
    partial class Form1
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
            this.btnConvert = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.date1 = new System.Windows.Forms.DateTimePicker();
            this.btnAutoPOConvert = new System.Windows.Forms.Button();
            this.dateTimeAutoPOStart = new System.Windows.Forms.DateTimePicker();
            this.dateTimeAutoPOEnd = new System.Windows.Forms.DateTimePicker();
            this.btnOTB = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(358, 50);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(75, 23);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "Convert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(355, 97);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "label1";
            // 
            // date1
            // 
            this.date1.Location = new System.Drawing.Point(69, 48);
            this.date1.Name = "date1";
            this.date1.Size = new System.Drawing.Size(200, 22);
            this.date1.TabIndex = 2;
            // 
            // btnAutoPOConvert
            // 
            this.btnAutoPOConvert.Location = new System.Drawing.Point(363, 214);
            this.btnAutoPOConvert.Name = "btnAutoPOConvert";
            this.btnAutoPOConvert.Size = new System.Drawing.Size(165, 23);
            this.btnAutoPOConvert.TabIndex = 3;
            this.btnAutoPOConvert.Text = "AutoPO Convert Rate";
            this.btnAutoPOConvert.UseVisualStyleBackColor = true;
            this.btnAutoPOConvert.Click += new System.EventHandler(this.btnAutoPOConvert_Click);
            // 
            // dateTimeAutoPOStart
            // 
            this.dateTimeAutoPOStart.Location = new System.Drawing.Point(69, 212);
            this.dateTimeAutoPOStart.Name = "dateTimeAutoPOStart";
            this.dateTimeAutoPOStart.Size = new System.Drawing.Size(200, 22);
            this.dateTimeAutoPOStart.TabIndex = 4;
            // 
            // dateTimeAutoPOEnd
            // 
            this.dateTimeAutoPOEnd.Location = new System.Drawing.Point(69, 252);
            this.dateTimeAutoPOEnd.Name = "dateTimeAutoPOEnd";
            this.dateTimeAutoPOEnd.Size = new System.Drawing.Size(200, 22);
            this.dateTimeAutoPOEnd.TabIndex = 5;
            // 
            // btnOTB
            // 
            this.btnOTB.Location = new System.Drawing.Point(359, 286);
            this.btnOTB.Name = "btnOTB";
            this.btnOTB.Size = new System.Drawing.Size(168, 26);
            this.btnOTB.TabIndex = 6;
            this.btnOTB.Text = "check OTB";
            this.btnOTB.UseVisualStyleBackColor = true;
            this.btnOTB.Click += new System.EventHandler(this.BtnOTB_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnOTB);
            this.Controls.Add(this.dateTimeAutoPOEnd);
            this.Controls.Add(this.dateTimeAutoPOStart);
            this.Controls.Add(this.btnAutoPOConvert);
            this.Controls.Add(this.date1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnConvert);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker date1;
        private System.Windows.Forms.Button btnAutoPOConvert;
        private System.Windows.Forms.DateTimePicker dateTimeAutoPOStart;
        private System.Windows.Forms.DateTimePicker dateTimeAutoPOEnd;
        private System.Windows.Forms.Button btnOTB;
    }
}

