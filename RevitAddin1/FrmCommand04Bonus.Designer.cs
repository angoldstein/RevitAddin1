﻿namespace RevitAddin1
{
    partial class FrmCommand04Bonus
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbLineStyles = new System.Windows.Forms.ComboBox();
            this.cmbWallTypes = new System.Windows.Forms.ComboBox();
            this.tbxWallHeight = new System.Windows.Forms.TextBox();
            this.cbxStructural = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(0, 4);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Line Styles";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(0, 67);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Select Wall Type";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(0, 142);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(90, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Enter Wall Height";
            // 
            // cmbLineStyles
            // 
            this.cmbLineStyles.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbLineStyles.FormattingEnabled = true;
            this.cmbLineStyles.Location = new System.Drawing.Point(3, 29);
            this.cmbLineStyles.Name = "cmbLineStyles";
            this.cmbLineStyles.Size = new System.Drawing.Size(338, 21);
            this.cmbLineStyles.TabIndex = 3;
            // 
            // cmbWallTypes
            // 
            this.cmbWallTypes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbWallTypes.FormattingEnabled = true;
            this.cmbWallTypes.Location = new System.Drawing.Point(3, 100);
            this.cmbWallTypes.Name = "cmbWallTypes";
            this.cmbWallTypes.Size = new System.Drawing.Size(338, 21);
            this.cmbWallTypes.TabIndex = 4;
            // 
            // tbxWallHeight
            // 
            this.tbxWallHeight.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbxWallHeight.Location = new System.Drawing.Point(3, 167);
            this.tbxWallHeight.Name = "tbxWallHeight";
            this.tbxWallHeight.Size = new System.Drawing.Size(338, 20);
            this.tbxWallHeight.TabIndex = 5;
            // 
            // cbxStructural
            // 
            this.cbxStructural.AutoSize = true;
            this.cbxStructural.Location = new System.Drawing.Point(3, 205);
            this.cbxStructural.Name = "cbxStructural";
            this.cbxStructural.Size = new System.Drawing.Size(112, 17);
            this.cbxStructural.TabIndex = 6;
            this.cbxStructural.Text = "Is Wall Structural?";
            this.cbxStructural.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(185, 236);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 25);
            this.btnOK.TabIndex = 7;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(268, 236);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(73, 25);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // FrmCommand04Bonus
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(348, 273);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.cbxStructural);
            this.Controls.Add(this.tbxWallHeight);
            this.Controls.Add(this.cmbWallTypes);
            this.Controls.Add(this.cmbLineStyles);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "FrmCommand04Bonus";
            this.Text = "Create Walls from Lines";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbLineStyles;
        private System.Windows.Forms.ComboBox cmbWallTypes;
        private System.Windows.Forms.TextBox tbxWallHeight;
        private System.Windows.Forms.CheckBox cbxStructural;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
    }
}