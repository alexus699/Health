namespace WindowsFormsApp1
{
    partial class Form3
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox02 = new System.Windows.Forms.TextBox();
            this.textBox03 = new System.Windows.Forms.TextBox();
            this.textBox01 = new System.Windows.Forms.TextBox();
            this.textBox04 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(16, 179);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 37);
            this.button1.TabIndex = 9;
            this.button1.Text = "Отмена";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button2_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(364, 179);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(105, 37);
            this.button2.TabIndex = 8;
            this.button2.Text = "Далее";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox02
            // 
            this.textBox02.Location = new System.Drawing.Point(369, 50);
            this.textBox02.Name = "textBox02";
            this.textBox02.Size = new System.Drawing.Size(100, 26);
            this.textBox02.TabIndex = 3;
            this.textBox02.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox01_KeyPress);
            // 
            // textBox03
            // 
            this.textBox03.Location = new System.Drawing.Point(369, 88);
            this.textBox03.Name = "textBox03";
            this.textBox03.Size = new System.Drawing.Size(100, 26);
            this.textBox03.TabIndex = 5;
            this.textBox03.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox01_KeyPress);
            // 
            // textBox01
            // 
            this.textBox01.Location = new System.Drawing.Point(369, 12);
            this.textBox01.Name = "textBox01";
            this.textBox01.Size = new System.Drawing.Size(100, 26);
            this.textBox01.TabIndex = 1;
            this.textBox01.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox01_KeyPress);
            // 
            // textBox04
            // 
            this.textBox04.Location = new System.Drawing.Point(369, 128);
            this.textBox04.Name = "textBox04";
            this.textBox04.Size = new System.Drawing.Size(100, 26);
            this.textBox04.TabIndex = 7;
            this.textBox04.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox01_KeyPress);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(238, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Максимальная оценка (ось Х):";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(341, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "Минимальное значение показателя (ось У):";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 91);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(348, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "Максимальное значение показателя (ось У):";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 131);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(152, 20);
            this.label4.TabIndex = 6;
            this.label4.Text = "Количество точек:";
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 237);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox04);
            this.Controls.Add(this.textBox01);
            this.Controls.Add(this.textBox03);
            this.Controls.Add(this.textBox02);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form3";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Новая оценка";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox02;
        private System.Windows.Forms.TextBox textBox03;
        private System.Windows.Forms.TextBox textBox01;
        private System.Windows.Forms.TextBox textBox04;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}