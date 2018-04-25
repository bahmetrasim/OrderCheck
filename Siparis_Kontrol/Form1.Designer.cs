namespace Siparis_Kontrol
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
            this.ChooseSheets = new System.Windows.Forms.ComboBox();
            this.Choose = new System.Windows.Forms.Button();
            this.Path = new System.Windows.Forms.TextBox();
            this.epplus = new System.Windows.Forms.Button();
            this.handling = new System.Windows.Forms.CheckBox();
            this.tbpaintwidth = new System.Windows.Forms.TextBox();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.Termin = new System.Windows.Forms.GroupBox();
            this.backlog = new System.Windows.Forms.RadioButton();
            this.allorders = new System.Windows.Forms.RadioButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tborderthick = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.NetProductTime = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.NeededProductTime = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.CALC = new System.Windows.Forms.Button();
            this.AVHours = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbquantity = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbstatus = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbstatusname = new System.Windows.Forms.TextBox();
            this.Termin.SuspendLayout();
            this.SuspendLayout();
            // 
            // ChooseSheets
            // 
            this.ChooseSheets.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.ChooseSheets.FormattingEnabled = true;
            this.ChooseSheets.Location = new System.Drawing.Point(504, 32);
            this.ChooseSheets.Name = "ChooseSheets";
            this.ChooseSheets.Size = new System.Drawing.Size(93, 24);
            this.ChooseSheets.TabIndex = 8;
            // 
            // Choose
            // 
            this.Choose.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Choose.Location = new System.Drawing.Point(42, 30);
            this.Choose.Name = "Choose";
            this.Choose.Size = new System.Drawing.Size(100, 25);
            this.Choose.TabIndex = 7;
            this.Choose.Text = "Choose File";
            this.Choose.UseVisualStyleBackColor = true;
            this.Choose.Click += new System.EventHandler(this.Choose_Click_1);
            // 
            // Path
            // 
            this.Path.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Path.Location = new System.Drawing.Point(159, 32);
            this.Path.Name = "Path";
            this.Path.Size = new System.Drawing.Size(316, 23);
            this.Path.TabIndex = 6;
            // 
            // epplus
            // 
            this.epplus.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.epplus.Location = new System.Drawing.Point(265, 127);
            this.epplus.Name = "epplus";
            this.epplus.Size = new System.Drawing.Size(100, 25);
            this.epplus.TabIndex = 9;
            this.epplus.Text = "EPPLUS";
            this.epplus.UseVisualStyleBackColor = true;
            this.epplus.Click += new System.EventHandler(this.epplus_Click_1);
            // 
            // handling
            // 
            this.handling.AutoSize = true;
            this.handling.Location = new System.Drawing.Point(12, 187);
            this.handling.Name = "handling";
            this.handling.Size = new System.Drawing.Size(130, 17);
            this.handling.TabIndex = 13;
            this.handling.Text = "Terminsizleri Hariç Tut";
            this.handling.UseVisualStyleBackColor = true;
            // 
            // tbpaintwidth
            // 
            this.tbpaintwidth.Location = new System.Drawing.Point(557, 161);
            this.tbpaintwidth.Name = "tbpaintwidth";
            this.tbpaintwidth.Size = new System.Drawing.Size(51, 20);
            this.tbpaintwidth.TabIndex = 14;
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.radioButton1.Location = new System.Drawing.Point(54, 89);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(96, 20);
            this.radioButton1.TabIndex = 10;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "QLIKVIEW";
            this.radioButton1.UseVisualStyleBackColor = true;
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.radioButton2.Location = new System.Drawing.Point(458, 89);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(67, 20);
            this.radioButton2.TabIndex = 16;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "AÜBT";
            this.radioButton2.UseVisualStyleBackColor = true;
            // 
            // Termin
            // 
            this.Termin.Controls.Add(this.backlog);
            this.Termin.Controls.Add(this.allorders);
            this.Termin.Location = new System.Drawing.Point(4, 115);
            this.Termin.Name = "Termin";
            this.Termin.Size = new System.Drawing.Size(155, 66);
            this.Termin.TabIndex = 17;
            this.Termin.TabStop = false;
            this.Termin.Text = "Termin";
            // 
            // backlog
            // 
            this.backlog.AutoSize = true;
            this.backlog.Location = new System.Drawing.Point(13, 41);
            this.backlog.Name = "backlog";
            this.backlog.Size = new System.Drawing.Size(133, 17);
            this.backlog.TabIndex = 1;
            this.backlog.TabStop = true;
            this.backlog.Text = "Mevcut Ay ve Backlog";
            this.backlog.UseVisualStyleBackColor = true;
            // 
            // allorders
            // 
            this.allorders.AutoSize = true;
            this.allorders.Location = new System.Drawing.Point(13, 18);
            this.allorders.Name = "allorders";
            this.allorders.Size = new System.Drawing.Size(91, 17);
            this.allorders.TabIndex = 0;
            this.allorders.TabStop = true;
            this.allorders.Text = "Tüm Siparişler";
            this.allorders.UseVisualStyleBackColor = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(453, 164);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 15);
            this.label1.TabIndex = 18;
            this.label1.Text = "Giriş Eni - Kolonu";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(413, 135);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(138, 15);
            this.label2.TabIndex = 20;
            this.label2.Text = "Sipariş Kalınlığı- Kolonu";
            // 
            // tborderthick
            // 
            this.tborderthick.Location = new System.Drawing.Point(557, 133);
            this.tborderthick.Name = "tborderthick";
            this.tborderthick.Size = new System.Drawing.Size(51, 20);
            this.tborderthick.TabIndex = 19;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(91, 274);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(51, 20);
            this.textBox3.TabIndex = 21;
            // 
            // NetProductTime
            // 
            this.NetProductTime.AutoSize = true;
            this.NetProductTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.NetProductTime.Location = new System.Drawing.Point(14, 235);
            this.NetProductTime.Name = "NetProductTime";
            this.NetProductTime.Size = new System.Drawing.Size(209, 17);
            this.NetProductTime.TabIndex = 22;
            this.NetProductTime.Text = "Toplam Net Üretim Süresi : ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label3.Location = new System.Drawing.Point(226, 237);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(0, 17);
            this.label3.TabIndex = 23;
            // 
            // NeededProductTime
            // 
            this.NeededProductTime.AutoSize = true;
            this.NeededProductTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.NeededProductTime.Location = new System.Drawing.Point(253, 273);
            this.NeededProductTime.Name = "NeededProductTime";
            this.NeededProductTime.Size = new System.Drawing.Size(144, 17);
            this.NeededProductTime.TabIndex = 24;
            this.NeededProductTime.Text = "İhtiyaç Süre (Saat)";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label4.Location = new System.Drawing.Point(423, 273);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 17);
            this.label4.TabIndex = 25;
            // 
            // CALC
            // 
            this.CALC.Location = new System.Drawing.Point(159, 270);
            this.CALC.Name = "CALC";
            this.CALC.Size = new System.Drawing.Size(53, 26);
            this.CALC.TabIndex = 26;
            this.CALC.Text = "CALC";
            this.CALC.UseVisualStyleBackColor = true;
            // 
            // AVHours
            // 
            this.AVHours.AutoSize = true;
            this.AVHours.Location = new System.Drawing.Point(1, 277);
            this.AVHours.Name = "AVHours";
            this.AVHours.Size = new System.Drawing.Size(79, 13);
            this.AVHours.TabIndex = 27;
            this.AVHours.Text = "Av.Hours Ratio";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label5.Location = new System.Drawing.Point(392, 199);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(163, 15);
            this.label5.TabIndex = 29;
            this.label5.Text = "Baz Alınacak Miktar - Kolonu";
            // 
            // tbquantity
            // 
            this.tbquantity.Location = new System.Drawing.Point(557, 198);
            this.tbquantity.Name = "tbquantity";
            this.tbquantity.Size = new System.Drawing.Size(51, 20);
            this.tbquantity.TabIndex = 28;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label6.Location = new System.Drawing.Point(412, 231);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(142, 15);
            this.label6.TabIndex = 31;
            this.label6.Text = "Genel Açıklama - Kolonu";
            // 
            // tbstatus
            // 
            this.tbstatus.Location = new System.Drawing.Point(557, 230);
            this.tbstatus.Name = "tbstatus";
            this.tbstatus.Size = new System.Drawing.Size(51, 20);
            this.tbstatus.TabIndex = 30;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label7.Location = new System.Drawing.Point(429, 258);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(54, 15);
            this.label7.TabIndex = 33;
            this.label7.Text = "Filtre Adı";
            // 
            // tbstatusname
            // 
            this.tbstatusname.Location = new System.Drawing.Point(504, 257);
            this.tbstatusname.Name = "tbstatusname";
            this.tbstatusname.Size = new System.Drawing.Size(130, 20);
            this.tbstatusname.TabIndex = 32;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(646, 306);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.tbstatusname);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.tbstatus);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tbquantity);
            this.Controls.Add(this.AVHours);
            this.Controls.Add(this.CALC);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.NeededProductTime);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.NetProductTime);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.tborderthick);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Termin);
            this.Controls.Add(this.radioButton2);
            this.Controls.Add(this.tbpaintwidth);
            this.Controls.Add(this.handling);
            this.Controls.Add(this.radioButton1);
            this.Controls.Add(this.epplus);
            this.Controls.Add(this.ChooseSheets);
            this.Controls.Add(this.Choose);
            this.Controls.Add(this.Path);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Termin.ResumeLayout(false);
            this.Termin.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox ChooseSheets;
        private System.Windows.Forms.Button Choose;
        private System.Windows.Forms.TextBox Path;
        private System.Windows.Forms.Button epplus;
        private System.Windows.Forms.CheckBox handling;
        private System.Windows.Forms.TextBox tbpaintwidth;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.GroupBox Termin;
        private System.Windows.Forms.RadioButton backlog;
        private System.Windows.Forms.RadioButton allorders;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tborderthick;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label NetProductTime;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label NeededProductTime;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button CALC;
        private System.Windows.Forms.Label AVHours;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbquantity;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbstatus;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox tbstatusname;
    }
}

