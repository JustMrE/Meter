namespace Meter.Forms
{
    partial class MonthSelect
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
            flowLayoutPanel1 = new FlowLayoutPanel();
            button1 = new Button();
            button2 = new Button();
            button3 = new Button();
            button4 = new Button();
            button5 = new Button();
            button6 = new Button();
            button7 = new Button();
            button8 = new Button();
            button9 = new Button();
            button10 = new Button();
            button11 = new Button();
            button12 = new Button();
            flowLayoutPanel1.SuspendLayout();
            SuspendLayout();
            // 
            // flowLayoutPanel1
            // 
            flowLayoutPanel1.Controls.Add(button1);
            flowLayoutPanel1.Controls.Add(button2);
            flowLayoutPanel1.Controls.Add(button3);
            flowLayoutPanel1.Controls.Add(button4);
            flowLayoutPanel1.Controls.Add(button5);
            flowLayoutPanel1.Controls.Add(button6);
            flowLayoutPanel1.Controls.Add(button7);
            flowLayoutPanel1.Controls.Add(button8);
            flowLayoutPanel1.Controls.Add(button9);
            flowLayoutPanel1.Controls.Add(button10);
            flowLayoutPanel1.Controls.Add(button11);
            flowLayoutPanel1.Controls.Add(button12);
            flowLayoutPanel1.Location = new Point(12, 12);
            flowLayoutPanel1.Name = "flowLayoutPanel1";
            flowLayoutPanel1.Size = new Size(244, 117);
            flowLayoutPanel1.TabIndex = 15;
            // 
            // button1
            // 
            button1.ForeColor = SystemColors.ControlText;
            button1.Location = new Point(3, 3);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 0;
            button1.Text = "январь";
            button1.UseVisualStyleBackColor = true;
            button1.Click += btn_Click;
            // 
            // button2
            // 
            button2.Location = new Point(84, 3);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 1;
            button2.Text = "февраль";
            button2.UseVisualStyleBackColor = true;
            button2.Click += btn_Click;
            // 
            // button3
            // 
            button3.Location = new Point(165, 3);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 2;
            button3.Text = "март";
            button3.UseVisualStyleBackColor = true;
            button3.Click += btn_Click;
            // 
            // button4
            // 
            button4.Location = new Point(3, 32);
            button4.Name = "button4";
            button4.Size = new Size(75, 23);
            button4.TabIndex = 5;
            button4.Text = "апрель";
            button4.UseVisualStyleBackColor = true;
            button4.Click += btn_Click;
            // 
            // button5
            // 
            button5.Location = new Point(84, 32);
            button5.Name = "button5";
            button5.Size = new Size(75, 23);
            button5.TabIndex = 4;
            button5.Text = "май";
            button5.UseVisualStyleBackColor = true;
            button5.Click += btn_Click;
            // 
            // button6
            // 
            button6.Location = new Point(165, 32);
            button6.Name = "button6";
            button6.Size = new Size(75, 23);
            button6.TabIndex = 3;
            button6.Text = "июнь";
            button6.UseVisualStyleBackColor = true;
            button6.Click += btn_Click;
            // 
            // button7
            // 
            button7.Location = new Point(3, 61);
            button7.Name = "button7";
            button7.Size = new Size(75, 23);
            button7.TabIndex = 8;
            button7.Text = "июль";
            button7.UseVisualStyleBackColor = true;
            button7.Click += btn_Click;
            // 
            // button8
            // 
            button8.Location = new Point(84, 61);
            button8.Name = "button8";
            button8.Size = new Size(75, 23);
            button8.TabIndex = 7;
            button8.Text = "август";
            button8.UseVisualStyleBackColor = true;
            button8.Click += btn_Click;
            // 
            // button9
            // 
            button9.Location = new Point(165, 61);
            button9.Name = "button9";
            button9.Size = new Size(75, 23);
            button9.TabIndex = 6;
            button9.Text = "сентябрь";
            button9.UseVisualStyleBackColor = true;
            button9.Click += btn_Click;
            // 
            // button10
            // 
            button10.Location = new Point(3, 90);
            button10.Name = "button10";
            button10.Size = new Size(75, 23);
            button10.TabIndex = 11;
            button10.Text = "октябрь";
            button10.UseVisualStyleBackColor = true;
            button10.Click += btn_Click;
            // 
            // button11
            // 
            button11.Location = new Point(84, 90);
            button11.Name = "button11";
            button11.Size = new Size(75, 23);
            button11.TabIndex = 10;
            button11.Text = "ноябрь";
            button11.UseVisualStyleBackColor = true;
            button11.Click += btn_Click;
            // 
            // button12
            // 
            button12.Location = new Point(165, 90);
            button12.Name = "button12";
            button12.Size = new Size(75, 23);
            button12.TabIndex = 9;
            button12.Text = "декабрь";
            button12.UseVisualStyleBackColor = true;
            button12.Click += btn_Click;
            // 
            // MonthSelect
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(267, 136);
            Controls.Add(flowLayoutPanel1);
            FormBorderStyle = FormBorderStyle.Fixed3D;
            Name = "MonthSelect";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "MonthSelect";
            Shown += MonthSelect_Shown;
            flowLayoutPanel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private FlowLayoutPanel flowLayoutPanel1;
        private Button button1;
        private Button button2;
        private Button button3;
        private Button button6;
        private Button button5;
        private Button button4;
        private Button button9;
        private Button button8;
        private Button button7;
        private Button button12;
        private Button button11;
        private Button button10;
    }
}