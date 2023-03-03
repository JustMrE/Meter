
using Meter.Forms;

namespace Meter
{
    partial class Menu
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        protected override void InitializeComponent()
        {
            base.InitializeComponent();
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1834, 118);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.flowLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Menu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Form1";
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.tableLayoutPanel5.ResumeLayout(false);
            this.ResumeLayout(false);
            //
            //button 2
            //
            this.button2.Visible = true;
            this.button2.Text = "EMCOS";
            this.button2.Click += new EventHandler(Button2_Click);

            //
            //textBox 1
            //
            this.textBox1.KeyPress += new KeyPressEventHandler(this.TextBox1_KeyPress);
        }

        /* protected override void InitializeComponent()
        {
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.RepairMenu = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.button14 = new System.Windows.Forms.Button();
            this.button15 = new System.Windows.Forms.Button();
            this.button16 = new System.Windows.Forms.Button();
            this.button17 = new System.Windows.Forms.Button();
            this.button18 = new System.Windows.Forms.Button();
            this.button19 = new System.Windows.Forms.Button();
            this.button20 = new System.Windows.Forms.Button();
            this.button21 = new System.Windows.Forms.Button();
            this.button22 = new System.Windows.Forms.Button();
            this.button23 = new System.Windows.Forms.Button();
            this.button24 = new System.Windows.Forms.Button();
            this.button25 = new System.Windows.Forms.Button();
            this.button26 = new System.Windows.Forms.Button();
            this.button27 = new System.Windows.Forms.Button();
            this.button28 = new System.Windows.Forms.Button();
            this.button29 = new System.Windows.Forms.Button();
            this.button30 = new System.Windows.Forms.Button();
            this.button31 = new System.Windows.Forms.Button();
            this.button32 = new System.Windows.Forms.Button();
            this.button33 = new System.Windows.Forms.Button();
            this.button34 = new System.Windows.Forms.Button();
            this.button35 = new System.Windows.Forms.Button();
            this.button36 = new System.Windows.Forms.Button();
            this.button37 = new System.Windows.Forms.Button();
            this.button38 = new System.Windows.Forms.Button();
            this.button39 = new System.Windows.Forms.Button();
            this.button40 = new System.Windows.Forms.Button();
            this.button41 = new System.Windows.Forms.Button();
            this.button42 = new System.Windows.Forms.Button();
            this.btnAdmin = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button44 = new System.Windows.Forms.Button();
            this.button45 = new System.Windows.Forms.Button();
            this.button46 = new System.Windows.Forms.Button();
            this.button47 = new System.Windows.Forms.Button();
            this.button48 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel4 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel5 = new System.Windows.Forms.TableLayoutPanel();
            this.flowLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.tableLayoutPanel4.SuspendLayout();
            this.tableLayoutPanel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel1.AutoScroll = true;
            this.flowLayoutPanel1.BackColor = System.Drawing.SystemColors.Info;
            this.flowLayoutPanel1.Controls.Add(this.textBox1);
            this.flowLayoutPanel1.Controls.Add(this.RepairMenu);
            this.flowLayoutPanel1.Controls.Add(this.button2);
            this.flowLayoutPanel1.Controls.Add(this.button3);
            this.flowLayoutPanel1.Controls.Add(this.button4);
            this.flowLayoutPanel1.Controls.Add(this.button5);
            this.flowLayoutPanel1.Controls.Add(this.button6);
            this.flowLayoutPanel1.Controls.Add(this.button7);
            this.flowLayoutPanel1.Controls.Add(this.button8);
            this.flowLayoutPanel1.Controls.Add(this.button9);
            this.flowLayoutPanel1.Controls.Add(this.button10);
            this.flowLayoutPanel1.Controls.Add(this.button11);
            this.flowLayoutPanel1.Controls.Add(this.button12);
            this.flowLayoutPanel1.Controls.Add(this.button13);
            this.flowLayoutPanel1.Controls.Add(this.button14);
            this.flowLayoutPanel1.Controls.Add(this.button15);
            this.flowLayoutPanel1.Controls.Add(this.button16);
            this.flowLayoutPanel1.Controls.Add(this.button17);
            this.flowLayoutPanel1.Controls.Add(this.button18);
            this.flowLayoutPanel1.Controls.Add(this.button19);
            this.flowLayoutPanel1.Controls.Add(this.button20);
            this.flowLayoutPanel1.Controls.Add(this.button21);
            this.flowLayoutPanel1.Controls.Add(this.button22);
            this.flowLayoutPanel1.Controls.Add(this.button23);
            this.flowLayoutPanel1.Controls.Add(this.button24);
            this.flowLayoutPanel1.Controls.Add(this.button25);
            this.flowLayoutPanel1.Controls.Add(this.button26);
            this.flowLayoutPanel1.Controls.Add(this.button27);
            this.flowLayoutPanel1.Controls.Add(this.button28);
            this.flowLayoutPanel1.Controls.Add(this.button29);
            this.flowLayoutPanel1.Controls.Add(this.button30);
            this.flowLayoutPanel1.Controls.Add(this.button31);
            this.flowLayoutPanel1.Controls.Add(this.button32);
            this.flowLayoutPanel1.Controls.Add(this.button33);
            this.flowLayoutPanel1.Controls.Add(this.button34);
            this.flowLayoutPanel1.Controls.Add(this.button35);
            this.flowLayoutPanel1.Controls.Add(this.button36);
            this.flowLayoutPanel1.Controls.Add(this.button37);
            this.flowLayoutPanel1.Controls.Add(this.button38);
            this.flowLayoutPanel1.Controls.Add(this.button39);
            this.flowLayoutPanel1.Controls.Add(this.button40);
            this.flowLayoutPanel1.Controls.Add(this.button41);
            this.flowLayoutPanel1.Controls.Add(this.button42);
            this.flowLayoutPanel1.Controls.Add(this.btnAdmin);
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(1383, 118);
            this.flowLayoutPanel1.TabIndex = 1;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(3, 3);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(120, 23);
            this.textBox1.TabIndex = 52;
            // 
            // RepairMenu
            // 
            this.RepairMenu.Location = new System.Drawing.Point(129, 3);
            this.RepairMenu.Name = "RepairMenu";
            this.RepairMenu.Size = new System.Drawing.Size(120, 23);
            this.RepairMenu.TabIndex = 1;
            this.RepairMenu.Text = "Починить Меню";
            this.RepairMenu.UseVisualStyleBackColor = true;
            this.RepairMenu.Click += new System.EventHandler(this.RepairMenu_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(255, 3);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(120, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Button";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(381, 3);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(120, 23);
            this.button3.TabIndex = 3;
            this.button3.Text = "Button";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Visible = false;
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(507, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(120, 23);
            this.button4.TabIndex = 4;
            this.button4.Text = "Button";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Visible = false;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(633, 3);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(120, 23);
            this.button5.TabIndex = 5;
            this.button5.Text = "Button";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Visible = false;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(759, 3);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(120, 23);
            this.button6.TabIndex = 6;
            this.button6.Text = "Button";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Visible = false;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(885, 3);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(120, 23);
            this.button7.TabIndex = 7;
            this.button7.Text = "Button";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Visible = false;
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(1011, 3);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(120, 23);
            this.button8.TabIndex = 8;
            this.button8.Text = "Button";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Visible = false;
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(1137, 3);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(120, 23);
            this.button9.TabIndex = 9;
            this.button9.Text = "Button";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Visible = false;
            // 
            // button10
            // 
            this.button10.Location = new System.Drawing.Point(3, 32);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(120, 23);
            this.button10.TabIndex = 10;
            this.button10.Text = "Button";
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Visible = false;
            // 
            // button11
            // 
            this.button11.Location = new System.Drawing.Point(129, 32);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(120, 23);
            this.button11.TabIndex = 11;
            this.button11.Text = "Button";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Visible = false;
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(255, 32);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(120, 23);
            this.button12.TabIndex = 12;
            this.button12.Text = "Button";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Visible = false;
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(381, 32);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(120, 23);
            this.button13.TabIndex = 13;
            this.button13.Text = "Button";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Visible = false;
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(507, 32);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(120, 23);
            this.button14.TabIndex = 14;
            this.button14.Text = "Button";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Visible = false;
            // 
            // button15
            // 
            this.button15.Location = new System.Drawing.Point(633, 32);
            this.button15.Name = "button15";
            this.button15.Size = new System.Drawing.Size(120, 23);
            this.button15.TabIndex = 15;
            this.button15.Text = "Button";
            this.button15.UseVisualStyleBackColor = true;
            this.button15.Visible = false;
            // 
            // button16
            // 
            this.button16.Location = new System.Drawing.Point(759, 32);
            this.button16.Name = "button16";
            this.button16.Size = new System.Drawing.Size(120, 23);
            this.button16.TabIndex = 16;
            this.button16.Text = "Button";
            this.button16.UseVisualStyleBackColor = true;
            this.button16.Visible = false;
            // 
            // button17
            // 
            this.button17.Location = new System.Drawing.Point(885, 32);
            this.button17.Name = "button17";
            this.button17.Size = new System.Drawing.Size(120, 23);
            this.button17.TabIndex = 17;
            this.button17.Text = "Button";
            this.button17.UseVisualStyleBackColor = true;
            this.button17.Visible = false;
            // 
            // button18
            // 
            this.button18.Location = new System.Drawing.Point(1011, 32);
            this.button18.Name = "button18";
            this.button18.Size = new System.Drawing.Size(120, 23);
            this.button18.TabIndex = 18;
            this.button18.Text = "Button";
            this.button18.UseVisualStyleBackColor = true;
            this.button18.Visible = false;
            // 
            // button19
            // 
            this.button19.Location = new System.Drawing.Point(1137, 32);
            this.button19.Name = "button19";
            this.button19.Size = new System.Drawing.Size(120, 23);
            this.button19.TabIndex = 19;
            this.button19.Text = "Button";
            this.button19.UseVisualStyleBackColor = true;
            this.button19.Visible = false;
            // 
            // button20
            // 
            this.button20.Location = new System.Drawing.Point(3, 61);
            this.button20.Name = "button20";
            this.button20.Size = new System.Drawing.Size(120, 23);
            this.button20.TabIndex = 20;
            this.button20.Text = "Button";
            this.button20.UseVisualStyleBackColor = true;
            this.button20.Visible = false;
            // 
            // button21
            // 
            this.button21.Location = new System.Drawing.Point(129, 61);
            this.button21.Name = "button21";
            this.button21.Size = new System.Drawing.Size(120, 23);
            this.button21.TabIndex = 21;
            this.button21.Text = "Button";
            this.button21.UseVisualStyleBackColor = true;
            this.button21.Visible = false;
            // 
            // button22
            // 
            this.button22.Location = new System.Drawing.Point(255, 61);
            this.button22.Name = "button22";
            this.button22.Size = new System.Drawing.Size(120, 23);
            this.button22.TabIndex = 22;
            this.button22.Text = "Button";
            this.button22.UseVisualStyleBackColor = true;
            this.button22.Visible = false;
            // 
            // button23
            // 
            this.button23.Location = new System.Drawing.Point(381, 61);
            this.button23.Name = "button23";
            this.button23.Size = new System.Drawing.Size(120, 23);
            this.button23.TabIndex = 23;
            this.button23.Text = "Button";
            this.button23.UseVisualStyleBackColor = true;
            this.button23.Visible = false;
            // 
            // button24
            // 
            this.button24.Location = new System.Drawing.Point(507, 61);
            this.button24.Name = "button24";
            this.button24.Size = new System.Drawing.Size(120, 23);
            this.button24.TabIndex = 24;
            this.button24.Text = "Button";
            this.button24.UseVisualStyleBackColor = true;
            this.button24.Visible = false;
            // 
            // button25
            // 
            this.button25.Location = new System.Drawing.Point(633, 61);
            this.button25.Name = "button25";
            this.button25.Size = new System.Drawing.Size(120, 23);
            this.button25.TabIndex = 25;
            this.button25.Text = "Button";
            this.button25.UseVisualStyleBackColor = true;
            this.button25.Visible = false;
            // 
            // button26
            // 
            this.button26.Location = new System.Drawing.Point(759, 61);
            this.button26.Name = "button26";
            this.button26.Size = new System.Drawing.Size(120, 23);
            this.button26.TabIndex = 26;
            this.button26.Text = "Button";
            this.button26.UseVisualStyleBackColor = true;
            this.button26.Visible = false;
            // 
            // button27
            // 
            this.button27.Location = new System.Drawing.Point(885, 61);
            this.button27.Name = "button27";
            this.button27.Size = new System.Drawing.Size(120, 23);
            this.button27.TabIndex = 27;
            this.button27.Text = "Button";
            this.button27.UseVisualStyleBackColor = true;
            this.button27.Visible = false;
            // 
            // button28
            // 
            this.button28.Location = new System.Drawing.Point(1011, 61);
            this.button28.Name = "button28";
            this.button28.Size = new System.Drawing.Size(120, 23);
            this.button28.TabIndex = 28;
            this.button28.Text = "Button";
            this.button28.UseVisualStyleBackColor = true;
            this.button28.Visible = false;
            // 
            // button29
            // 
            this.button29.Location = new System.Drawing.Point(1137, 61);
            this.button29.Name = "button29";
            this.button29.Size = new System.Drawing.Size(120, 23);
            this.button29.TabIndex = 29;
            this.button29.Text = "Button";
            this.button29.UseVisualStyleBackColor = true;
            this.button29.Visible = false;
            // 
            // button30
            // 
            this.button30.Location = new System.Drawing.Point(3, 90);
            this.button30.Name = "button30";
            this.button30.Size = new System.Drawing.Size(120, 23);
            this.button30.TabIndex = 30;
            this.button30.Text = "Button";
            this.button30.UseVisualStyleBackColor = true;
            this.button30.Visible = false;
            // 
            // button31
            // 
            this.button31.Location = new System.Drawing.Point(129, 90);
            this.button31.Name = "button31";
            this.button31.Size = new System.Drawing.Size(120, 23);
            this.button31.TabIndex = 31;
            this.button31.Text = "Button";
            this.button31.UseVisualStyleBackColor = true;
            this.button31.Visible = false;
            // 
            // button32
            // 
            this.button32.Location = new System.Drawing.Point(255, 90);
            this.button32.Name = "button32";
            this.button32.Size = new System.Drawing.Size(120, 23);
            this.button32.TabIndex = 32;
            this.button32.Text = "Button";
            this.button32.UseVisualStyleBackColor = true;
            this.button32.Visible = false;
            // 
            // button33
            // 
            this.button33.Location = new System.Drawing.Point(381, 90);
            this.button33.Name = "button33";
            this.button33.Size = new System.Drawing.Size(120, 23);
            this.button33.TabIndex = 33;
            this.button33.Text = "Button";
            this.button33.UseVisualStyleBackColor = true;
            this.button33.Visible = false;
            // 
            // button34
            // 
            this.button34.Location = new System.Drawing.Point(507, 90);
            this.button34.Name = "button34";
            this.button34.Size = new System.Drawing.Size(120, 23);
            this.button34.TabIndex = 34;
            this.button34.Text = "Button";
            this.button34.UseVisualStyleBackColor = true;
            this.button34.Visible = false;
            // 
            // button35
            // 
            this.button35.Location = new System.Drawing.Point(633, 90);
            this.button35.Name = "button35";
            this.button35.Size = new System.Drawing.Size(120, 23);
            this.button35.TabIndex = 35;
            this.button35.Text = "Button";
            this.button35.UseVisualStyleBackColor = true;
            this.button35.Visible = false;
            // 
            // button36
            // 
            this.button36.Location = new System.Drawing.Point(759, 90);
            this.button36.Name = "button36";
            this.button36.Size = new System.Drawing.Size(120, 23);
            this.button36.TabIndex = 36;
            this.button36.Text = "Button";
            this.button36.UseVisualStyleBackColor = true;
            this.button36.Visible = false;
            // 
            // button37
            // 
            this.button37.Location = new System.Drawing.Point(885, 90);
            this.button37.Name = "button37";
            this.button37.Size = new System.Drawing.Size(120, 23);
            this.button37.TabIndex = 37;
            this.button37.Text = "Button";
            this.button37.UseVisualStyleBackColor = true;
            this.button37.Visible = false;
            // 
            // button38
            // 
            this.button38.Location = new System.Drawing.Point(1011, 90);
            this.button38.Name = "button38";
            this.button38.Size = new System.Drawing.Size(120, 23);
            this.button38.TabIndex = 38;
            this.button38.Text = "Button";
            this.button38.UseVisualStyleBackColor = true;
            this.button38.Visible = false;
            // 
            // button39
            // 
            this.button39.Location = new System.Drawing.Point(1137, 90);
            this.button39.Name = "button39";
            this.button39.Size = new System.Drawing.Size(120, 23);
            this.button39.TabIndex = 39;
            this.button39.Text = "Button";
            this.button39.UseVisualStyleBackColor = true;
            this.button39.Visible = false;
            // 
            // button40
            // 
            this.button40.Location = new System.Drawing.Point(3, 119);
            this.button40.Name = "button40";
            this.button40.Size = new System.Drawing.Size(120, 23);
            this.button40.TabIndex = 40;
            this.button40.Text = "Button";
            this.button40.UseVisualStyleBackColor = true;
            this.button40.Visible = false;
            // 
            // button41
            // 
            this.button41.Location = new System.Drawing.Point(129, 119);
            this.button41.Name = "button41";
            this.button41.Size = new System.Drawing.Size(120, 23);
            this.button41.TabIndex = 41;
            this.button41.Text = "Button";
            this.button41.UseVisualStyleBackColor = true;
            this.button41.Visible = false;
            // 
            // button42
            // 
            this.button42.Location = new System.Drawing.Point(255, 119);
            this.button42.Name = "button42";
            this.button42.Size = new System.Drawing.Size(120, 23);
            this.button42.TabIndex = 42;
            this.button42.Text = "Button";
            this.button42.UseVisualStyleBackColor = true;
            this.button42.Visible = false;
            // 
            // btnAdmin
            // 
            this.btnAdmin.Location = new System.Drawing.Point(381, 119);
            this.btnAdmin.Name = "btnAdmin";
            this.btnAdmin.Size = new System.Drawing.Size(120, 23);
            this.btnAdmin.TabIndex = 43;
            this.btnAdmin.Text = "Admin";
            this.btnAdmin.UseVisualStyleBackColor = true;
            this.btnAdmin.Click += new System.EventHandler(this.btnAdmin_Click);
            // 
            // button44
            // 
            this.button44.Location = new System.Drawing.Point(3, 3);
            this.button44.Name = "button44";
            this.button44.Size = new System.Drawing.Size(120, 23);
            this.button44.TabIndex = 44;
            this.button44.Text = "Архивироапть";
            this.button44.UseVisualStyleBackColor = true;
            // 
            // button45
            // 
            this.button45.Location = new System.Drawing.Point(3, 32);
            this.button45.Name = "button45";
            this.button45.Size = new System.Drawing.Size(120, 23);
            this.button45.TabIndex = 45;
            this.button45.Text = "Открыть месяц";
            this.button45.UseVisualStyleBackColor = true;
            // 
            // button46
            // 
            this.button46.Location = new System.Drawing.Point(3, 61);
            this.button46.Name = "button46";
            this.button46.Size = new System.Drawing.Size(120, 23);
            this.button46.TabIndex = 46;
            this.button46.Text = "Button";
            this.button46.UseVisualStyleBackColor = true;
            this.button46.Visible = false;
            // 
            // button47
            // 
            this.button47.Location = new System.Drawing.Point(3, 90);
            this.button47.Name = "button47";
            this.button47.Size = new System.Drawing.Size(120, 23);
            this.button47.TabIndex = 47;
            this.button47.Text = "Button";
            this.button47.UseVisualStyleBackColor = true;
            this.button47.Visible = false;
            // 
            // button48
            // 
            this.button48.Location = new System.Drawing.Point(3, 3);
            this.button48.Name = "button48";
            this.button48.Size = new System.Drawing.Size(60, 23);
            this.button48.TabIndex = 48;
            this.button48.Text = "L3";
            this.button48.UseVisualStyleBackColor = true;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(3, 3);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(246, 23);
            this.textBox2.TabIndex = 53;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 15;
            this.listBox1.Location = new System.Drawing.Point(3, 32);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(246, 0);
            this.listBox1.TabIndex = 2;
            this.listBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom;
            this.listBox1.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel3);
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel4);
            this.tableLayoutPanel2.Controls.Add(this.tableLayoutPanel5);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Right;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(1383, 0);
            this.tableLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel2.Name = "flowLayoutPanel2";
            this.tableLayoutPanel2.Size = new System.Drawing.Size(451, 0);
            this.tableLayoutPanel2.TabIndex = 54;
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.GrowStyle = TableLayoutPanelGrowStyle.AddColumns;
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel3.Controls.Add(this.button44);
            this.tableLayoutPanel3.Controls.Add(this.button45);
            this.tableLayoutPanel3.Controls.Add(this.button46);
            this.tableLayoutPanel3.Controls.Add(this.button47);
            this.tableLayoutPanel3.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel3.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.Size = new System.Drawing.Size(131, 0);
            this.tableLayoutPanel3.TabIndex = 54;
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            // 
            // tableLayoutPanel4
            // 
            this.tableLayoutPanel4.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel4.Controls.Add(this.textBox2);
            this.tableLayoutPanel4.Controls.Add(this.listBox1);
            this.tableLayoutPanel4.Location = new System.Drawing.Point(1383, 0);
            this.tableLayoutPanel4.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel4.Name = "tableLayoutPanel4";
            this.tableLayoutPanel4.Size = new System.Drawing.Size(251, 0);
            this.tableLayoutPanel4.TabIndex = 54;
            this.tableLayoutPanel4.ColumnCount = 1;
            this.tableLayoutPanel4.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            // 
            // tableLayoutPanel5
            // 
            this.tableLayoutPanel5.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel5.Controls.Add(this.button48);
            this.tableLayoutPanel5.Location = new System.Drawing.Point(382, 0);
            this.tableLayoutPanel5.Margin = new System.Windows.Forms.Padding(0);
            this.tableLayoutPanel5.Name = "tableLayoutPanel5";
            this.tableLayoutPanel5.Size = new System.Drawing.Size(69, 0);
            this.tableLayoutPanel5.TabIndex = 56;
            this.tableLayoutPanel5.ColumnCount = 1;
            this.tableLayoutPanel5.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            // 
            // Menu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1834, 118);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.flowLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Menu";
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Form1";
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel3.ResumeLayout(false);
            this.tableLayoutPanel4.ResumeLayout(false);
            this.tableLayoutPanel4.PerformLayout();
            this.tableLayoutPanel5.ResumeLayout(false);
            this.ResumeLayout(false);

        } */

        #endregion
        /* private FlowLayoutPanel flowLayoutPanel1;
        private Button RepairMenu;
        private Button button2;
        private Button button3;
        private Button button4;
        private Button button5;
        private Button button6;
        private Button button7;
        private Button button8;
        private Button button9;
        private Button button10;
        private Button button11;
        private Button button12;
        private Button button13;
        private Button button14;
        private Button button15;
        private Button button16;
        private Button button17;
        private Button button18;
        private Button button19;
        private Button button20;
        private Button button21;
        private Button button22;
        private Button button23;
        private Button button24;
        private Button button25;
        private Button button26;
        private Button button27;
        private Button button28;
        private Button button29;
        private Button button30;
        private Button button31;
        private Button button32;
        private Button button33;
        private Button button34;
        private Button button35;
        private Button button36;
        private Button button37;
        private Button button38;
        private Button button39;
        private Button button40;
        private Button button41;
        private Button button42; 
        private Button btnAdmin;
        private TextBox textBox1;
        private ListBox listBox1;
        private Button button44;
        private Button button45;
        private Button button46;
        private Button button47;
        private Button button48;
        private TextBox textBox2;
        private TableLayoutPanel tableLayoutPanel2;
        private TableLayoutPanel tableLayoutPanel3;
        private TableLayoutPanel tableLayoutPanel4;
        private TableLayoutPanel tableLayoutPanel5; */
    }
}