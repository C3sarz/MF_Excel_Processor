namespace MF_Excel_Processor
{
    partial class Main_Form
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
            this.OpenFileButton = new System.Windows.Forms.Button();
            this.MainTextBox = new System.Windows.Forms.TextBox();
            this.DataTextBox = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.CleanupButton = new System.Windows.Forms.Button();
            this.CellSearchTextBox = new System.Windows.Forms.TextBox();
            this.cellBox1 = new System.Windows.Forms.TextBox();
            this.cellBox2 = new System.Windows.Forms.TextBox();
            this.RowConfirmButton = new System.Windows.Forms.Button();
            this.RowBox1 = new System.Windows.Forms.TextBox();
            this.FilterStartButton = new System.Windows.Forms.Button();
            this.RowBox2 = new System.Windows.Forms.TextBox();
            this.ListBox = new System.Windows.Forms.TextBox();
            this.MissingBox = new System.Windows.Forms.TextBox();
            this.YearTextBox = new System.Windows.Forms.TextBox();
            this.MonthTextBox = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.CancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Location = new System.Drawing.Point(84, 146);
            this.OpenFileButton.Margin = new System.Windows.Forms.Padding(4);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(189, 64);
            this.OpenFileButton.TabIndex = 0;
            this.OpenFileButton.Text = "Abrir archivo";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // MainTextBox
            // 
            this.MainTextBox.Location = new System.Drawing.Point(281, 146);
            this.MainTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.MainTextBox.Multiline = true;
            this.MainTextBox.Name = "MainTextBox";
            this.MainTextBox.ReadOnly = true;
            this.MainTextBox.Size = new System.Drawing.Size(188, 50);
            this.MainTextBox.TabIndex = 1;
            // 
            // DataTextBox
            // 
            this.DataTextBox.Location = new System.Drawing.Point(84, 27);
            this.DataTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.DataTextBox.Multiline = true;
            this.DataTextBox.Name = "DataTextBox";
            this.DataTextBox.ReadOnly = true;
            this.DataTextBox.Size = new System.Drawing.Size(268, 94);
            this.DataTextBox.TabIndex = 2;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(541, 172);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(189, 64);
            this.button2.TabIndex = 3;
            this.button2.Text = "Mostrar contenidos de celda";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.ShowCellButton_Click);
            // 
            // CleanupButton
            // 
            this.CleanupButton.Location = new System.Drawing.Point(84, 218);
            this.CleanupButton.Margin = new System.Windows.Forms.Padding(4);
            this.CleanupButton.Name = "CleanupButton";
            this.CleanupButton.Size = new System.Drawing.Size(189, 64);
            this.CleanupButton.TabIndex = 4;
            this.CleanupButton.Text = "Cerrar Excel Cargado";
            this.CleanupButton.UseVisualStyleBackColor = true;
            this.CleanupButton.Click += new System.EventHandler(this.CleanupButton_Click);
            // 
            // CellSearchTextBox
            // 
            this.CellSearchTextBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CellSearchTextBox.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CellSearchTextBox.Location = new System.Drawing.Point(541, 103);
            this.CellSearchTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.CellSearchTextBox.Name = "CellSearchTextBox";
            this.CellSearchTextBox.ReadOnly = true;
            this.CellSearchTextBox.Size = new System.Drawing.Size(184, 18);
            this.CellSearchTextBox.TabIndex = 6;
            this.CellSearchTextBox.Text = "Ubicacion de celda:";
            // 
            // cellBox1
            // 
            this.cellBox1.Location = new System.Drawing.Point(541, 129);
            this.cellBox1.Margin = new System.Windows.Forms.Padding(4);
            this.cellBox1.Name = "cellBox1";
            this.cellBox1.Size = new System.Drawing.Size(87, 22);
            this.cellBox1.TabIndex = 7;
            // 
            // cellBox2
            // 
            this.cellBox2.Location = new System.Drawing.Point(636, 129);
            this.cellBox2.Margin = new System.Windows.Forms.Padding(4);
            this.cellBox2.Name = "cellBox2";
            this.cellBox2.Size = new System.Drawing.Size(95, 22);
            this.cellBox2.TabIndex = 8;
            // 
            // RowConfirmButton
            // 
            this.RowConfirmButton.Enabled = false;
            this.RowConfirmButton.Location = new System.Drawing.Point(187, 433);
            this.RowConfirmButton.Margin = new System.Windows.Forms.Padding(4);
            this.RowConfirmButton.Name = "RowConfirmButton";
            this.RowConfirmButton.Size = new System.Drawing.Size(165, 64);
            this.RowConfirmButton.TabIndex = 12;
            this.RowConfirmButton.Text = "Confirmar filas (categoria,datos)";
            this.RowConfirmButton.UseVisualStyleBackColor = true;
            this.RowConfirmButton.Click += new System.EventHandler(this.RowConfirmButton_Click);
            // 
            // RowBox1
            // 
            this.RowBox1.Location = new System.Drawing.Point(360, 475);
            this.RowBox1.Margin = new System.Windows.Forms.Padding(4);
            this.RowBox1.Name = "RowBox1";
            this.RowBox1.Size = new System.Drawing.Size(87, 22);
            this.RowBox1.TabIndex = 13;
            this.RowBox1.Text = "9";
            // 
            // FilterStartButton
            // 
            this.FilterStartButton.Enabled = false;
            this.FilterStartButton.Location = new System.Drawing.Point(281, 345);
            this.FilterStartButton.Margin = new System.Windows.Forms.Padding(4);
            this.FilterStartButton.Name = "FilterStartButton";
            this.FilterStartButton.Size = new System.Drawing.Size(189, 64);
            this.FilterStartButton.TabIndex = 15;
            this.FilterStartButton.Text = "Empezar proceso (Filtro)";
            this.FilterStartButton.UseVisualStyleBackColor = true;
            this.FilterStartButton.Click += new System.EventHandler(this.FilterStartButton_Click);
            // 
            // RowBox2
            // 
            this.RowBox2.Location = new System.Drawing.Point(360, 446);
            this.RowBox2.Margin = new System.Windows.Forms.Padding(4);
            this.RowBox2.Name = "RowBox2";
            this.RowBox2.Size = new System.Drawing.Size(87, 22);
            this.RowBox2.TabIndex = 19;
            this.RowBox2.Text = "5";
            // 
            // ListBox
            // 
            this.ListBox.Location = new System.Drawing.Point(753, 129);
            this.ListBox.Margin = new System.Windows.Forms.Padding(4);
            this.ListBox.Multiline = true;
            this.ListBox.Name = "ListBox";
            this.ListBox.ReadOnly = true;
            this.ListBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.ListBox.Size = new System.Drawing.Size(261, 354);
            this.ListBox.TabIndex = 20;
            this.ListBox.WordWrap = false;
            // 
            // MissingBox
            // 
            this.MissingBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.MissingBox.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MissingBox.Location = new System.Drawing.Point(753, 103);
            this.MissingBox.Margin = new System.Windows.Forms.Padding(4);
            this.MissingBox.Name = "MissingBox";
            this.MissingBox.ReadOnly = true;
            this.MissingBox.Size = new System.Drawing.Size(184, 18);
            this.MissingBox.TabIndex = 21;
            this.MissingBox.Text = "Productos no encontrados:";
            // 
            // YearTextBox
            // 
            this.YearTextBox.Location = new System.Drawing.Point(588, 384);
            this.YearTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.YearTextBox.Name = "YearTextBox";
            this.YearTextBox.Size = new System.Drawing.Size(87, 22);
            this.YearTextBox.TabIndex = 22;
            // 
            // MonthTextBox
            // 
            this.MonthTextBox.Location = new System.Drawing.Point(588, 345);
            this.MonthTextBox.Margin = new System.Windows.Forms.Padding(4);
            this.MonthTextBox.Name = "MonthTextBox";
            this.MonthTextBox.Size = new System.Drawing.Size(87, 22);
            this.MonthTextBox.TabIndex = 23;
            // 
            // textBox3
            // 
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox3.Location = new System.Drawing.Point(509, 345);
            this.textBox3.Margin = new System.Windows.Forms.Padding(4);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(71, 18);
            this.textBox3.TabIndex = 24;
            this.textBox3.Text = "Mes:";
            // 
            // textBox4
            // 
            this.textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox4.Font = new System.Drawing.Font("Microsoft New Tai Lue", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox4.Location = new System.Drawing.Point(509, 384);
            this.textBox4.Margin = new System.Windows.Forms.Padding(4);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(71, 18);
            this.textBox4.TabIndex = 25;
            this.textBox4.Text = "Año:";
            // 
            // CancelButton
            // 
            this.CancelButton.Enabled = false;
            this.CancelButton.Location = new System.Drawing.Point(84, 345);
            this.CancelButton.Margin = new System.Windows.Forms.Padding(4);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(189, 64);
            this.CancelButton.TabIndex = 26;
            this.CancelButton.Text = "Cancelar";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1067, 554);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.MonthTextBox);
            this.Controls.Add(this.YearTextBox);
            this.Controls.Add(this.MissingBox);
            this.Controls.Add(this.ListBox);
            this.Controls.Add(this.RowBox2);
            this.Controls.Add(this.FilterStartButton);
            this.Controls.Add(this.RowBox1);
            this.Controls.Add(this.RowConfirmButton);
            this.Controls.Add(this.cellBox2);
            this.Controls.Add(this.cellBox1);
            this.Controls.Add(this.CellSearchTextBox);
            this.Controls.Add(this.CleanupButton);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.DataTextBox);
            this.Controls.Add(this.MainTextBox);
            this.Controls.Add(this.OpenFileButton);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Main_Form";
            this.Text = " ";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OpenFileButton;
        private System.Windows.Forms.TextBox MainTextBox;
        private System.Windows.Forms.TextBox DataTextBox;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button CleanupButton;
        private System.Windows.Forms.TextBox CellSearchTextBox;
        private System.Windows.Forms.TextBox cellBox1;
        private System.Windows.Forms.TextBox cellBox2;
        private System.Windows.Forms.Button RowConfirmButton;
        private System.Windows.Forms.TextBox RowBox1;
        private System.Windows.Forms.Button FilterStartButton;
        private System.Windows.Forms.TextBox RowBox2;
        private System.Windows.Forms.TextBox ListBox;
        private System.Windows.Forms.TextBox MissingBox;
        private System.Windows.Forms.TextBox YearTextBox;
        private System.Windows.Forms.TextBox MonthTextBox;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.Button CancelButton;
    }
}

