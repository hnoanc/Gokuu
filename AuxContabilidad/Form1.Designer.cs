namespace AuxContabilidad
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.pbProgreso = new System.Windows.Forms.ProgressBar();
            this.lblTitulo = new System.Windows.Forms.Label();
            this.btnFindExcel = new System.Windows.Forms.Button();
            this.txtRutaExcel = new System.Windows.Forms.TextBox();
            this.txtRutaFacturas = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnFindFolder = new System.Windows.Forms.Button();
            this.fbdFindFacturas = new System.Windows.Forms.FolderBrowserDialog();
            this.ofdFindExcel = new System.Windows.Forms.OpenFileDialog();
            this.btnProceder = new System.Windows.Forms.Button();
            this.lblPorcentaje = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnConectar = new System.Windows.Forms.Button();
            this.txtPass = new System.Windows.Forms.TextBox();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.cbxNombre = new System.Windows.Forms.ComboBox();
            this.txtIP = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // pbProgreso
            // 
            this.pbProgreso.Location = new System.Drawing.Point(132, 257);
            this.pbProgreso.Name = "pbProgreso";
            this.pbProgreso.Size = new System.Drawing.Size(422, 25);
            this.pbProgreso.TabIndex = 5;
            // 
            // lblTitulo
            // 
            this.lblTitulo.AutoSize = true;
            this.lblTitulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitulo.Location = new System.Drawing.Point(128, 20);
            this.lblTitulo.Name = "lblTitulo";
            this.lblTitulo.Size = new System.Drawing.Size(178, 20);
            this.lblTitulo.TabIndex = 4;
            this.lblTitulo.Text = "Seleccione el archivo";
            // 
            // btnFindExcel
            // 
            this.btnFindExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFindExcel.Location = new System.Drawing.Point(520, 54);
            this.btnFindExcel.Name = "btnFindExcel";
            this.btnFindExcel.Size = new System.Drawing.Size(34, 28);
            this.btnFindExcel.TabIndex = 3;
            this.btnFindExcel.Text = "...";
            this.btnFindExcel.UseVisualStyleBackColor = true;
            this.btnFindExcel.Click += new System.EventHandler(this.btnFindExcel_Click);
            // 
            // txtRutaExcel
            // 
            this.txtRutaExcel.Location = new System.Drawing.Point(132, 54);
            this.txtRutaExcel.Margin = new System.Windows.Forms.Padding(2);
            this.txtRutaExcel.Multiline = true;
            this.txtRutaExcel.Name = "txtRutaExcel";
            this.txtRutaExcel.ReadOnly = true;
            this.txtRutaExcel.Size = new System.Drawing.Size(383, 28);
            this.txtRutaExcel.TabIndex = 6;
            // 
            // txtRutaFacturas
            // 
            this.txtRutaFacturas.Location = new System.Drawing.Point(132, 133);
            this.txtRutaFacturas.Margin = new System.Windows.Forms.Padding(2);
            this.txtRutaFacturas.Multiline = true;
            this.txtRutaFacturas.Name = "txtRutaFacturas";
            this.txtRutaFacturas.ReadOnly = true;
            this.txtRutaFacturas.Size = new System.Drawing.Size(383, 28);
            this.txtRutaFacturas.TabIndex = 9;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(128, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(246, 20);
            this.label1.TabIndex = 8;
            this.label1.Text = "Seleccione la carpeta destino";
            // 
            // btnFindFolder
            // 
            this.btnFindFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFindFolder.Location = new System.Drawing.Point(520, 133);
            this.btnFindFolder.Name = "btnFindFolder";
            this.btnFindFolder.Size = new System.Drawing.Size(34, 28);
            this.btnFindFolder.TabIndex = 7;
            this.btnFindFolder.Text = "...";
            this.btnFindFolder.UseVisualStyleBackColor = true;
            this.btnFindFolder.Click += new System.EventHandler(this.btnFindFolder_Click);
            // 
            // ofdFindExcel
            // 
            this.ofdFindExcel.FileName = "openFileDialog1";
            // 
            // btnProceder
            // 
            this.btnProceder.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(255)))), ((int)(((byte)(128)))));
            this.btnProceder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProceder.Location = new System.Drawing.Point(253, 179);
            this.btnProceder.Name = "btnProceder";
            this.btnProceder.Size = new System.Drawing.Size(175, 41);
            this.btnProceder.TabIndex = 10;
            this.btnProceder.Text = "Ayudame Goku!!!";
            this.btnProceder.UseVisualStyleBackColor = false;
            this.btnProceder.Click += new System.EventHandler(this.btnProceder_Click);
            // 
            // lblPorcentaje
            // 
            this.lblPorcentaje.AutoSize = true;
            this.lblPorcentaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPorcentaje.Location = new System.Drawing.Point(330, 234);
            this.lblPorcentaje.Name = "lblPorcentaje";
            this.lblPorcentaje.Size = new System.Drawing.Size(24, 20);
            this.lblPorcentaje.TabIndex = 11;
            this.lblPorcentaje.Text = "%";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.btnConectar);
            this.groupBox1.Controls.Add(this.txtPass);
            this.groupBox1.Controls.Add(this.txtUser);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.cbxNombre);
            this.groupBox1.Controls.Add(this.txtIP);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(676, 105);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // btnConectar
            // 
            this.btnConectar.Location = new System.Drawing.Point(526, 78);
            this.btnConectar.Name = "btnConectar";
            this.btnConectar.Size = new System.Drawing.Size(136, 23);
            this.btnConectar.TabIndex = 24;
            this.btnConectar.Text = "Conectar";
            this.btnConectar.UseVisualStyleBackColor = true;
            this.btnConectar.Click += new System.EventHandler(this.btnConectar_Click);
            // 
            // txtPass
            // 
            this.txtPass.Location = new System.Drawing.Point(410, 52);
            this.txtPass.Name = "txtPass";
            this.txtPass.PasswordChar = '*';
            this.txtPass.Size = new System.Drawing.Size(252, 20);
            this.txtPass.TabIndex = 23;
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(410, 26);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(252, 20);
            this.txtUser.TabIndex = 22;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(348, 52);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(56, 13);
            this.label4.TabIndex = 21;
            this.label4.Text = "Password:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(348, 29);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "User:";
            // 
            // cbxNombre
            // 
            this.cbxNombre.FormattingEnabled = true;
            this.cbxNombre.Location = new System.Drawing.Point(98, 22);
            this.cbxNombre.Name = "cbxNombre";
            this.cbxNombre.Size = new System.Drawing.Size(215, 21);
            this.cbxNombre.TabIndex = 16;
            this.cbxNombre.SelectedIndexChanged += new System.EventHandler(this.cbxNombre_SelectedIndexChanged);
            // 
            // txtIP
            // 
            this.txtIP.Location = new System.Drawing.Point(98, 49);
            this.txtIP.Name = "txtIP";
            this.txtIP.Size = new System.Drawing.Size(215, 20);
            this.txtIP.TabIndex = 19;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(20, 13);
            this.label2.TabIndex = 18;
            this.label2.Text = "IP:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(47, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Nombre:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lblTitulo);
            this.groupBox2.Controls.Add(this.btnFindExcel);
            this.groupBox2.Controls.Add(this.pbProgreso);
            this.groupBox2.Controls.Add(this.lblPorcentaje);
            this.groupBox2.Controls.Add(this.txtRutaExcel);
            this.groupBox2.Controls.Add(this.btnProceder);
            this.groupBox2.Controls.Add(this.btnFindFolder);
            this.groupBox2.Controls.Add(this.txtRutaFacturas);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(12, 113);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(676, 307);
            this.groupBox2.TabIndex = 14;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "groupBox2";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(248, 75);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(65, 23);
            this.button1.TabIndex = 25;
            this.button1.Text = "Olvidar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 428);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Aux. Contabilidad";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ProgressBar pbProgreso;
        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Button btnFindExcel;
        private System.Windows.Forms.TextBox txtRutaExcel;
        private System.Windows.Forms.TextBox txtRutaFacturas;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnFindFolder;
        private System.Windows.Forms.FolderBrowserDialog fbdFindFacturas;
        private System.Windows.Forms.OpenFileDialog ofdFindExcel;
        private System.Windows.Forms.Button btnProceder;
        private System.Windows.Forms.Label lblPorcentaje;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cbxNombre;
        private System.Windows.Forms.TextBox txtIP;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnConectar;
        private System.Windows.Forms.TextBox txtPass;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button button1;
    }
}

