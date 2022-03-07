namespace WindowsFormsApp2
{
    partial class Form2
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
            this.Generar = new System.Windows.Forms.Button();
            this.Origen = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.ConAnual = new System.Windows.Forms.CheckBox();
            this.Impuesto = new System.Windows.Forms.ComboBox();
            this.FechaOpcion = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.DiferenciarMails = new System.Windows.Forms.CheckBox();
            this.txtCantidadCorte = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.ConCabecera = new System.Windows.Forms.CheckBox();
            this.txtArchivoOrigen = new System.Windows.Forms.TextBox();
            this.txtCantidad = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.barraLeidos = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.barraGenerados = new System.Windows.Forms.ProgressBar();
            this.txturl = new System.Windows.Forms.TextBox();
            this.lburl = new System.Windows.Forms.Label();
            this.BarraReloj = new System.Windows.Forms.ProgressBar();
            this.LabelZip = new System.Windows.Forms.Label();
            this.ModoOriginal = new System.Windows.Forms.RadioButton();
            this.ModoNuevo = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // Generar
            // 
            this.Generar.Location = new System.Drawing.Point(626, 228);
            this.Generar.Name = "Generar";
            this.Generar.Size = new System.Drawing.Size(75, 23);
            this.Generar.TabIndex = 0;
            this.Generar.Text = "&Generar";
            this.Generar.UseVisualStyleBackColor = true;
            this.Generar.Click += new System.EventHandler(this.button1_Click);
            // 
            // Origen
            // 
            this.Origen.InitialDirectory = "\\\\arba.gov.ar\\DE\\GGTI\\Gerencia de Produccion\\Mantenimiento\\Boleta Electronica\\Ori" +
    "gen";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(28, 172);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(65, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "Origen";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(25, 301);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(86, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Mails Generados";
            // 
            // ConAnual
            // 
            this.ConAnual.AutoSize = true;
            this.ConAnual.Location = new System.Drawing.Point(345, 22);
            this.ConAnual.Name = "ConAnual";
            this.ConAnual.Size = new System.Drawing.Size(75, 17);
            this.ConAnual.TabIndex = 9;
            this.ConAnual.Text = "Con Anual";
            this.ConAnual.UseVisualStyleBackColor = true;
            // 
            // Impuesto
            // 
            this.Impuesto.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.Impuesto.FormattingEnabled = true;
            this.Impuesto.Items.AddRange(new object[] {
            "Impuesto Automotor",
            "Impuesto a las Embarcaciones",
            "Impuesto Urbano Edificado",
            "Impuesto Urbano Baldío",
            "Impuesto Rural",
            "Impuesto Complementario"});
            this.Impuesto.Location = new System.Drawing.Point(81, 22);
            this.Impuesto.Name = "Impuesto";
            this.Impuesto.Size = new System.Drawing.Size(223, 21);
            this.Impuesto.TabIndex = 10;
            this.Impuesto.SelectedIndexChanged += new System.EventHandler(this.Impuesto_SelectedIndexChanged);
            // 
            // FechaOpcion
            // 
            this.FechaOpcion.Location = new System.Drawing.Point(110, 52);
            this.FechaOpcion.Name = "FechaOpcion";
            this.FechaOpcion.Size = new System.Drawing.Size(194, 20);
            this.FechaOpcion.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(25, 25);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(50, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Impuesto";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(25, 55);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(74, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Fecha Opción";
            // 
            // DiferenciarMails
            // 
            this.DiferenciarMails.AutoSize = true;
            this.DiferenciarMails.Location = new System.Drawing.Point(597, 22);
            this.DiferenciarMails.Name = "DiferenciarMails";
            this.DiferenciarMails.Size = new System.Drawing.Size(104, 17);
            this.DiferenciarMails.TabIndex = 20;
            this.DiferenciarMails.Text = "Diferenciar Mails";
            this.DiferenciarMails.UseVisualStyleBackColor = true;
            // 
            // txtCantidadCorte
            // 
            this.txtCantidadCorte.Location = new System.Drawing.Point(639, 52);
            this.txtCantidadCorte.Name = "txtCantidadCorte";
            this.txtCantidadCorte.Size = new System.Drawing.Size(62, 20);
            this.txtCantidadCorte.TabIndex = 22;
            this.txtCantidadCorte.Text = "150000";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(568, 55);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(65, 13);
            this.label8.TabIndex = 21;
            this.label8.Text = "Cortar cada ";
            // 
            // ConCabecera
            // 
            this.ConCabecera.AutoSize = true;
            this.ConCabecera.Checked = true;
            this.ConCabecera.CheckState = System.Windows.Forms.CheckState.Checked;
            this.ConCabecera.Location = new System.Drawing.Point(463, 22);
            this.ConCabecera.Name = "ConCabecera";
            this.ConCabecera.Size = new System.Drawing.Size(94, 17);
            this.ConCabecera.TabIndex = 23;
            this.ConCabecera.Text = "Con Cabecera";
            this.ConCabecera.UseVisualStyleBackColor = true;
            // 
            // txtArchivoOrigen
            // 
            this.txtArchivoOrigen.Location = new System.Drawing.Point(110, 174);
            this.txtArchivoOrigen.Multiline = true;
            this.txtArchivoOrigen.Name = "txtArchivoOrigen";
            this.txtArchivoOrigen.ReadOnly = true;
            this.txtArchivoOrigen.Size = new System.Drawing.Size(591, 48);
            this.txtArchivoOrigen.TabIndex = 25;
            // 
            // txtCantidad
            // 
            this.txtCantidad.Location = new System.Drawing.Point(424, 52);
            this.txtCantidad.Name = "txtCantidad";
            this.txtCantidad.Size = new System.Drawing.Size(117, 20);
            this.txtCantidad.TabIndex = 27;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(317, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 26;
            this.label1.Text = "Cant. Suscripciones";
            // 
            // barraLeidos
            // 
            this.barraLeidos.Location = new System.Drawing.Point(129, 259);
            this.barraLeidos.Maximum = 0;
            this.barraLeidos.Name = "barraLeidos";
            this.barraLeidos.Size = new System.Drawing.Size(572, 23);
            this.barraLeidos.Step = 1;
            this.barraLeidos.TabIndex = 28;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 259);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 13);
            this.label2.TabIndex = 29;
            this.label2.Text = "Suscripciones leidas";
            // 
            // barraGenerados
            // 
            this.barraGenerados.Location = new System.Drawing.Point(129, 301);
            this.barraGenerados.Maximum = 0;
            this.barraGenerados.Name = "barraGenerados";
            this.barraGenerados.Size = new System.Drawing.Size(572, 23);
            this.barraGenerados.TabIndex = 30;
            // 
            // txturl
            // 
            this.txturl.Location = new System.Drawing.Point(81, 134);
            this.txturl.Name = "txturl";
            this.txturl.Size = new System.Drawing.Size(620, 20);
            this.txturl.TabIndex = 32;
            // 
            // lburl
            // 
            this.lburl.AutoSize = true;
            this.lburl.Location = new System.Drawing.Point(25, 137);
            this.lburl.Name = "lburl";
            this.lburl.Size = new System.Drawing.Size(29, 13);
            this.lburl.TabIndex = 31;
            this.lburl.Text = "URL";
            // 
            // BarraReloj
            // 
            this.BarraReloj.Location = new System.Drawing.Point(240, 341);
            this.BarraReloj.Name = "BarraReloj";
            this.BarraReloj.Size = new System.Drawing.Size(288, 23);
            this.BarraReloj.TabIndex = 33;
            // 
            // LabelZip
            // 
            this.LabelZip.AutoSize = true;
            this.LabelZip.Location = new System.Drawing.Point(333, 367);
            this.LabelZip.Name = "LabelZip";
            this.LabelZip.Size = new System.Drawing.Size(87, 13);
            this.LabelZip.TabIndex = 34;
            this.LabelZip.Text = "Procesando .ZIP";
            this.LabelZip.UseMnemonic = false;
            // 
            // ModoOriginal
            // 
            this.ModoOriginal.AutoSize = true;
            this.ModoOriginal.Location = new System.Drawing.Point(286, 99);
            this.ModoOriginal.Name = "ModoOriginal";
            this.ModoOriginal.Size = new System.Drawing.Size(100, 17);
            this.ModoOriginal.TabIndex = 35;
            this.ModoOriginal.TabStop = true;
            this.ModoOriginal.Text = "a Carpeta Local";
            this.ModoOriginal.UseVisualStyleBackColor = true;
            // 
            // ModoNuevo
            // 
            this.ModoNuevo.AutoSize = true;
            this.ModoNuevo.Location = new System.Drawing.Point(400, 99);
            this.ModoNuevo.Name = "ModoNuevo";
            this.ModoNuevo.Size = new System.Drawing.Size(54, 17);
            this.ModoNuevo.TabIndex = 36;
            this.ModoNuevo.TabStop = true;
            this.ModoNuevo.Text = "a Red";
            this.ModoNuevo.UseVisualStyleBackColor = true;
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(738, 396);
            this.Controls.Add(this.ModoNuevo);
            this.Controls.Add(this.ModoOriginal);
            this.Controls.Add(this.LabelZip);
            this.Controls.Add(this.BarraReloj);
            this.Controls.Add(this.txturl);
            this.Controls.Add(this.lburl);
            this.Controls.Add(this.barraGenerados);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.barraLeidos);
            this.Controls.Add(this.txtCantidad);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtArchivoOrigen);
            this.Controls.Add(this.ConCabecera);
            this.Controls.Add(this.txtCantidadCorte);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.DiferenciarMails);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.FechaOpcion);
            this.Controls.Add(this.Impuesto);
            this.Controls.Add(this.ConAnual);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.Generar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "Form2";
            this.Text = "Generar bases para campaña por mail";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Generar;
        private System.Windows.Forms.OpenFileDialog Origen;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox ConAnual;
        private System.Windows.Forms.ComboBox Impuesto;
        private System.Windows.Forms.DateTimePicker FechaOpcion;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.CheckBox DiferenciarMails;
        private System.Windows.Forms.TextBox txtCantidadCorte;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckBox ConCabecera;
        private System.Windows.Forms.TextBox txtArchivoOrigen;
        private System.Windows.Forms.TextBox txtCantidad;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar barraLeidos;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ProgressBar barraGenerados;
        private System.Windows.Forms.TextBox txturl;
        private System.Windows.Forms.Label lburl;
        private System.Windows.Forms.ProgressBar BarraReloj;
        private System.Windows.Forms.Label LabelZip;
        private System.Windows.Forms.RadioButton ModoOriginal;
        private System.Windows.Forms.RadioButton ModoNuevo;
    }
}

