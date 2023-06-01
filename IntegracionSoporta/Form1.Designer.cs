namespace IntegracionSoporta
{
    partial class Form1
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
        private void InitializeComponent()
        {
            this.lbLista = new System.Windows.Forms.ListBox();
            this.btProcesar = new System.Windows.Forms.Button();
            this.btCerrar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbLista
            // 
            this.lbLista.FormattingEnabled = true;
            this.lbLista.ItemHeight = 15;
            this.lbLista.Location = new System.Drawing.Point(12, 15);
            this.lbLista.Name = "lbLista";
            this.lbLista.Size = new System.Drawing.Size(776, 379);
            this.lbLista.TabIndex = 0;
            // 
            // btProcesar
            // 
            this.btProcesar.Location = new System.Drawing.Point(322, 420);
            this.btProcesar.Name = "btProcesar";
            this.btProcesar.Size = new System.Drawing.Size(75, 23);
            this.btProcesar.TabIndex = 1;
            this.btProcesar.Text = "Iniciar";
            this.btProcesar.UseVisualStyleBackColor = true;
            this.btProcesar.Click += new System.EventHandler(this.btProcesar_Click);
            // 
            // btCerrar
            // 
            this.btCerrar.Location = new System.Drawing.Point(422, 419);
            this.btCerrar.Name = "btCerrar";
            this.btCerrar.Size = new System.Drawing.Size(75, 23);
            this.btCerrar.TabIndex = 2;
            this.btCerrar.Text = "Cerrar";
            this.btCerrar.UseVisualStyleBackColor = true;
            this.btCerrar.Click += new System.EventHandler(this.btCerrar_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btCerrar);
            this.Controls.Add(this.btProcesar);
            this.Controls.Add(this.lbLista);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Integración con Soporta";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private ListBox lbLista;
        private Button btProcesar;
        private Button btCerrar;
    }
}