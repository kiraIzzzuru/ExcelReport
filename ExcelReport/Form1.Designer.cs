
namespace ExcelReport
{
	partial class FrmReport
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
			this.btnLoadExcel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnLoadExcel
			// 
			this.btnLoadExcel.Location = new System.Drawing.Point(223, 209);
			this.btnLoadExcel.Name = "btnLoadExcel";
			this.btnLoadExcel.Size = new System.Drawing.Size(246, 31);
			this.btnLoadExcel.TabIndex = 0;
			this.btnLoadExcel.Text = "Запустить";
			this.btnLoadExcel.UseVisualStyleBackColor = true;
			this.btnLoadExcel.Click += new System.EventHandler(this.btnLoadExcel_Click);
			// 
			// FrmReport
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(674, 292);
			this.Controls.Add(this.btnLoadExcel);
			this.Name = "FrmReport";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Сбор отчета";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button btnLoadExcel;
	}
}

