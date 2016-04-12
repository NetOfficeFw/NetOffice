namespace NetOffice.DeveloperToolbox.Controls.Text
{
	partial class Scroller
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

		#region Component Designer generated code

		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			this.m_timer = new System.Windows.Forms.Timer(this.components);
			this.SuspendLayout();
			// 
			// m_timer
			// 
			this.m_timer.Interval = 50;
			this.m_timer.Tick += new System.EventHandler(this.OnTimerTick);
			// 
			// Scroller
			// 
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
			this.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
			this.Name = "Scroller";
			this.Size = new System.Drawing.Size(447, 429);
			this.Paint += new System.Windows.Forms.PaintEventHandler(this.OnPaint);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Timer m_timer;
	}
}
