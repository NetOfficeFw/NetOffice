namespace NOTools.ConsoleMonitor
{
    partial class ChannelViewControl
    {
        /// <summary> 
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">True, wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls False.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary> 
        /// Erforderliche Methode für die Designerunterstützung. 
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ChannelViewControl));
            this.ListViewChannels = new System.Windows.Forms.ListView();
            this.colChannel = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colTime = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colMachine = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colAppDomain = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colMessage = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageListColumns = new System.Windows.Forms.ImageList(this.components);
            this.SuspendLayout();
            // 
            // ListViewChannels
            // 
            this.ListViewChannels.AutoArrange = false;
            this.ListViewChannels.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ListViewChannels.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colChannel,
            this.colTime,
            this.colMachine,
            this.colAppDomain,
            this.colMessage});
            this.ListViewChannels.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ListViewChannels.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ListViewChannels.FullRowSelect = true;
            this.ListViewChannels.GridLines = true;
            this.ListViewChannels.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.ListViewChannels.LabelWrap = false;
            this.ListViewChannels.Location = new System.Drawing.Point(0, 0);
            this.ListViewChannels.MultiSelect = false;
            this.ListViewChannels.Name = "ListViewChannels";
            this.ListViewChannels.ShowGroups = false;
            this.ListViewChannels.Size = new System.Drawing.Size(640, 480);
            this.ListViewChannels.SmallImageList = this.imageListColumns;
            this.ListViewChannels.TabIndex = 1;
            this.ListViewChannels.UseCompatibleStateImageBehavior = false;
            this.ListViewChannels.View = System.Windows.Forms.View.Details;
            // 
            // colChannel
            // 
            this.colChannel.Text = "Channel";
            this.colChannel.Width = 110;
            // 
            // colTime
            // 
            this.colTime.Text = "Last Update";
            this.colTime.Width = 110;
            // 
            // colMachine
            // 
            this.colMachine.Text = "Machine";
            this.colMachine.Width = 110;
            // 
            // colAppDomain
            // 
            this.colAppDomain.Text = "AppDomain";
            this.colAppDomain.Width = 110;
            // 
            // colMessage
            // 
            this.colMessage.Text = "Last Message";
            this.colMessage.Width = 150;
            // 
            // imageListColumns
            // 
            this.imageListColumns.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListColumns.ImageStream")));
            this.imageListColumns.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListColumns.Images.SetKeyName(0, "channels.png");
            this.imageListColumns.Images.SetKeyName(1, "clock.png");
            this.imageListColumns.Images.SetKeyName(2, "machine.png");
            this.imageListColumns.Images.SetKeyName(3, "appdomain.png");
            this.imageListColumns.Images.SetKeyName(4, "value.png");
            // 
            // ChannelViewControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ListViewChannels);
            this.Name = "ChannelViewControl";
            this.Size = new System.Drawing.Size(640, 480);
            this.Resize += new System.EventHandler(this.ChannelViewControl_Resize);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView ListViewChannels;
        private System.Windows.Forms.ColumnHeader colChannel;
        private System.Windows.Forms.ColumnHeader colTime;
        private System.Windows.Forms.ColumnHeader colMessage;
        private System.Windows.Forms.ColumnHeader colMachine;
        private System.Windows.Forms.ColumnHeader colAppDomain;
        private System.Windows.Forms.ImageList imageListColumns;
    }
}
