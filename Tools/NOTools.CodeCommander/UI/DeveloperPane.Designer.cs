namespace NOTools.CodeCommander.UI
{
    partial class DeveloperPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeveloperPane));
            this.tabControlMain = new System.Windows.Forms.TabControl();
            this.tabPageCommands = new System.Windows.Forms.TabPage();
            this.tabPageProperties = new System.Windows.Forms.TabPage();
            this.tabPageInfo = new System.Windows.Forms.TabPage();
            this.panelSettings = new System.Windows.Forms.Panel();
            this.imageListMain = new System.Windows.Forms.ImageList(this.components);
            this.DirtyLittleTimer = new System.Windows.Forms.Timer(this.components);
            this.propertyPane1 = new NOTools.CodeCommander.UI.PropertyPane();
            this.infoPane1 = new NOTools.CodeCommander.UI.InfoPane();
            this.commandPane1 = new NOTools.CodeCommander.UI.CommandPane();
            this.tabControlMain.SuspendLayout();
            this.tabPageCommands.SuspendLayout();
            this.tabPageProperties.SuspendLayout();
            this.tabPageInfo.SuspendLayout();
            this.panelSettings.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControlMain
            // 
            this.tabControlMain.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControlMain.Controls.Add(this.tabPageCommands);
            this.tabControlMain.Controls.Add(this.tabPageProperties);
            this.tabControlMain.Controls.Add(this.tabPageInfo);
            this.tabControlMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlMain.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControlMain.HotTrack = true;
            this.tabControlMain.ImageList = this.imageListMain;
            this.tabControlMain.Location = new System.Drawing.Point(0, 0);
            this.tabControlMain.Name = "tabControlMain";
            this.tabControlMain.SelectedIndex = 0;
            this.tabControlMain.Size = new System.Drawing.Size(320, 420);
            this.tabControlMain.TabIndex = 6;
            // 
            // tabPageCommands
            // 
            this.tabPageCommands.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPageCommands.Controls.Add(this.commandPane1);
            this.tabPageCommands.ImageIndex = 1;
            this.tabPageCommands.Location = new System.Drawing.Point(4, 28);
            this.tabPageCommands.Name = "tabPageCommands";
            this.tabPageCommands.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageCommands.Size = new System.Drawing.Size(312, 388);
            this.tabPageCommands.TabIndex = 1;
            this.tabPageCommands.Text = "Commands";
            // 
            // tabPageProperties
            // 
            this.tabPageProperties.BackColor = System.Drawing.Color.LightSteelBlue;
            this.tabPageProperties.Controls.Add(this.propertyPane1);
            this.tabPageProperties.ImageIndex = 0;
            this.tabPageProperties.Location = new System.Drawing.Point(4, 28);
            this.tabPageProperties.Name = "tabPageProperties";
            this.tabPageProperties.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageProperties.Size = new System.Drawing.Size(312, 388);
            this.tabPageProperties.TabIndex = 0;
            this.tabPageProperties.Text = "Properties";
            this.tabPageProperties.UseVisualStyleBackColor = true;
            // 
            // tabPageInfo
            // 
            this.tabPageInfo.Controls.Add(this.panelSettings);
            this.tabPageInfo.ImageIndex = 2;
            this.tabPageInfo.Location = new System.Drawing.Point(4, 28);
            this.tabPageInfo.Name = "tabPageInfo";
            this.tabPageInfo.Size = new System.Drawing.Size(312, 388);
            this.tabPageInfo.TabIndex = 2;
            this.tabPageInfo.Text = "Info";
            this.tabPageInfo.UseVisualStyleBackColor = true;
            // 
            // panelSettings
            // 
            this.panelSettings.BackColor = System.Drawing.Color.LightSteelBlue;
            this.panelSettings.Controls.Add(this.infoPane1);
            this.panelSettings.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelSettings.Location = new System.Drawing.Point(0, 0);
            this.panelSettings.Name = "panelSettings";
            this.panelSettings.Size = new System.Drawing.Size(312, 388);
            this.panelSettings.TabIndex = 4;
            // 
            // imageListMain
            // 
            this.imageListMain.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListMain.ImageStream")));
            this.imageListMain.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListMain.Images.SetKeyName(0, "Tab1.png");
            this.imageListMain.Images.SetKeyName(1, "Tab2.png");
            this.imageListMain.Images.SetKeyName(2, "Tab3.png");
            // 
            // DirtyLittleTimer
            // 
            this.DirtyLittleTimer.Enabled = true;
            this.DirtyLittleTimer.Interval = 55;
            this.DirtyLittleTimer.Tick += new System.EventHandler(this.DirtyLittleTimer_Tick);
            // 
            // propertyPane1
            // 
            this.propertyPane1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.propertyPane1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.propertyPane1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.propertyPane1.Location = new System.Drawing.Point(3, 3);
            this.propertyPane1.Margin = new System.Windows.Forms.Padding(4);
            this.propertyPane1.Name = "propertyPane1";
            this.propertyPane1.Size = new System.Drawing.Size(306, 382);
            this.propertyPane1.TabIndex = 0;
            // 
            // infoPane1
            // 
            this.infoPane1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.infoPane1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.infoPane1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.infoPane1.Location = new System.Drawing.Point(0, 0);
            this.infoPane1.Margin = new System.Windows.Forms.Padding(4);
            this.infoPane1.Name = "infoPane1";
            this.infoPane1.Size = new System.Drawing.Size(312, 388);
            this.infoPane1.TabIndex = 0;
            // 
            // commandPane1
            // 
            this.commandPane1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.commandPane1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.commandPane1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.commandPane1.Location = new System.Drawing.Point(3, 3);
            this.commandPane1.Margin = new System.Windows.Forms.Padding(4);
            this.commandPane1.Name = "commandPane1";
            this.commandPane1.Size = new System.Drawing.Size(306, 382);
            this.commandPane1.TabIndex = 0;
            // 
            // DeveloperPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.Controls.Add(this.tabControlMain);
            this.DoubleBuffered = true;
            this.Name = "DeveloperPane";
            this.Size = new System.Drawing.Size(320, 420);
            this.tabControlMain.ResumeLayout(false);
            this.tabPageCommands.ResumeLayout(false);
            this.tabPageProperties.ResumeLayout(false);
            this.tabPageInfo.ResumeLayout(false);
            this.panelSettings.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControlMain;
        private System.Windows.Forms.TabPage tabPageProperties;
        private System.Windows.Forms.TabPage tabPageCommands;
        private System.Windows.Forms.TabPage tabPageInfo;
        private System.Windows.Forms.ImageList imageListMain;
        private System.Windows.Forms.Timer DirtyLittleTimer;
        private System.Windows.Forms.Panel panelSettings;
        private InfoPane infoPane1;
        private PropertyPane propertyPane1;
        private CommandPane commandPane1;
    }
}
