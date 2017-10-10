namespace TutorialsBase
{
    partial class AreaForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AreaForm));
            this.panelTutorials = new System.Windows.Forms.Panel();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.OnlineTabPage = new System.Windows.Forms.TabPage();
            this.webBrowserTutorialContent = new System.Windows.Forms.WebBrowser();
            this.OfflineTabPage = new System.Windows.Forms.TabPage();
            this.panelShowTutorialLink = new System.Windows.Forms.Panel();
            this.labelOffHint = new System.Windows.Forms.Label();
            this.linkLabelTutorialContent = new System.Windows.Forms.LinkLabel();
            this.SampleTabPage = new System.Windows.Forms.TabPage();
            this.panelTutorialArea = new System.Windows.Forms.Panel();
            this.buttonRunTutorial = new System.Windows.Forms.Button();
            this.AreaTabPage = new System.Windows.Forms.TabPage();
            this.panelTutorials.SuspendLayout();
            this.tabControl1.SuspendLayout();
            this.OnlineTabPage.SuspendLayout();
            this.OfflineTabPage.SuspendLayout();
            this.panelShowTutorialLink.SuspendLayout();
            this.SampleTabPage.SuspendLayout();
            this.panelTutorialArea.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTutorials
            // 
            this.panelTutorials.BackColor = System.Drawing.SystemColors.Control;
            this.panelTutorials.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelTutorials.Controls.Add(this.tabControl1);
            this.panelTutorials.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTutorials.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.panelTutorials.Location = new System.Drawing.Point(0, 0);
            this.panelTutorials.Name = "panelTutorials";
            this.panelTutorials.Size = new System.Drawing.Size(784, 562);
            this.panelTutorials.TabIndex = 4;
            this.panelTutorials.Resize += new System.EventHandler(this.panelTutorials_Resize);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.OnlineTabPage);
            this.tabControl1.Controls.Add(this.OfflineTabPage);
            this.tabControl1.Controls.Add(this.SampleTabPage);
            this.tabControl1.Controls.Add(this.AreaTabPage);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(782, 560);
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // OnlineTabPage
            // 
            this.OnlineTabPage.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.OnlineTabPage.Controls.Add(this.webBrowserTutorialContent);
            this.OnlineTabPage.Location = new System.Drawing.Point(4, 25);
            this.OnlineTabPage.Name = "OnlineTabPage";
            this.OnlineTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.OnlineTabPage.Size = new System.Drawing.Size(774, 531);
            this.OnlineTabPage.TabIndex = 2;
            this.OnlineTabPage.Text = "Introduction";
            // 
            // webBrowserTutorialContent
            // 
            this.webBrowserTutorialContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowserTutorialContent.IsWebBrowserContextMenuEnabled = false;
            this.webBrowserTutorialContent.Location = new System.Drawing.Point(3, 3);
            this.webBrowserTutorialContent.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowserTutorialContent.Name = "webBrowserTutorialContent";
            this.webBrowserTutorialContent.ScriptErrorsSuppressed = true;
            this.webBrowserTutorialContent.Size = new System.Drawing.Size(768, 525);
            this.webBrowserTutorialContent.TabIndex = 1;
            this.webBrowserTutorialContent.WebBrowserShortcutsEnabled = false;
            // 
            // OfflineTabPage
            // 
            this.OfflineTabPage.Controls.Add(this.panelShowTutorialLink);
            this.OfflineTabPage.Location = new System.Drawing.Point(4, 25);
            this.OfflineTabPage.Name = "OfflineTabPage";
            this.OfflineTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.OfflineTabPage.Size = new System.Drawing.Size(774, 531);
            this.OfflineTabPage.TabIndex = 0;
            this.OfflineTabPage.Text = "Introduction";
            this.OfflineTabPage.UseVisualStyleBackColor = true;
            // 
            // panelShowTutorialLink
            // 
            this.panelShowTutorialLink.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.panelShowTutorialLink.Controls.Add(this.labelOffHint);
            this.panelShowTutorialLink.Controls.Add(this.linkLabelTutorialContent);
            this.panelShowTutorialLink.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelShowTutorialLink.Location = new System.Drawing.Point(3, 3);
            this.panelShowTutorialLink.Name = "panelShowTutorialLink";
            this.panelShowTutorialLink.Size = new System.Drawing.Size(768, 525);
            this.panelShowTutorialLink.TabIndex = 7;
            // 
            // labelOffHint
            // 
            this.labelOffHint.AutoSize = true;
            this.labelOffHint.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.labelOffHint.Location = new System.Drawing.Point(25, 18);
            this.labelOffHint.Name = "labelOffHint";
            this.labelOffHint.Size = new System.Drawing.Size(566, 20);
            this.labelOffHint.TabIndex = 7;
            this.labelOffHint.Text = "Connect to online documentation is disabled. Enable in program options or visit:";
            // 
            // linkLabelTutorialContent
            // 
            this.linkLabelTutorialContent.AutoSize = true;
            this.linkLabelTutorialContent.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.linkLabelTutorialContent.Location = new System.Drawing.Point(25, 49);
            this.linkLabelTutorialContent.Name = "linkLabelTutorialContent";
            this.linkLabelTutorialContent.Size = new System.Drawing.Size(180, 20);
            this.linkLabelTutorialContent.TabIndex = 6;
            this.linkLabelTutorialContent.TabStop = true;
            this.linkLabelTutorialContent.Tag = "http://netoffice.codeplex.com";
            this.linkLabelTutorialContent.Text = "linkLabelTutorialContent";
            this.linkLabelTutorialContent.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabelTutorialContent_LinkClicked);
            // 
            // SampleTabPage
            // 
            this.SampleTabPage.Controls.Add(this.panelTutorialArea);
            this.SampleTabPage.Location = new System.Drawing.Point(4, 25);
            this.SampleTabPage.Name = "SampleTabPage";
            this.SampleTabPage.Padding = new System.Windows.Forms.Padding(3);
            this.SampleTabPage.Size = new System.Drawing.Size(774, 531);
            this.SampleTabPage.TabIndex = 1;
            this.SampleTabPage.Text = "Run Code Sample";
            this.SampleTabPage.UseVisualStyleBackColor = true;
            // 
            // panelTutorialArea
            // 
            this.panelTutorialArea.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.panelTutorialArea.Controls.Add(this.buttonRunTutorial);
            this.panelTutorialArea.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTutorialArea.Location = new System.Drawing.Point(3, 3);
            this.panelTutorialArea.Name = "panelTutorialArea";
            this.panelTutorialArea.Size = new System.Drawing.Size(768, 525);
            this.panelTutorialArea.TabIndex = 0;
            // 
            // buttonRunTutorial
            // 
            this.buttonRunTutorial.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonRunTutorial.Font = new System.Drawing.Font("MS Reference Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonRunTutorial.Image = ((System.Drawing.Image)(resources.GetObject("buttonRunTutorial.Image")));
            this.buttonRunTutorial.Location = new System.Drawing.Point(237, 112);
            this.buttonRunTutorial.Name = "buttonRunTutorial";
            this.buttonRunTutorial.Size = new System.Drawing.Size(274, 285);
            this.buttonRunTutorial.TabIndex = 0;
            this.buttonRunTutorial.Text = "Click here to run tutorial";
            this.buttonRunTutorial.UseVisualStyleBackColor = true;
            this.buttonRunTutorial.Click += new System.EventHandler(this.buttonRunTutorial_Click);
            // 
            // AreaTabPage
            // 
            this.AreaTabPage.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.AreaTabPage.Location = new System.Drawing.Point(4, 25);
            this.AreaTabPage.Name = "AreaTabPage";
            this.AreaTabPage.Size = new System.Drawing.Size(774, 531);
            this.AreaTabPage.TabIndex = 3;
            this.AreaTabPage.Text = "Run Code Sample";
            // 
            // AreaForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(201)))), ((int)(((byte)(227)))), ((int)(((byte)(243)))));
            this.ClientSize = new System.Drawing.Size(784, 562);
            this.Controls.Add(this.panelTutorials);
            this.MinimumSize = new System.Drawing.Size(800, 600);
            this.Name = "AreaForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "AreaForm";
            this.panelTutorials.ResumeLayout(false);
            this.tabControl1.ResumeLayout(false);
            this.OnlineTabPage.ResumeLayout(false);
            this.OfflineTabPage.ResumeLayout(false);
            this.panelShowTutorialLink.ResumeLayout(false);
            this.panelShowTutorialLink.PerformLayout();
            this.SampleTabPage.ResumeLayout(false);
            this.panelTutorialArea.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelTutorials;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage OfflineTabPage;
        private System.Windows.Forms.Panel panelShowTutorialLink;
        private System.Windows.Forms.Label labelOffHint;
        private System.Windows.Forms.LinkLabel linkLabelTutorialContent;
        private System.Windows.Forms.TabPage SampleTabPage;
        private System.Windows.Forms.Panel panelTutorialArea;
        private System.Windows.Forms.Button buttonRunTutorial;
        private System.Windows.Forms.TabPage OnlineTabPage;
        private System.Windows.Forms.WebBrowser webBrowserTutorialContent;
        private System.Windows.Forms.TabPage AreaTabPage;
    }
}