using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using stdole;

using NetOffice;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;
using VBIDE = NetOffice.VBIDEApi;
using NOTools = NetOffice.OfficeApi.Tools;

namespace ClientApplication
{
    public class Form1 : System.Windows.Forms.Form
    { 
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            Excel.Application app = null;
            try
            {
                app = new Excel.Application();
                NOTools.CommonUtils utils = new NOTools.CommonUtils(app, typeof(Form1).Assembly);
                utils.Dialog.SuppressOnAutomation = false;
                utils.Dialog.SuppressOnHide = false;
                utils.Dialog.ShowDiagnostics(true);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
            finally
            {
                app.Quit();
                app.Dispose();
                Close();
            }
        }

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
            this.SuspendLayout();
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(292, 273);
            this.Name = "Form1";
            this.Text = "ClientApplication";
            this.Shown += new System.EventHandler(this.Form1_Shown);
            this.ResumeLayout(false);

        }

        #endregion
    }
}
