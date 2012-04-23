using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using stdole;

using NetOffice;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using VBIDE = NetOffice.VBIDEApi;
using Word = NetOffice.WordApi;
using Outlook = NetOffice.OutlookApi;
using PowerPoint = NetOffice.PowerPointApi;
using Access = NetOffice.AccessApi;
using DAO = NetOffice.DAOApi;
using ADODB = NetOffice.ADODBApi;
using OWC10 = NetOffice.OWC10Api;
using MSDATASRC = NetOffice.MSDATASRCApi;
using MSComctlLib = NetOffice.MSComctlLibApi;
using MSProject = NetOffice.MSProjectApi;
using MSHTML = NetOffice.MSHTMLApi;

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

            // not necessary any longer
            // LateBindingApi.Core.Factory.Initialize();

            
            /*>> your testcode here <<*/
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
            this.ResumeLayout(false);

        }

        #endregion
    }
}
