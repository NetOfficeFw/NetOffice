using System;
using System.Reflection;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using NetOffice;
using NetOffice.Attributes;
using Excel = NetOffice.ExcelApi;
using Office = NetOffice.OfficeApi;
//using Point = NetOffice.PowerPointApi;
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

            try
            {
                //new RunExcel01().Run();
                new RunWord01().Run();
                ////new RunExcel02().Run();
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void PerformanceTrace_Alert(PerformanceTrace sender, PerformanceTrace.PerformanceAlertEventArgs args)
        {
            Console.WriteLine(args);
        }

        private void Form1_Shown(object sender, EventArgs e)
        {          
            try
            {
                //new MultiRegisterClient().Test();
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.ToString());
            }
            finally
            {
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
