using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.WordApi.Tools;
using LinqToWikipedia;

namespace Sample.Addin
{ 
    /// <summary>
    /// Custom pane for Word. The control implements the ITaskPane interface from NetOffice.WordApi.Tools
    /// Learn more about the NetOffice Tools namespace: http://netoffice.codeplex.com/wikipage?title=Tools_EN
    /// </summary>
    public partial class WikipediaPane : UserControl, ITaskPane 
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public WikipediaPane()
        {
            try
            {
                InitializeComponent();
                Datacontext = new WikipediaContext();
                ClearResult();
            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }
        }
    
        WikipediaContext Datacontext { get; set; }

        #region UI Trigger

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            try
            {
                ClearError();
                var opensearch = (from wikipedia in Datacontext.OpenSearch where wikipedia.Keyword == textBoxKeyWords.Text select wikipedia).Take(100);
                gridResults.DataSource = opensearch.ToList();

            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }
        }

        private void gridResults_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gridResults.Rows.Count > e.RowIndex)
                {
                    DataGridViewRow row = gridResults.Rows[e.RowIndex];
                    textBoxDescription.Text = (row.DataBoundItem as LinqToWikipedia.WikipediaOpenSearchResult).Description;
                }         
            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }              
        }

        private void gridResults_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (gridResults.Rows.Count > e.RowIndex)
                {
                    DataGridViewRow row = gridResults.Rows[e.RowIndex];
                    System.Diagnostics.Process.Start((row.DataBoundItem as LinqToWikipedia.WikipediaOpenSearchResult).Url.ToString());
                }       
            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }
        }

        private void textBoxKeyWords_KeyPress(object sender, KeyPressEventArgs e)
        {
            try
            {
                if (e.KeyChar == 13)
                {
                    buttonSearch_Click(buttonSearch, new EventArgs());
                    e.Handled = true;
                }
            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }
           
        }

        #endregion

        #region Word Trigger

        void Application_WindowSelectionChangeEvent(NetOffice.WordApi.Selection Sel)
        {
            try
            {
                if (checkBoxAutoSearch.Checked)
                {
                    textBoxKeyWords.Text = Sel.Text.Replace("\r", "");
                    buttonSearch_Click(buttonSearch, new EventArgs());
                }
                Sel.Dispose();
            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }
        }

        #endregion

        #region ITaskPane Member

        public void OnConnection(NetOffice.WordApi.Application application, object[] customArguments)
        {
            application.WindowSelectionChangeEvent += new NetOffice.WordApi.Application_WindowSelectionChangeEventHandler(Application_WindowSelectionChangeEvent);
        }
      
        #endregion

        #region Methods

        private void ClearError()
        {
            labelMessage.ForeColor = Color.FromKnownColor(KnownColor.DarkBlue);
            labelMessage.Text = "Double click in the result list to open the article in your Web Browser.";
        }

        private void DisplayError(Exception exception)
        {
            labelMessage.ForeColor = Color.Red;
            labelMessage.Text = "An error ocurred. Details:" + exception.Message;
        }

        private void ClearResult()
        {
            gridResults.Rows.Clear();
        }

        #endregion
    }
}
