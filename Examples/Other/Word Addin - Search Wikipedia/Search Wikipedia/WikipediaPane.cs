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
    /// </summary>
    public partial class WikipediaPane : UserControl, ITaskPane
    {
        #region Delegates

        /// <summary>
        /// Error Async Invoker. Of course ActionT works as well but this example is also available in .NET 2.0
        /// </summary>
        /// <param name="exception">exception to display</param>
        public delegate void DisplayErrorAction(Exception exception);

        #endregion

        #region Ctor

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

        #endregion

        #region Properties

        /// <summary>
        /// The connection to wikipedia
        /// </summary>
        internal WikipediaContext Datacontext { get; private set; }
        
        #endregion

        #region Methods

        private bool WaitForCancelation(int timeOutMS)
        {
            DateTime start = DateTime.Now;
            while (backgroundWorker1.IsBusy && (DateTime.Now - start).TotalMilliseconds < timeOutMS)
            {
                ;
            }
            return !backgroundWorker1.IsBusy;
        }

        #endregion

        #region UI Trigger

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            try
            {
                // run async(otherwise search is block the Word UI) we wait a second for cancelation if backgroundworker is currently busy
                ClearError();
                if (backgroundWorker1.IsBusy)
                    backgroundWorker1.CancelAsync();
                if(WaitForCancelation(1000))
                    backgroundWorker1.RunWorkerAsync(textBoxKeyWords.Text);
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

        private void Application_WindowSelectionChangeEvent(NetOffice.WordApi.Selection Sel)
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

        public void OnConnection(NetOffice.WordApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            application.WindowSelectionChangeEvent += new NetOffice.WordApi.Application_WindowSelectionChangeEventHandler(Application_WindowSelectionChangeEvent);
        }

        public void OnDisconnection()
        { 
        
        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {

        }

        public void OnVisibleStateChanged(bool visible)
        {

        }

        #endregion

        #region Methods

        private void ClearError()
        {
            if (labelMessage.InvokeRequired)
            {
                MethodInvoker invoker = ClearError;
                invoker.Invoke();
            }
            else
            {
                labelMessage.ForeColor = Color.FromKnownColor(KnownColor.DarkBlue);
                labelMessage.Text = "Double click in the result list to open the article in your Web Browser.";
            }
        }

        private void DisplayError(Exception exception)
        {
            if (labelMessage.InvokeRequired)
            {
                DisplayErrorAction invoker = DisplayError;
                invoker.Invoke(exception);
            }
            else
            {
                labelMessage.ForeColor = Color.Red;
                labelMessage.Text = "An error ocurred. Details:" + exception.Message;
            }
        }

        private void ClearResult()
        {
            if (gridResults.InvokeRequired)
            {
                MethodInvoker invoker = ClearResult;
                invoker.Invoke();
            }
            else
            {
                gridResults.Rows.Clear();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                string search = e.Argument as string;
                var opensearch = (from wikipedia in Datacontext.OpenSearch where wikipedia.Keyword == search select wikipedia).Take(100);
                e.Result = opensearch.ToList();
            }
            catch (Exception exception)
            {
                DisplayError(exception);                
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                ClearError();
                List<WikipediaOpenSearchResult> result = e.Result as List<WikipediaOpenSearchResult>;
                gridResults.DataSource = result;
            }
            catch (Exception exception)
            {
                DisplayError(exception);
            }
        }

        #endregion
    }
}
