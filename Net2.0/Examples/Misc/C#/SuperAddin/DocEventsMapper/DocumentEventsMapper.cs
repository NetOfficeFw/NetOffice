using System;
using System.Collections.Generic;
using System.Text;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using Word = NetOffice.WordApi;

namespace SuperAddin
{
    public class BeforeSaveArgs : EventArgs, IDisposable
    {
        public Word.Document Doc;
        public Excel.Workbook Wb;

        internal BeforeSaveArgs(Word.Document doc)
        {
            Doc = doc;
        }

        internal BeforeSaveArgs(Excel.Workbook wb)
        {
            Wb = wb;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (null != Doc)
                Doc.Dispose();

            if (null != Wb)
                Wb.Dispose();
        }

        #endregion
    }
    public delegate void BeforeSaveHandler(BeforeSaveArgs args, ref bool SaveAsUI, ref bool Cancel);

    public class OpenArgs : EventArgs, IDisposable
    {
        public Word.Document Doc;
        public Excel.Workbook Wb;

        internal OpenArgs(Word.Document doc)
        {
            Doc = doc;
        }

        internal OpenArgs(Excel.Workbook wb)
        {
            Wb = wb;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (null != Doc)
                Doc.Dispose();

            if (null != Wb)
                Wb.Dispose();
        }

        #endregion
    }
    public delegate void OpenHandler(OpenArgs args);

    public class BeforeCloseArgs : EventArgs, IDisposable
    {
        Word.Document Doc;
        Excel.Workbook Wb;

        internal BeforeCloseArgs(Word.Document doc)
        {
            Doc = doc;
        }

        internal BeforeCloseArgs(Excel.Workbook wb)
        {
            Wb = wb;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (null != Doc)
                Doc.Dispose();

            if (null != Wb)
                Wb.Dispose();
        }

        #endregion
    }
    public delegate void BeforeCloseHandler(BeforeCloseArgs args, ref bool Cancel);

    public class BeforePrintArgs : EventArgs, IDisposable
    {
        Word.Document Doc;
        Excel.Workbook Wb;

        internal BeforePrintArgs(Word.Document doc)
        {
            Doc = doc;
        }

        internal BeforePrintArgs(Excel.Workbook wb)
        {
            Wb = wb;
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (null != Doc)
                Doc.Dispose();

            if (null != Wb)
                Wb.Dispose();
        }

        #endregion
    } 
    public delegate void BeforePrintHandler(BeforePrintArgs args, ref bool Cancel);

    /// <summary>
    /// combines common excel and word events
    /// </summary>
    public class DocumentEventsMapper
    {
        #region Fields

        HostApplication _parent;

        #endregion

        #region Construction

        public DocumentEventsMapper(HostApplication application)
        {
            _parent = application;

            Excel.Application excelApp = _parent.Application as Excel.Application;
            if (null != excelApp)
            {
                excelApp.WorkbookBeforePrintEvent += new Excel.Application_WorkbookBeforePrintEventHandler(excelApp_WorkbookBeforePrintEvent);
                excelApp.WorkbookBeforeCloseEvent += new Excel.Application_WorkbookBeforeCloseEventHandler(excelApp_WorkbookBeforeCloseEvent);
                excelApp.WorkbookOpenEvent += new Excel.Application_WorkbookOpenEventHandler(excelApp_WorkbookOpenEvent);
                excelApp.WorkbookBeforeSaveEvent += new Excel.Application_WorkbookBeforeSaveEventHandler(excelApp_WorkbookBeforeSaveEvent);
                return;
            }

            Word.Application wordApp = _parent.Application as Word.Application;
            if (null != wordApp)
            {
                wordApp.DocumentBeforePrintEvent += new Word.Application_DocumentBeforePrintEventHandler(wordApp_DocumentBeforePrintEvent);
                wordApp.DocumentBeforeCloseEvent += new Word.Application_DocumentBeforeCloseEventHandler(wordApp_DocumentBeforeCloseEvent);
                wordApp.DocumentOpenEvent += new Word.Application_DocumentOpenEventHandler(wordApp_DocumentOpenEvent);
                wordApp.DocumentBeforeSaveEvent += new Word.Application_DocumentBeforeSaveEventHandler(wordApp_DocumentBeforeSaveEvent);
                return;
            }
        }
        
        #endregion

        #region Word Trigger

        void wordApp_DocumentBeforeSaveEvent(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                _parent.RaiseBeforeSaveEvent(new BeforeSaveArgs(Doc), ref SaveAsUI, ref Cancel);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform wordApp_DocumentBeforeSaveEvent.", throwedException);
            }   
        }

        void wordApp_DocumentOpenEvent(Word.Document Doc)
        {
            try
            {
                _parent.RaiseBeforeOpenEvent(new OpenArgs(Doc));
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform wordApp_DocumentOpenEvent.", throwedException);
            }      
        }

        void wordApp_DocumentBeforeCloseEvent(Word.Document Doc, ref bool Cancel)
        {
            try
            {
                _parent.RaiseBeforeCloseEvent(new BeforeCloseArgs(Doc), ref Cancel);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform wordApp_DocumentBeforeCloseEvent.", throwedException);
            }      
        }

        void wordApp_DocumentBeforePrintEvent(Word.Document Doc, ref bool Cancel)
        {
            try
            {
                _parent.RaiseBeforePrintEvent(new BeforePrintArgs(Doc), ref Cancel);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform wordApp_DocumentBeforePrintEvent.", throwedException);
            }      
        }

        #endregion

        #region Excel Trigger

        void excelApp_WorkbookBeforeSaveEvent(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                _parent.RaiseBeforeSaveEvent(new BeforeSaveArgs(Wb), ref SaveAsUI, ref Cancel);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform excelApp_WorkbookBeforeSaveEvent.", throwedException);
            }   
        }

        void excelApp_WorkbookOpenEvent(Excel.Workbook Wb)
        {
            try
            {
                _parent.RaiseBeforeOpenEvent(new OpenArgs(Wb));
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform excelApp_WorkbookOpenEvent.", throwedException);
            }      
        }

        void excelApp_WorkbookBeforeCloseEvent(Excel.Workbook Wb, ref bool Cancel)
        {
            try
            {
                _parent.RaiseBeforeCloseEvent(new BeforeCloseArgs(Wb), ref Cancel);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform excelApp_WorkbookBeforeCloseEvent.", throwedException);
            }      

        }

        void excelApp_WorkbookBeforePrintEvent(Excel.Workbook Wb, ref bool Cancel)
        {
            try
            {
                _parent.RaiseBeforePrintEvent(new BeforePrintArgs(Wb), ref Cancel);
            }
            catch (Exception throwedException)
            {
                FormShowError.LogError("An error ocurred while perform excelApp_WorkbookBeforePrintEvent.", throwedException);
            }      
        }

        #endregion  
    }
}
