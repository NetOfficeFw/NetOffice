using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Tools;
using NetOffice.OfficeApi.Tools;
/*
    Document Inspector Example
*/
namespace Word05AddinCS4
{   
    [COMAddin("Word05 Sample Addin CS4", "Document Inspector Example", LoadBehavior.LoadAtStartup)]
    [DocumentInspector("Word05AddinCS4 Inspector", "This is a sample inspector", "12,14,15,16", 1)]
    [ProgId("Word05AddinCS4.Connect"), Guid("E3630EC8-AB2C-42E7-A2F2-FB3757A896A8"), Codebase, Timestamp]
    public class Addin : Word.Tools.COMAddin
    {
        protected override ToolsDocumentInspector OnCreateDocumentInspector()
        {
            return new CustomDocumentInspector();
        }
    }
}
