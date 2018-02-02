using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.WordApi.Tools;
/*
    Custom Addin Object Example
    Demonstrate how to spend a callable instance to VBA
*/
namespace Word06AddinCS4
{
    [COMAddin("Word06 Sample Addin CS4", "Custom Addin Object Example", LoadBehavior.LoadAtStartup)]
    [ProgId("Word06AddinCS4.Connect"), Guid("CDF6C37B-D346-479D-8B39-71DDAF0BB261"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        public Addin()
        {
        }

        protected override object OnCreateObjectInstance()
        {
            return new TimeComponent();
        }
    }
}