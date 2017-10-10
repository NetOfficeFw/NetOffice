using System;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using NetOffice.OutlookApi.Tools;
/*
    FormRegions Example
*/
namespace Outlook06AddinCS4
{
    [COMAddin("Outlook06 Sample Addin CS4", "FormRegions Example", LoadBehavior.LoadAtStartup)]
    [FormRegion("CustomFormRegion_1_CS4", "IPM.Note", "CustomFormRegion1.xml", "CustomFormRegion1.ofs", "Icon1.ico")]
    [ProgId("Outlook06AddinCS4.Connect"), Guid("ABDD3141-9E5C-4306-96B4-B705F94C1C55"), Codebase, Timestamp]
    public class Addin : COMAddin
    {
        protected override OpenFormRegion OnCreateOpenFormRegion(Outlook.FormRegion form)
        {
            return new CustomFormRegion1(form);
        }
    }
}