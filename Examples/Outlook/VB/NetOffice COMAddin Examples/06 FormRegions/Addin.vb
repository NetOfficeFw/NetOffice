Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.OutlookApi.Tools
Imports NetOffice.OutlookApi
'
'FormRegions Example
'
<COMAddin("Outlook06 Sample Addin VB4", "FormRegions Example", LoadBehavior.LoadAtStartup)>
<FormRegion("CustomFormRegion_1_VB4", "IPM.Note", "CustomFormRegion1.xml", "CustomFormRegion1.ofs", "Icon1.ico")>
<ProgId("Outlook06AddinVB4.Connect"), Guid("D39EBAEE-169B-441C-BDC3-7F73C50DBB45"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Protected Overrides Function OnCreateOpenFormRegion(form As FormRegion) As OpenFormRegion
        Return New CustomFormRegion1(form)
    End Function

End Class