Imports System.Runtime.InteropServices
Imports NetOffice
Imports NetOffice.Tools
Imports Word = NetOffice.WordApi
Imports NetOffice.WordApi.Tools
Imports NetOffice.OfficeApi.Tools
'
'Document Inspector Example
'
<COMAddin("Word05 Sample Addin VB4", "Document Inspector Example", LoadBehavior.LoadAtStartup)>
<DocumentInspector("Word05AddinVB4 Inspector", "This is a sample inspector", "12,14,15,16", 1)>
<ProgId("Word05AddinVB4.Connect"), Guid("D34CE190-3E6F-454A-9121-6FC6BC9CEC09"), Codebase, Timestamp>
Public Class Addin
    Inherits Word.Tools.COMAddin

    Protected Overrides Function OnCreateDocumentInspector() As ToolsDocumentInspector

        Return New CustomDocumentInspector()

    End Function

End Class