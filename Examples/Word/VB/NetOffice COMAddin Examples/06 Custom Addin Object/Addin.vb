Imports System.Runtime.InteropServices
Imports NetOffice.Tools
Imports NetOffice.WordApi.Tools
'
'Custom Addin Object Example
'Demonstrate how to spend a callable instance to VBA
'
<COMAddin("Word06 Sample Addin VB4", "Custom Addin Object Example", LoadBehavior.LoadAtStartup)>
<ProgId("Word06AddinVB4.Connect"), Guid("C3D40A14-3845-4616-A2CF-EE6CF909B251"), Codebase, Timestamp>
Public Class Addin
    Inherits COMAddin

    Protected Overrides Function OnCreateObjectInstance() As Object

        Return New TimeComponent()

    End Function

End Class
