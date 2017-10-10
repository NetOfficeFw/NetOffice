Imports System.Runtime.InteropServices
Imports NetOffice
Imports NetOffice.Tools
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Tools
'
'Ribbons & Panes Addin Example
'
<COMAddin("Excel02 Sample Addin VB4", "Ribbons & Panes Addin Example", LoadBehavior.LoadAtStartup)>
<ProgId("Excel02AddinVB4.Connect"), Guid("ED20A119-9E0D-4D7B-B50A-71CFFD44A7C7"), Codebase, Timestamp>
<CustomUI("RibbonUI.xml", True)>
<CustomPane(GetType(SamplePane), "Excel CPU Usage", False, PaneDockPosition.msoCTPDockPositionTop, PaneDockPositionRestrict.msoCTPDockPositionRestrictNoVertical, 60, 60)>
Public Class Addin
    Inherits COMAddin

    ' taskpane visibility has been changed. we upate the checkbutton in the ribbon ui for show/hide taskpane
    Protected Overrides Sub TaskPaneVisibleStateChanged(ByVal customTaskPaneInst As NetOffice.OfficeApi._CustomTaskPane)

        If Not IsNothing(RibbonUI) Then
            RibbonUI.InvalidateControl("paneVisibleToogleButton")
        End If

    End Sub

    '  defined in RibbonUI.xml to make sure the checkbutton state is up-to-date and synchronized with taskpane visibility.
    Public Function OnGetPressedPanelToggle(ByVal control As Office.IRibbonControl) As Boolean

        If TaskPanes.Count > 0 Then
            Return TaskPanes(0).Visible
        Else
            Return False
        End If

    End Function

    ' defined in RibbonUI.xml to track the user clicked ouer checkbutton. then we upate the panel visibility at hand.
    Public Sub OnCheckPanelToggle(ByVal control As Office.IRibbonControl, ByVal pressed As Boolean)

        If TaskPanes.Count > 0 Then
            TaskPanes(0).Visible = pressed
        End If

    End Sub

    ' defined in RibbonUI.xml to catch the user click for the about button
    Public Sub OnClickAboutButton(ByVal control As Office.IRibbonControl)

        Utils.Dialog.ShowDiagnostics()

    End Sub

End Class