Imports Inventor
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module CreateRibbonPanel

    Function GetRibbonPanel(panelName As String, myRibbonTab As RibbonTab) As RibbonPanel

        Dim Custom_RibbonTab As RibbonTab = CheckForExisting.GetRibbonTab(myRibbonTab.DisplayName, myRibbonTab.Parent.InternalName)

        Dim ribbon_Panel As RibbonPanel = CheckForExisting.GetRibbonPanel(Custom_RibbonTab, panelName)

        Dim panel_ID As String = "id_" & Replace(panelName, " ", "_")

        If ribbon_Panel Is Nothing Then
            ribbon_Panel = Custom_RibbonTab.RibbonPanels.Add(panelName, panel_ID, g_addInClientID)
            System.Diagnostics.Debug.WriteLine("*******  " & panelName & " panel created")
        End If

        Return ribbon_Panel

    End Function

End Module
