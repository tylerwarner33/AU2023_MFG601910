Imports Inventor

Module CreateComboBoxDefintion

    Function CreateButtonDef(
        ribbonTab As RibbonTab,
        ribbonPanelName As String,
        isInButtonStack As Boolean,
        isInSlideOut As Boolean,
        buttonLabel As String,
        dropDownWidth As Long,
        toolTipSimple As String) As ComboBoxDefinition

        Dim buttonNameNoSpaces = buttonLabel.Replace(System.Environment.NewLine, "_").Replace(vbLf, "_").Replace(vbCr, "_").Replace(" ", "_")
        Dim internalName As String = buttonNameNoSpaces & "_" & g_addInClientID & "_Button_InternalName"
        Dim description As String = buttonNameNoSpaces & "_Button"
        Dim controlDefs As Inventor.ControlDefinitions = g_inventorApplication.CommandManager.ControlDefinitions

        Dim toolButtonDef As ButtonDefinition = CheckForExisting.GetButtonDef(internalName)

        If Not toolButtonDef Is Nothing Then
            Return toolButtonDef
            Exit Function
        End If

        Dim ribbon_Panel = CreateRibbonPanel.GetRibbonPanel(ribbonPanelName, ribbonTab)

        toolButtonDef = controlDefs.AddComboBoxDefinition(
            buttonLabel,
            internalName,
            CommandTypesEnum.kEditMaskCmdType,
            dropDownWidth,
            g_addInClientID)

        toolButtonDef.DescriptionText = description
        toolButtonDef.ToolTipText = toolTipSimple

        If isInButtonStack = False Then
            If isInSlideOut = False Then
                ribbon_Panel.CommandControls.AddComboBox(toolButtonDef)
            Else
                ribbon_Panel.SlideoutControls.AddComboBox(toolButtonDef)
            End If
        End If

        Return toolButtonDef

    End Function

End Module
