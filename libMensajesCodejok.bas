Attribute VB_Name = "libMensajes"
Option Explicit




Public Function MsgBoxC(TEXTO As String, Botones As Long, Optional Titulo As String, Optional Extendido As Boolean) As Long
'   Botones
'--------------------------------------------------------
'    vbOKOnly 0 Sólo el botón Aceptar (predeterminado)
'    vbOKCancel 1 Los botones Aceptar y Cancelar
'    vbYesNoCancel 3 Los botones Sí, No y Cancelar.
'    VbYesNo 4 Los botones Sí y No
'   vbCritical 16 Mensaje crítico
'   vbQuestion 32 Consulta de advertencia
'   vbExclamation 48 Mensaje de advertencia
'   vbInformation 64 Mensaje de información
    
    Dim B As Boolean
    Dim miBoton As String
    If Botones = vbInformation Then
        'Solo quiere el boton de aceptar
         miBoton = "Aceptar|"
    Else
        miBoton = ""
    End If
    
    
    
    MuestraMsgCodejock2 "Ariconta6", Titulo, TEXTO, "", miBoton, Botones, Extendido
        
    MsgBoxC = RespuestaMsgBox
    

    
End Function





Public Function MsgBoxA(TEXTO As String, Botones As Long, Optional Titulo As String, Optional Extendido As Boolean) As Long
'   Botones
'--------------------------------------------------------
'    vbOKOnly 0 Sólo el botón Aceptar (predeterminado)
'    vbOKCancel 1 Los botones Aceptar y Cancelar
'    vbYesNoCancel 3 Los botones Sí, No y Cancelar.
'    VbYesNo 4 Los botones Sí y No
'   vbCritical 16 Mensaje crítico
'   vbQuestion 32 Consulta de advertencia
'   vbExclamation 48 Mensaje de advertencia
'   vbInformation 64 Mensaje de información
    
    Dim B As Boolean
    Dim miBoton As String
    If Botones = vbInformation Then
        'Solo quiere el boton de aceptar
         miBoton = "Aceptar|"
    Else
        miBoton = ""
    End If
    
    
    
    'MuestraMsgCodejock2 "Ariconta6", Titulo, TEXTO, "", miBoton, Botones, Extendido
     
    MuestraMsgAriadna "Ariconta6", Titulo, TEXTO, "", miBoton, Botones, Extendido
    MsgBoxA = RespuestaMsgBox
End Function








'Ejemplo codejock
'  C:\Program Files (x86)\Codejock Software\ActiveX\Xtreme SuitePro ActiveX v17.2.0\Samples\Controls\VB\VistaTaskDialog
'
'
'
Public Sub MuestraMsgCodejock2(Titulo As String, Ppal As String, Contenido As String, Pie As String, BtnPersonalizados As String, Botones As Long, Extendido As Boolean)
Dim B As Boolean
Dim i As Integer
Dim K As Integer
Dim C As String
Dim Icono As Long
Dim BtnPorDefecto As Long


    Load frmMenBox

    frmMenBox.TaskDialog1.Reset
    
    'Always look like manifest was used, even if no manifest file is present
    frmMenBox.TaskDialog1.MessageBoxStyle = True
    
    frmMenBox.TaskDialog1.WindowTitle = Titulo
    If Extendido Then frmMenBox.TaskDialog1.DialogWidth = 250
    
    frmMenBox.TaskDialog1.MainInstructionText = Ppal
    frmMenBox.TaskDialog1.ContentText = Contenido
    frmMenBox.TaskDialog1.FooterText = Pie
    
    'frmMenBox.TaskDialog1.VerificationText = editVerification.Text
    
       
    
    
    'From Expanded Tab
    'frmMenBox.TaskDialog1.ExpandedInformationText = editExpandedInformation.Text
    'frmMenBox.TaskDialog1.ExpandedControlText = editExpandedControlText.Text
    'frmMenBox.TaskDialog1.CollapsedControlText = editCollapsedControlText.Text

    '-------------------------------------------------------------------------
    ' General tab.
    '-------------------------------------------------------------------------
    
    'If chkUsePreferredWidth.Value = xtpChecked Then
    '    frmMenBox.TaskDialog1.DialogWidth = Int(editPreferredWidth.Text)
    'End If

    'frmMenBox.TaskDialog1.VerifyCheckState = IIf(chkVerificationChecked.Value = xtpChecked, True, False)
    'frmMenBox.TaskDialog1.EnableHyperlinks = IIf(chkEnableHyperlinks.Value = xtpChecked, True, False)
    'frmMenBox.TaskDialog1.AllowCancellation = IIf(chkAllowDialogCancellation.Value = xtpChecked, True, False)
    'frmMenBox.TaskDialog1.RelativePosition = IIf(chkPositionRelativeToWindow.Value = xtpChecked, True, False)
    'frmMenBox.TaskDialog1.RtlLayout = IIf(chkRightToLeftLayout.Value = xtpChecked, True, False)
    
    '-------------------------------------------------------------------------
    ' Expanded tab.
    '-------------------------------------------------------------------------

    'frmMenBox.TaskDialog1.ExpandedByDefault = IIf(chkExpandedByDefault.Value = xtpChecked, True, False)
    'frmMenBox.TaskDialog1.ExpandFooterArea = IIf(chkExpandedFooterArea.Value = xtpChecked, True, False)
    
    '-------------------------------------------------------------------------
    ' Buttons tab.
    '-------------------------------------------------------------------------
    'EditMode = False
    frmMenBox.TaskDialog1.CommonButtons = 0
    frmMenBox.TaskDialog1.DefaultButton = -1
    RespuestaMsgBox = -1   'Ya la establezco aqui
    
    If BtnPersonalizados <> "" Then
        'Botones personalizados + cancel
        i = 1
        Do
            K = InStr(i, BtnPersonalizados, "|")
            If K > 0 Then
                C = Mid(BtnPersonalizados, i, K - i)
                frmMenBox.TaskDialog1.AddButton C, 30000 + i    'Int(listCustomCommandButtons.ListItems(i).ListSubItems(1).Text)
                i = K + 1
            End If
        Loop Until K = 0
        If Botones <> 64 Then
            'No lleva el vbinformation
            frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonCancel
            frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonCancel
        End If
    Else
        
        If Botones = 0 Then
            'Solo OK
            frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonOk
            frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonOk
    
        Else
           'Dependera de lo que hay enviado en botones
           '    vbOKOnly 0 Sólo el botón Aceptar (predeterminado)
            '    vbOKCancel 1 Los botones Aceptar y Cancelar
            '    vbYesNoCancel 3 Los botones Sí, No y Cancelar.
            '    VbYesNo 4 Los botones Sí y No
            BtnPorDefecto = -1
            If (Botones And 0) > 0 Then
                frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonOk
                BtnPorDefecto = xtpTaskButtonOk
            End If
            B = (Botones And 4) = 4 Or (Botones And 3) = 3
            If B Then
                frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonYes
                If BtnPorDefecto < 0 Then BtnPorDefecto = xtpTaskButtonYes
            End If
            B = (Botones And 4) = 4 Or (Botones And 3) = 3
            If B Then
                frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonNo
                If BtnPorDefecto < 0 Then BtnPorDefecto = xtpTaskButtonNo
            End If
            B = (Botones And 1) = 1 Or (Botones And 3) = 3
            If B Then
                frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonCancel
                If BtnPorDefecto < 0 Then BtnPorDefecto = xtpTaskButtonCancel
            End If
        
        End If
        
    End If
   
    
    
    
    'frmMenBox.TaskDialog1.EnableCommandLinks = IIf(chkUseCommandLinks = xtpChecked, True, False)
    'frmMenBox.TaskDialog1.ShowCommandLinkIcons = IIf(chkShowCommandLinkIcons = xtpChecked, True, False)
    
    
    'Default button
    
    
    If BtnPersonalizados <> "" Then
        frmMenBox.TaskDialog1.DefaultButton = 30000 + 1
        
    Else
        
        If Botones - 256 >= 0 Then
            'Ha seleecionado el boton pordefecto
            'Si es mayor que 256 es el boton 2 por defecto
            ' si es mayor=512 en tonces es el cancelar
            B = True
            If Botones - 512 >= 0 Then
                'Tercer boton,. el de cancelar
                frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonCancel
            Else
                frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonNo
            End If
                  
        Else
            If BtnPorDefecto >= 0 Then frmMenBox.TaskDialog1.DefaultButton = BtnPorDefecto
        End If
    End If
    
    '-------------------------------------------------------------------------
    ' Icons tab.
    '-------------------------------------------------------------------------
    
    '   vbCritical 16 Mensaje crítico
    '   vbQuestion 32 Consulta de advertencia
    '   vbExclamation 48 Mensaje de advertencia
    '   vbInformation 64 Mensaje de información
    
    'cmbMainIcon.AddItem "None", xtpTaskIconNone
    'cmbMainIcon.AddItem "Warning Icon", xtpTaskIconWarning
    'cmbMainIcon.AddItem "Error Icon", xtpTaskIconError
    'cmbMainIcon.AddItem "Information Icon", xtpTaskIconInformation
    'cmbMainIcon.AddItem "Shield Icon", xtpTaskIconShield
    'cmbMainIcon.AddItem "Question Icon", xtpTaskIconQuestion
        
        
    If Botones > 256 Then Botones = Botones - IIf(Botones > 511, 512, 256)
    If Botones >= 64 Then
        Icono = xtpTaskIconInformation
    ElseIf Botones >= 48 Then
        Icono = xtpTaskIconWarning
    ElseIf Botones >= 32 Then
        Icono = xtpTaskIconQuestion
    ElseIf Botones >= 16 Then
        Icono = xtpTaskIconError
    Else
        Icono = -1
    End If
    
    If Icono = -1 Then
        'Para el pie
        Dim nTemp As Long
        'If (editMainCustomIconPath.Text = "") Then
        '    frmMenBox.TaskDialog1.MainIcon = cmbMainIcon.ListIndex
        'Else
        nTemp = 0
        frmMenBox.TaskDialog1.MainIcon = xtpTaskIconCustom
        nTemp = LoadIcon(App.Path & "\styles\Ariconta.ico", 0, 0)
        If Not nTemp = 0 Then
            frmMenBox.TaskDialog1.MainIconHandle = nTemp
        Else
            frmMenBox.TaskDialog1.MainIconHandle = frmppal.Icon
        End If
    Else
        frmMenBox.TaskDialog1.MainIcon = Icono
    End If
    
    'End If
    
    'If (editFooterCustomIconPath.Text = "") Then
    '    frmMenBox.TaskDialog1.FooterIcon = cmbFooterIcon.ListIndex
    'Else
    '    nTemp = 0
    '    frmMenBox.TaskDialog1.FooterIcon = xtpTaskIconCustom
    '    nTemp = LoadIcon(editFooterCustomIconPath.Text, 0, 0)
    '    If Not nTemp = 0 Then
    '        frmMenBox.TaskDialog1.FooterIconHandle = nTemp
    '    Else
    '        frmMenBox.TaskDialog1.FooterIconHandle = Me.Icon
    '    End If
    'End If
    
    frmMenBox.Show vbModal
    'St op
    
End Sub






'SOLO ACEPTAMOS 2 botones personalizados
Public Sub MuestraMsgAriadna(Titulo As String, Ppal As String, Contenido As String, Pie As String, BtnPersonalizados As String, Botones As Long, Extendido As Boolean)
Dim B As Boolean
Dim i As Integer
Dim K As Integer
Dim C As String
Dim Icono As Long
Dim N As Integer
Dim BtnPorDefecto As Long


    Load frmMenBoxAriadna

    frmMenBoxAriadna.Limpiar
    
    If Titulo <> "" Then frmMenBoxAriadna.Caption = Titulo
    
    
    'If Extendido Then frmMenBox.TaskDialog1.DialogWidth = 250
    
    'frmMenBox.TaskDialog1.MainInstructionText = Ppal
    frmMenBoxAriadna.lblTitulo.Caption = Ppal
    frmMenBoxAriadna.lblTexto.Caption = Contenido
    frmMenBoxAriadna.lblPie.Caption = Pie
    
    
    
    RespuestaMsgBox = -1   'Ya la establezco aqui
    
    If BtnPersonalizados <> "" Then
        'Botones personalizados + cancel
        N = 1
        i = 1
        Do
            K = InStr(i, BtnPersonalizados, "|")
            If K > 0 Then
                C = Mid(BtnPersonalizados, i, K - i)
                If N < 3 Then
                    'frmMenBox.TaskDialog1.AddButton C, 30000 + i    'Int(listCustomCommandButtons.ListItems(i).ListSubItems(1).Text)
                    
                    frmMenBoxAriadna.cmdOtro(N).Caption = C
                    frmMenBoxAriadna.cmdOtro(N).Width = IIf(Len(C) > 11, 1500, 1335)
                    frmMenBoxAriadna.cmdOtro(N).Tag = 30000 + i
                    If N = 1 Then
                        frmMenBoxAriadna.cmdOtro1visible = True
                    Else
                        frmMenBoxAriadna.cmdOtro2visible = True
                    End If
                    N = N + 1
                Else
                    MsgBox "2 botones personalizados", vbCritical
                End If
                i = K + 1
            End If
        Loop Until K = 0
        If Botones <> 64 Then
            'No lleva el vbinformation
            'frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonCancel
            'frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonCancel
            frmMenBoxAriadna.cmdCancelvisible = True
            frmMenBoxAriadna.cmdCancel.Default = True
        End If
    Else
        
        If Botones = 0 Then
            'Solo OK
            'frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonOk
            'frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonOk
                frmMenBoxAriadna.cmdOkvisible = True
                frmMenBoxAriadna.cmdOk.Default = True
                
        Else
           'Dependera de lo que hay enviado en botones
           '    vbOKOnly 0 Sólo el botón Aceptar (predeterminado)
            '    vbOKCancel 1 Los botones Aceptar y Cancelar
            '    vbYesNoCancel 3 Los botones Sí, No y Cancelar.
            '    VbYesNo 4 Los botones Sí y No
            BtnPorDefecto = -1
            If (Botones And 0) > 0 Then
                'frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonOk
                frmMenBoxAriadna.cmdOkvisible = True
                frmMenBoxAriadna.cmdOk.Default = True
                BtnPorDefecto = xtpTaskButtonOk
            End If
            B = (Botones And 4) = 4 Or (Botones And 3) = 3
            If B Then
                'frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonYes
                frmMenBoxAriadna.cmdYesvisible = True
                If BtnPorDefecto < 0 Then
                    BtnPorDefecto = xtpTaskButtonYes
                    frmMenBoxAriadna.cmdYes.Default = True
                End If
            End If
            B = (Botones And 4) = 4 Or (Botones And 3) = 3
            If B Then
                'frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonNo
                frmMenBoxAriadna.cmdNovisible = True
                If BtnPorDefecto < 0 Then
                    frmMenBoxAriadna.cmdNo.Default = True
                    BtnPorDefecto = xtpTaskButtonNo
                End If
            End If
            B = (Botones And 1) = 1 Or (Botones And 3) = 3
            If B Then
                'frmMenBox.TaskDialog1.CommonButtons = frmMenBox.TaskDialog1.CommonButtons Or xtpTaskButtonCancel
                frmMenBoxAriadna.cmdCancelvisible = True
                If BtnPorDefecto < 0 Then
                    BtnPorDefecto = xtpTaskButtonCancel
                    frmMenBoxAriadna.cmdCancel.Default = True
                End If
            End If


            If BtnPorDefecto = -1 Then
                frmMenBoxAriadna.cmdOkvisible = True
                frmMenBoxAriadna.cmdOk.Default = True
            End If
        End If
        
    End If
   
    
    
    
    
    
    If BtnPersonalizados <> "" Then
        'frmMenBox.TaskDialog1.DefaultButton = 30000 + 1
        
    Else
        
        If Botones - 256 >= 0 Then
            'Ha seleecionado el boton pordefecto
            'Si es mayor que 256 es el boton 2 por defecto
            ' si es mayor=512 en tonces es el cancelar
            B = True
            If Botones - 512 >= 0 Then
                'Tercer boton,. el de cancelar
                'frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonCancel
                frmMenBoxAriadna.cmdCancel.Default = True
            Else
                'frmMenBox.TaskDialog1.DefaultButton = xtpTaskButtonNo
                frmMenBoxAriadna.cmdNo.Default = True
            End If
                  
        Else
            If BtnPorDefecto >= 0 Then
                'frmMenBox.TaskDialog1.DefaultButton = BtnPorDefecto
            End If
        End If
    End If
    
    '-------------------------------------------------------------------------
    ' Icons tab.
    '-------------------------------------------------------------------------
    
    '   vbCritical 16 Mensaje crítico
    '   vbQuestion 32 Consulta de advertencia
    '   vbExclamation 48 Mensaje de advertencia
    '   vbInformation 64 Mensaje de información
    
    'cmbMainIcon.AddItem "None", xtpTaskIconNone
    'cmbMainIcon.AddItem "Warning Icon", xtpTaskIconWarning
    'cmbMainIcon.AddItem "Error Icon", xtpTaskIconError
    'cmbMainIcon.AddItem "Information Icon", xtpTaskIconInformation
    'cmbMainIcon.AddItem "Shield Icon", xtpTaskIconShield
    'cmbMainIcon.AddItem "Question Icon", xtpTaskIconQuestion
        
        
    If Botones > 256 Then Botones = Botones - IIf(Botones > 511, 512, 256)
    If Botones >= 64 Then
        'Icono = xtpTaskIconInformation
        Icono = 0
    ElseIf Botones >= 48 Then
        'Icono = xtpTaskIconWarning
        Icono = 2
    ElseIf Botones >= 32 Then
        'Icono = xtpTaskIconQuestion
        Icono = 1
    ElseIf Botones >= 16 Then
        'Icono = xtpTaskIconError
        Icono = 3
    Else
        Icono = 4 'Icono de Ariadna
    End If
    
    'frmMenBox.TaskDialog1.MainIcon = Icono
    frmMenBoxAriadna.Image1(Icono).visible = True
    
    
    frmMenBoxAriadna.AjustaTamañosYPosicion
'    frmMenBox.Show vbModal
    frmMenBoxAriadna.Show vbModal
    
End Sub























Function LoadIcon(Path As String, cx As Long, cy As Long) As Long
Dim LR_LOADFROMFILE As Long
Dim IMAGE_ICON As Long
    IMAGE_ICON = 0
    LR_LOADFROMFILE = &H10
    LoadIcon = LoadImage(App.hInstance, Path, IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
End Function





































'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'
'
'     Trozo de abajo por si queremos poner en marcha las notificaciones POPUP
'
'
'
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------------------------------------








'Para los mensajes. Estaba en frmmain
'Private Const ID_POPUP0 = 0
'Dim HaLanzadoPopUp2 As Byte
'Private TextoMensaje As String
'Private EncabezadoMsg As String
'Const IDOK = 1
'Const IDCLOSE = 2
'Const IDSITE = 3
'Const IDMINIMIZE = 4



'''''''***************************************************************************************************
'''''''***************************************************************************************************
'''''''***************************************************************************************************
'''''''    Mostrajes OFFICE
'''''''   Ver el proyecto original en :  codjecock\samples\control\vb\popup
'''''''
'''''''***************************************************************************************************
'''''''***************************************************************************************************
'''''''***************************************************************************************************
'''''''
''''''' HaLanzadoElCero  : significa que ha lanzado un popaup hace un momento, entonces tiene que cojer
''''''' el index 1
''''''Public Sub MostrarMensaje(Tipo As Integer, Encabezado As String, TextoMens As String, HaLanzadoElCero As Boolean)
''''''
''''''
''''''    On Error Resume Next
''''''
''''''    Dim X As Integer
''''''    Dim lastPane As Integer
''''''
''''''
''''''
''''''    TextoMensaje = TextoMens
''''''    EncabezadoMsg = Encabezado
''''''    'lastPane = IIf(chkMultiplePopup, ID_POPUP2, ID_POPUP0)
''''''    If HaLanzadoElCero Then
''''''        lastPane = 1
''''''    Else
''''''        lastPane = ID_POPUP0
''''''
''''''    End If
''''''
''''''    For X = lastPane To lastPane
''''''
''''''
''''''
''''''        Dim Popup As XtremeSuiteControls.PopupControl
''''''        Set Popup = PopupControl1(X)
''''''
''''''        Popup.Animation = 2  'cmbAnimation.ListIndex
''''''        Popup.AnimateDelay = 250
''''''        Popup.ShowDelay = 2000
''''''        Popup.Transparency = 200
''''''        Popup.DefaultLocation = 0
''''''
''''''
''''''        Select Case Tipo
''''''   '         Case 0: SetOffice2000Theme Popup
''''''   '         Case 1: SetOfficeXPTheme Popup
''''''   '         Case 2: SetOffice2003Theme Popup
''''''   '         Case 3: SetOffice2007Theme Popup
''''''   '         Case 4: SetOffice2013Theme Popup
''''''   '         Case 5: SetMSNTheme Popup
''''''   '         Case 6: SetBlueTheme Popup
''''''            Case 7: SetRedTheme Popup
''''''
''''''
''''''   '         Case 8: SetGreenTheme Popup
''''''
''''''            Case 9: SetBlackTheme Popup
''''''
''''''
''''''   '        Case 10: SetToolTipTheme Popup
''''''        End Select
''''''    Next
''''''
''''''    DesEnablar
''''''
''''''    If lastPane = 0 Then PopupControl1(lastPane).Show
''''''
''''''
''''''    If lastPane = 1 Then
''''''        PopupControl1(lastPane).Right = PopupControl1(ID_POPUP0).Right
''''''        PopupControl1(lastPane).Bottom = (PopupControl1(ID_POPUP0).Bottom - PopupControl1(ID_POPUP0).Height)
''''''        PopupControl1(lastPane).AnimateDelay = PopupControl1(ID_POPUP0).AnimateDelay + 256
''''''        PopupControl1(lastPane).ShowDelay = PopupControl1(ID_POPUP0).ShowDelay + 1000
''''''        PopupControl1(lastPane).Show
''''''
''''''    End If
''''''    'If chkMultiplePopup Then
''''''    '    PopupControl(ID_POPUP1).Right = PopupControl(ID_POPUP0).Right
''''''    '    PopupControl(ID_POPUP1).Bottom = (PopupControl(ID_POPUP0).Bottom - PopupControl(ID_POPUP0).Height)
''''''    '    PopupControl(ID_POPUP1).AnimateDelay = PopupControl(ID_POPUP0).AnimateDelay + 256
''''''    '    PopupControl(ID_POPUP1).ShowDelay = PopupControl(ID_POPUP0).ShowDelay + 1000
''''''    '    PopupControl(ID_POPUP1).Show
''''''    '
''''''    '    PopupControl(ID_POPUP2).Right = PopupControl(ID_POPUP1).Right
''''''    '    PopupControl(ID_POPUP2).Bottom = (PopupControl(ID_POPUP1).Bottom - PopupControl(ID_POPUP1).Height)
''''''    '    PopupControl(ID_POPUP2).AnimateDelay = PopupControl(ID_POPUP1).AnimateDelay + 256
''''''    '    PopupControl(ID_POPUP2).ShowDelay = PopupControl(ID_POPUP1).ShowDelay + 1000
''''''    '    PopupControl(ID_POPUP2).Show
''''''    'End If
''''''
''''''End Sub
''''''
''''''Private Sub DesEnablar()
''''''    On Error GoTo eDese
''''''    Dim Control
''''''    For Each Control In Controls
''''''        Control.Enabled = False
''''''    Next
''''''    Exit Sub
''''''eDese:
''''''    If Err.Number <> 438 Then MuestraError Err.Number
''''''End Sub
''''''
''''''
''''''Sub SetRedTheme(Popup As XtremeSuiteControls.PopupControl)
''''''    Dim Item As PopupControlItem
''''''
''''''    Popup.RemoveAllItems
''''''    Popup.Icons.RemoveAll
''''''
''''''    Set Item = Popup.AddItem(0, 0, 170, 130, "", RGB(255, 50, 50), RGB(255, 255, 0))
''''''
''''''    Set Item = Popup.AddItem(5, 25, 170 - 5, 130 - 5, "", RGB(230, 70, 70), RGB(255, 255, 0))
''''''
''''''    'Set Item = Popup.AddItem(104, 27, 170, 45, "more...")
''''''
''''''    Set Item = Popup.AddItem(0, 70, 170, 100, TextoMensaje)
''''''    Item.TextAlignment = DT_CENTER Or DT_WORDBREAK
''''''    Item.TextColor = RGB(255, 255, 0)
''''''    Item.CalculateHeight
''''''    Item.id = IDSITE
''''''
''''''    Set Item = Popup.AddItem(12, 30, 12, 47, "")
''''''    'Item.SetIcon LoadIcon("Icons\icon3.ico", 32, 32), xtpPopupItemIconNormal
''''''
''''''    Set Item = Popup.AddItem(5, 0, 170, 25, EncabezadoMsg)
''''''    Item.TextAlignment = DT_SINGLELINE Or DT_VCENTER
''''''    Item.TextColor = RGB(255, 255, 255)
''''''    Item.Bold = True
''''''    Item.Hyperlink = False
''''''
''''''    Set Item = Popup.AddItem(151, 6, 164, 19, "")
''''''    'Item.SetIcons LoadBitmap("Icons\CloseMSN.bmp"), 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
''''''    Item.id = IDCLOSE
''''''
''''''    Popup.VisualTheme = xtpPopupThemeCustom
''''''    Popup.setSize 170, 130
''''''
''''''End Sub
''''''Sub SetBlackTheme(Popup As XtremeSuiteControls.PopupControl)
''''''    Dim Item As PopupControlItem
''''''
''''''    Popup.RemoveAllItems
''''''    Popup.Icons.RemoveAll
''''''
''''''    Set Item = Popup.AddItem(0, 0, 170, 130, "", RGB(10, 10, 10), RGB(255, 255, 255))
''''''
''''''    Set Item = Popup.AddItem(5, 25, 170 - 5, 130 - 5, "", RGB(70, 70, 70), RGB(200, 200, 200))
''''''
''''''    Set Item = Popup.AddItem(104, 27, 170, 45, "")
''''''    Item.TextColor = RGB(150, 150, 150)
''''''
''''''    Set Item = Popup.AddItem(0, 70, 170, 100, TextoMensaje)
''''''    Item.TextAlignment = DT_CENTER Or DT_WORDBREAK
''''''    Item.TextColor = RGB(255, 255, 255)
''''''    Item.CalculateHeight
''''''    Item.id = IDSITE
''''''
''''''    Set Item = Popup.AddItem(12, 30, 12, 47, "")
''''''    'Item.SetIcon LoadIcon("Icons\icon3.ico", 32, 32), xtpPopupItemIconNormal
''''''
''''''    Set Item = Popup.AddItem(5, 0, 170, 25, EncabezadoMsg)
''''''    Item.TextAlignment = DT_SINGLELINE Or DT_VCENTER
''''''    Item.TextColor = RGB(255, 255, 255)
''''''    Item.Bold = True
''''''    Item.Hyperlink = False
''''''
''''''    'Set Item = Popup.AddItem(151, 6, 164, 19, "")
''''''    'Item.SetIcons LoadBitmap("Icons\CloseMSN.bmp"), 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
''''''    'Item.id = IDCLOSE
''''''
''''''    Popup.VisualTheme = xtpPopupThemeCustom
''''''    Popup.setSize 170, 130
''''''
''''''End Sub
''''''
''''''
''''''
''''''

