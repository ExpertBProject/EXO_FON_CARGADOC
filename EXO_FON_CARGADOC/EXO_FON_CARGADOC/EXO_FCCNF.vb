Imports System.Xml
Imports SAPbouiCOM

Public Class EXO_FCCNF
    Inherits EXO_UIAPI.EXO_DLLBase

    Public Sub New(ByRef oObjGlobal As EXO_UIAPI.EXO_UIAPI, ByRef actualizar As Boolean, usaLicencia As Boolean, idAddOn As Integer)
        MyBase.New(oObjGlobal, actualizar, False, idAddOn)
        cargamenu()
        If actualizar Then
            cargaCampos()
        End If
    End Sub
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        objGlobal.SBOApp.LoadBatchActions(menuXML)
        Dim res As String = objGlobal.SBOApp.GetLastBatchResults
        Path = objGlobal.refDi.OGEN.pathGeneral.ToString.Trim
        If objGlobal.SBOApp.Menus.Exists("EXO-MnCDoc") = True Then
            Path &= "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnCDOC.png") = True Then
                    objGlobal.SBOApp.Menus.Item("EXO-MnCDoc").Image = Path & "\MnCDOC.png"
                End If
            End If
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.funciones.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Private Sub cargaCampos()
        If objGlobal.refDi.comunes.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing
            'MnCFRP
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_FCCNF.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_FCCNF", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Cargamos campos personalizados en las lineas de facturas
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDFs_INV1.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDFs_INV1", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'Cargamos campos temporales
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOC.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UT_EXO_TMPDOC", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCL.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UT_EXO_TMPDOCL", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCLT.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UT_EXO_TMPDOCLT", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If
    End Sub

    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCFCF"
                        'Cargamos UDO Configurador.
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_FCCNF")
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(infoEvento As ItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_FCCNF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                    If EventHandler_COMBO_SELECT(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_FCCNF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_FCCNF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                    If EventHandler_LOST_FOCUS(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_FCCNF"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_Before(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_Before(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Dim oConds As SAPbouiCOM.Conditions = Nothing
        Dim oCond As SAPbouiCOM.Condition = Nothing
        Dim iEntra As Integer = 0
        EventHandler_Choose_FromList_Before = False
        Dim sChoosefromlist As String = ""

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            oConds = New SAPbouiCOM.Conditions

            If pVal.ItemUID <> "15_U_E" And pVal.ItemUID <> "txt_IC" Then
                Select Case pVal.ItemUID
                    Case "txtCSRV" 'CCC de SRV por defecto
                        oCond = oConds.Add
                        oCond.Alias = "ActType" ' Propiedad Cliente principal
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                        oCond.CondVal = "N"
                        sChoosefromlist = "CFL_CSRV"
                    Case "txtIV" 'Impuesto de ventas por defecto
                        oCond = oConds.Add
                        oCond.Alias = "Category" ' Categoría del impuesto
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "O"
                        sChoosefromlist = "CFL_IV"
                    Case "txtIC" 'Impuesto de compras por defecto
                        oCond = oConds.Add
                        oCond.Alias = "Category" ' Categoría del impuesto
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = "I"
                        sChoosefromlist = "CFL_IC"
                End Select
                oForm.ChooseFromLists.Item(sChoosefromlist).SetConditions(oConds)
            End If

            EventHandler_Choose_FromList_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oConds, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oCond, Object))
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If oForm.Visible = True Then
                sSQL = "SELECT * FROM ""@EXO_CFCNF"" "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    oForm.Mode = BoFormMode.fm_OK_MODE
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objGlobal.SBOApp.ActivateMenuItem("1290") ' Ir al primer registro
                    ElseIf oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        CargaComboLínea(oForm, 1, CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString)
                        CType(oForm.Items.Item("O_U_E").Specific, SAPbouiCOM.EditText).Active = True
                    End If
                Else
                    oForm.Mode = BoFormMode.fm_ADD_MODE
                End If
                HabDesHabCampos(pVal, CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString)
                CargaComboSerie(oForm)
                CargaComboViaPago(oForm)
                CargaComboCPago(oForm)
            End If

            EventHandler_FORM_VISIBLE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function HabDesHabCampos(ByRef pVal As ItemEvent, ByVal sValorCampo As String) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        HabDesHabCampos = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If oForm.Visible = True And oForm.TypeEx = "UDO_FT_EXO_FCCNF" Then
                Select Case sValorCampo
                    Case "1" 'TXT
                        CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Enabled = False
                        CType(oForm.Items.Item("17_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Enabled = True
                        CType(oForm.Items.Item("17_U_Cb").Specific, SAPbouiCOM.ComboBox).Select("1", BoSearchKey.psk_ByValue)
                    Case "2" 'Excel
                        CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Enabled = True
                        oForm.DataSources.DBDataSources.Item("@EXO_CFCNF").SetValue("U_EXO_STXT", 0, "0")
                        CType(oForm.Items.Item("17_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Enabled = False
                    Case "3" 'XML
                        CType(oForm.Items.Item("16_U_E").Specific, SAPbouiCOM.EditText).Item.Enabled = False
                        oForm.DataSources.DBDataSources.Item("@EXO_CFCNF").SetValue("U_EXO_STXT", 0, "0")
                        CType(oForm.Items.Item("17_U_Cb").Specific, SAPbouiCOM.ComboBox).Item.Enabled = False
                End Select
            End If

            HabDesHabCampos = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_COMBO_SELECT(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_COMBO_SELECT = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If oForm.Mode = BoFormMode.fm_ADD_MODE Or oForm.Mode = BoFormMode.fm_UPDATE_MODE Then
                If pVal.ItemChanged = True And pVal.ItemUID = "14_U_Cb" Then ' Tipo Fichero a importar
                    HabDesHabCampos(pVal, CType(oForm.Items.Item("14_U_Cb").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString)
                End If
                If (pVal.ItemUID = "0_U_G" And pVal.ColUID = "C_0_1") Then 'Cod. Campo según Tipo Cabecera
                    If CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString <> "" Then
                        If pVal.Row <= 0 Then
                            CargaComboLínea(oForm, 1, CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString)
                        Else
                            CargaComboLínea(oForm, pVal.Row, CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString)
                        End If
                    Else
                        objGlobal.SBOApp.MessageBox("No ha seleccionado la estructura de ficheros." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione Cód. de los Parámetros de Campos de SAP.")
                        objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha seleccionado la estructura de ficheros." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione Cód. de los Parámetros de Campos de SAP.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Active = True
                    End If
                End If

            End If
            EventHandler_COMBO_SELECT = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_LOST_FOCUS(ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        EventHandler_LOST_FOCUS = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)
            If pVal.ItemUID = "15_U_E" Then 'Cod. Campo según Tipo Cabecera
                If CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString <> "" Then
                    CargaComboLínea(oForm, 1, CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Value.ToString)
                Else
                    objGlobal.SBOApp.MessageBox("No ha seleccionado la estructura de ficheros." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione Cód. de los Parámetros de Campos de SAP.")
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - No ha seleccionado la estructura de ficheros." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione Cód. de los Parámetros de Campos de SAP.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    CType(oForm.Items.Item("15_U_E").Specific, SAPbouiCOM.EditText).Active = True
                End If
            End If
            EventHandler_LOST_FOCUS = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function CargaComboLínea(ByRef oForm As SAPbouiCOM.Form, ByVal iLinea As Integer, ByVal sCode As String) As Boolean

        CargaComboLínea = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sTipo As String = ""
        Dim sTabla As String = ""
        Try
            oForm.Freeze(True)
            sTipo = CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_1").Cells.Item(iLinea).Specific, SAPbouiCOM.ComboBox).Selected.Value
            Select Case sTipo
                Case "C" : sTabla = "@EXO_CSAPC"
                Case "L" : sTabla = "@EXO_CSAPL"
                Case Else
                    objGlobal.SBOApp.MessageBox("Ha ocurrido un error inesperado en el campo ""Tipo Campo"".")
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Ha ocurrido un error inesperado en el campo ""Tipo Campo"".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
            End Select
            sSQL = "Select ""U_EXO_COD"",""U_EXO_DES"" FROM """ & sTabla & """ where ""Code""='" & sCode & "' Order by ""LineId"" "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(CType(oForm.Items.Item("0_U_G").Specific, SAPbouiCOM.Matrix).Columns.Item("C_0_2").Cells.Item(iLinea).Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            CargaComboLínea = True

        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaComboSerie(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboSerie = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sTipo As String = ""
        Try
            sSQL = "Select ""Series"",""SeriesName"" FROM ""NNM1"" WHERE ""ObjectCode""=2 and ""DocSubType""='C' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbSCLI").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = "Select ""Series"",""SeriesName"" FROM ""NNM1"" WHERE ""ObjectCode""=2 and ""DocSubType""='S' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cbSPRO").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            CargaComboSerie = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaComboViaPago(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboViaPago = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sTipo As String = ""
        Try
            sSQL = "Select ""PayMethCod"",""Descript"" FROM ""OPYM"" WHERE ""Type""='I' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_VIAPV").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = "Select ""PayMethCod"",""Descript"" FROM ""OPYM"" WHERE ""Type""='O' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_VPC").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            CargaComboViaPago = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function CargaComboCPago(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboCPago = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sTipo As String = ""
        Try
            sSQL = "Select ""GroupNum"",""PymntGroup"" FROM ""OCTG""  "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_CPV").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            sSQL = "Select ""GroupNum"",""PymntGroup"" FROM ""OCTG""  "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.funcionesUI.cargaCombo(CType(oForm.Items.Item("cb_CPC").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

            CargaComboCPago = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class

