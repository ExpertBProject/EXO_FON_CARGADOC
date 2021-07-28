Imports System.Xml
Imports SAPbouiCOM
Imports OfficeOpenXml
Imports EPPlus
Imports System.IO
Public Class EXO_CIC
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef general As EXO_Generales.EXO_General, ByRef actualizar As Boolean)
        MyBase.New(general, actualizar)
        ' cargamenu()       
    End Sub
    Private Sub cargamenu()
        Dim Path As String = ""
        Dim menuXML As String = objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_MENU.xml")
        SboApp.LoadBatchActions(menuXML)
        Dim res As String = SboApp.GetLastBatchResults

        If SboApp.Menus.Exists("EXO-MnCDoc") = True Then
            Path = objGlobal.conexionSAP.path & "\02.Menus"
            If Path <> "" Then
                If IO.File.Exists(Path & "\MnCDOC.png") = True Then
                    SboApp.Menus.Item("EXO-MnCDoc").Image = Path & "\MnCDOC.png"
                End If
            End If
        End If
    End Sub
    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function
    Public Overrides Function SBOApp_MenuEvent(ByRef infoEvento As EXO_Generales.EXO_MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCIC"
                        'Cargamos pantalla de gestión.
                        If CargarFormCDOC() = False Then
                            Exit Function
                        End If
                End Select
            End If

            Return MyBase.SBOApp_MenuEvent(infoEvento)

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Public Function CargarFormCDOC() As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim Path As String = ""
        Dim oFP As SAPbouiCOM.FormCreationParams = Nothing
        Dim EXO_Xml As New EXO_Generales.EXO_XML(objGlobal.conexionSAP.refCompañia, objGlobal.conexionSAP.refSBOApp)

        CargarFormCDOC = False

        Try
            Path = objGlobal.conexionSAP.pathPantallas
            If Path = "" Then
                Return False
            End If

            oFP = CType(SboApp.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams), SAPbouiCOM.FormCreationParams)
            oFP.XmlData = objGlobal.conexionSAP.leerEmbebido(Me.GetType(), "EXO_CIC.srf")

            Try
                oForm = SboApp.Forms.AddEx(oFP)
            Catch ex As Exception
                If ex.Message.StartsWith("Form - already exists") = True Then
                    SboApp.StatusBar.SetText("El formulario ya está abierto.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Function
                ElseIf ex.Message.StartsWith("Se produjo un error interno") = True Then 'Falta de autorización
                    Exit Function
                End If
            End Try
            CargaComboFormato(oForm)
            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Item.Enabled = False
            If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues.Count > 1 Then
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("CARGAIC", BoSearchKey.psk_ByValue)
            Else
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("--", BoSearchKey.psk_ByValue)
            End If
            CargarFormCDOC = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Visible = True
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CIC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CIC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                    If EventHandler_Matrix_Link_Press_Before(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CIC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                    If EventHandler_Choose_FromList_After(infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "EXO_CIC"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If

            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_Choose_FromList_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = SboApp.Forms.Item(pVal.FormUID)

        EventHandler_Choose_FromList_After = False

        Try

            If pVal.ItemUID = "grd_DOC" AndAlso pVal.ChooseFromListUID = "CFL_0" Then
                Dim oCFLEvento As EXO_Generales.EXO_infoItemEvent = Nothing
                Dim oDataTable As EXO_Generales.EXO_infoItemEvent.EXO_SeleccionadosCHFL = Nothing
                oCFLEvento = CType(pVal, EXO_Generales.EXO_infoItemEvent)

                oDataTable = oCFLEvento.SelectedObjects
                If oDataTable IsNot Nothing Then
                    Try
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Comercial", pVal.Row, oDataTable.GetValue("SlpName", 0).ToString)

                    Catch ex As Exception
                        oForm.DataSources.DataTables.Item("DT_DOC").SetValue("Comercial", pVal.Row, oDataTable.GetValue("SlpName", 0).ToString)
                    End Try
                End If
            End If

            EventHandler_Choose_FromList_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            oForm.Freeze(False)
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_Matrix_Link_Press_Before(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sModo As String = ""
        EventHandler_Matrix_Link_Press_Before = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)

            If pVal.ItemUID = "grd_DOC" Then
                If pVal.ColUID = "DocEntry" Then
                    sModo = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Modo", pVal.Row).ToString
                    If sModo = "F" Then
                        CType(CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry"), SAPbouiCOM.EditTextColumn).LinkedObjectType = CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).DataTable.GetValue("Tipo", pVal.Row).ToString
                    Else
                        CType(CType(oForm.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item("DocEntry"), SAPbouiCOM.EditTextColumn).LinkedObjectType = 112
                    End If
                End If
            End If
            EventHandler_Matrix_Link_Press_Before = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_FORM_VISIBLE(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        EventHandler_FORM_VISIBLE = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)
            If oForm.Visible = True Then
                oForm.Items.Item("btn_Carga").Enabled = False
            End If

            EventHandler_FORM_VISIBLE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim sTipoArchivo As String = ""
        Dim sArchivoOrigen As String = ""
        Dim sArchivo As String = objGlobal.conexionSAP.pathHistorico & "\DOC_CARGADOS\" & objGlobal.conexionSAP.SBOApp.Company.DatabaseName & "\IC\"
        Dim sNomFICH As String = ""
        EventHandler_ItemPressed_After = False

        Try
            oForm = SboApp.Forms.Item(pVal.FormUID)
            'Comprobamos que exista el directorio y sino, lo creamos
            If System.IO.Directory.Exists(sArchivo) = False Then
                System.IO.Directory.CreateDirectory(sArchivo)
            End If
            Select Case pVal.ItemUID
                Case "btn_Carga"
                    If SboApp.MessageBox("¿Está seguro que quiere tratar el fichero seleccionado?", 1, "Sí", "No") = 1 Then
                        sArchivoOrigen = CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value
                        sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                        sArchivo = sArchivo & sNomFICH
                        'Hacemos copia de seguridad para tratarlo
                        Copia_Seguridad(sArchivoOrigen, sArchivo)
                        oForm.Items.Item("btn_Carga").Enabled = False
                        SboApp.StatusBar.SetText("Creando/Actualizando IC ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oForm.Freeze(True)
                        'Ahora abrimos el fichero para tratarlo
                        TratarFichero(sArchivo, CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm)
                        oForm.Freeze(False)
                        SboApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        SboApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                        oForm.Items.Item("btn_Carga").Enabled = True
                    Else
                        SboApp.StatusBar.SetText("Se ha cancelado el proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    End If
                Case "btn_Fich"
                    'Cargar Fichero para leer
                    If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString <> "--" Then
                        If CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString = "XML" Then
                            sTipoArchivo = "XML|*.xml"
                        Else
                            sSQL = "Select ""U_EXO_TEXP"" FROM ""@EXO_CFCNF""  WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                Select Case oRs.Fields.Item("U_EXO_TEXP").Value.ToString
                                    Case "1" : sTipoArchivo = "Ficheros CSV|*.csv|Texto|*.txt"
                                    Case "2" : sTipoArchivo = "Libro de Excel|*.xlsx|Excel 97-2003|*.xls"
                                    Case Else
                                        SboApp.StatusBar.SetText("(EXO) - Error inesperado. No ha encontrado el tipo de fichero a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oForm.Items.Item("btn_Carga").Enabled = False
                                        Exit Function
                                End Select
                            End If
                        End If
                        'Tenemos que controlar que es cliente o web
                        If objGlobal.conexionSAP.SBOApp.ClientType = SAPbouiCOM.BoClientType.ct_Browser Then
                            sArchivoOrigen = objGlobal.conexionSAP.SBOApp.GetFileFromBrowser() 'Modificar
                        Else
                            'Controlar el tipo de fichero que vamos a abrir según campo de formato
                            sArchivoOrigen = objGlobal.Functions.OpenDialogFiles("Abrir archivo como", sTipoArchivo)
                        End If

                        If Len(sArchivoOrigen) = 0 Then
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = ""
                            SboApp.MessageBox("Debe indicar un archivo a importar.")
                            SboApp.StatusBar.SetText("(EXO) - Debe indicar un archivo a importar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                            oForm.Items.Item("btn_Carga").Enabled = False
                            Exit Function
                        Else
                            CType(oForm.Items.Item("txt_Fich").Specific, SAPbouiCOM.EditText).Value = sArchivoOrigen
                            'sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                            'sArchivo = sArchivo & sNomFICH
                            ''Hacemos copia de seguridad para tratarlo
                            'Copia_Seguridad(sArchivoOrigen, sArchivo)
                            ''Ahora abrimos el fichero para tratarlo
                            'TratarFichero(sArchivo, CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm)
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    Else
                        SboApp.MessageBox("No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.")
                        SboApp.StatusBar.SetText("(EXO) - No ha seleccionado el formato a importar." & ChrW(10) & ChrW(13) & " Antes de continuar Seleccione un formato de los que se ha creado en la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Active = True
                    End If
            End Select

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oColumnTxt, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oColumnChk, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function

    Private Sub TratarFichero_Excel(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
#Region "Variables"
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsDir As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRCampos As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCampo As String = ""

        Dim sClienteColumna As String = "" : Dim sTIC As String = "" : Dim sCodigoSAP As String = ""
        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCliNomComercial As String = "" : Dim sCodCliente As String = "" : Dim sBusinessPartnerPais As String = ""
        Dim sBPStreet As String = "" : Dim sBPZipCode As String = "" : Dim sBPCity As String = "" : Dim sBPCounty As String = ""
        Dim sLicTradNum As String = "" : Dim sPhone1 As String = "" : Dim sFax As String = "" : Dim sMovil As String = "" : Dim sMoneda As String = ""
        Dim sImpuesto As String = ""

        Dim sPais As String = ""
        Dim sViaPago As String = "" : Dim sCondPago As String = ""
        Dim sCondicion As String = ""

        Dim sExiste As String = ""
        Dim sCodCampos As String = "" ':Dim iLinea As Integer = 0 

        Dim pck As ExcelPackage = Nothing
        Dim iLin As Integer = 0
        Dim oOCRD As SAPbobsCOM.BusinessPartners = CType(Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
        Dim sExisteIC As String = ""
#End Region
        Try
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Dim excel As New FileInfo(sArchivo)
                pck = New ExcelPackage(excel)
                Dim workbook = pck.Workbook
                Dim worksheet = workbook.Worksheets.First()
                sSQL = "SELECT ""U_EXO_FEXCEL"",""U_EXO_CSAP"",""U_EXO_TDOC"" FROM ""@EXO_CFCNF"" WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    iLin = oRs.Fields.Item("U_EXO_FEXCEL").Value

                    sCodCampos = oRs.Fields.Item("U_EXO_CSAP").Value
                    sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPC"" WHERE ""Code""='" & sCodCampos & "'"
                    oRCampos.DoQuery(sSQL)
                    If oRCampos.RecordCount > 0 Then
#Region "Matrix de Cabecera"
                        Dim sCamposC(oRCampos.RecordCount, 3) As String
                        For I = 1 To oRCampos.RecordCount
                            sCamposC(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                            sCamposC(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                            sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='C' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                            sCampo = EXO_GLOBALES.GetValueDB(Company, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
                            If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
                                sMensaje &= "Por favor, Revise la parametrización."
                                SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                SboApp.MessageBox(sMensaje)
                                Exit Sub
                            End If
                            sCamposC(I, 3) = sCampo
                            Select Case oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                Case "CardCode", "ADDID" : sClienteColumna = sCampo
                            End Select
                            oRCampos.MoveNext()
                        Next
#End Region
                        Do
                            If sCliente <> worksheet.Cells(sClienteColumna & iLin).Text Then
                                'Grabamos la cabecera
                                For C = 1 To sCamposC.GetUpperBound(0)
                                    Select Case sCamposC(C, 1)
                                        Case "CardType"
#Region "CardType"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTIC = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    If worksheet.Cells("A" & iLin).Text <> "" Then
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        SboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    Else
                                                        Exit Do
                                                    End If
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sTIC = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "CardCode"
#Region "CardCode"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    'SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "CardName"
#Region "CardName"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "CardFName"
#Region "CardFName"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliNomComercial = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCliNomComercial = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "AddID"
#Region "AddID"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "LicTradNum"
#Region "LicTradNum"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sLicTradNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sLicTradNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "Phone1"
#Region "Phone1"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sPhone1 = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sPhone1 = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "Fax"
#Region "Fax"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFax = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sFax = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "Cellular"
#Region "Cellular"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sMovil = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sMovil = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "Currency"
#Region "Currency"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "ECVatGroup"
#Region "ECVatGroup"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sImpuesto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sImpuesto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "Country"
#Region "Country"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBusinessPartnerPais = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBusinessPartnerPais = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "Street"
#Region "Street"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPStreet = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPStreet = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "ZipCode"
#Region "ZipCode"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPZipCode = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPZipCode = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "City"
#Region "City"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPCity = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPCity = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "County"
#Region "County"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPCounty = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sBPCounty = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                        Case "County"
#Region "Country"
                                            If sCamposC(C, 2) = "Y" Then
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sPais = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                Else
                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                    sMensaje &= "Por favor, Revise el documento Excel."
                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    SboApp.MessageBox(sMensaje)
                                                    Exit Sub
                                                End If
                                            Else
                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                    sPais = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                End If
                                            End If
#End Region
                                    End Select
                                Next
#Region "IC"
                                'Buscaremos si existe el interlocutor                              
                                'Para ello, componemos el NIF
                                If IsNumeric(Left(sLicTradNum, 2)) Or IsNumeric(Mid(sLicTradNum, 2, 2)) Then
                                    Select Case sPais.ToUpper
                                        Case "", "ESP" : sPais = "ES"
                                        Case Else : sPais = "ES"
                                    End Select
                                    sLicTradNum = sPais & sLicTradNum.ToUpper
                                End If
                                If sTIC = "" Then
                                    If Left(sCliente, 3) = "430" Then
                                        sTIC = "C"
                                    ElseIf Left(sCliente, 3) = "400" Then
                                        sTIC = "S"
                                    End If
                                End If
                                sExisteIC = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OCRD""", """CardCode""", """LicTradNum""='" & sLicTradNum & "' and ""CardType""='" & sTIC & "' ")
                                oOCRD = CType(Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
                                If sExisteIC = "" Then
#Region "Creamos"
                                    'Componemos el Codigo para saber si es cliente o Proveedor
                                    'If sTIC = "" Then
                                    '    If Left(sCliente, 3) = "430" Then
                                    '        sTIC = "C"
                                    '    ElseIf Left(sCliente, 3) = "400" Then
                                    '        sTIC = "S"
                                    '    End If
                                    'End If
                                    Dim sGrupo As String = "" : Dim sSerie As String = ""
                                    If sTIC <> "" Then
                                        Select Case sTIC
                                            Case "C"
                                                oOCRD.CardType = SAPbobsCOM.BoCardTypes.cCustomer
                                                If sArchivo.Contains("BCN") = True Then
                                                    sCliente = "CT" & Right(sCodCliente, 5)
                                                Else
                                                    sCliente = "CT1" & Right(sCodCliente, 4)
                                                End If

                                                sGrupo = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OCRG""", """GroupCode""", """GroupName""='Clientes'")
                                                sSerie = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_SERIEV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S"
                                                oOCRD.CardType = SAPbobsCOM.BoCardTypes.cSupplier
                                                If sArchivo.Contains("BCN") = True Then
                                                    sCliente = "P" & Right(sCodCliente, 5)
                                                Else
                                                    sCliente = "P1" & Right(sCodCliente, 4)
                                                End If

                                                sGrupo = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OCRG""", """GroupCode""", """GroupName""='Acreedores'")
                                                If sGrupo = "" Then
                                                    sGrupo = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OCRG""", """GroupCode""", """GroupName""='Proveedores'")
                                                End If
                                                sSerie = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_SERIEC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    End If
                                    oOCRD.Series = sSerie
                                    If sSerie = "1" Or sSerie = "2" Then
                                        oOCRD.CardCode = sCliente
                                    End If
                                    oOCRD.CardName = sCliNombre
                                    oOCRD.CardForeignName = sCliNomComercial
                                    oOCRD.GroupCode = sGrupo
                                    If sCodCliente <> "" Then : oOCRD.AdditionalID = sCodCliente : End If
                                    If sBusinessPartnerPais <> "" Then
                                        Select Case sBusinessPartnerPais.ToUpper
                                            Case "", "ESP" : sBusinessPartnerPais = "ES"
                                            Case Else : sBusinessPartnerPais = "ES"
                                        End Select
                                        oOCRD.Country = sBusinessPartnerPais
                                    End If
                                    If sLicTradNum <> "" Then : oOCRD.FederalTaxID = sLicTradNum : End If
                                    'Condición de pago
                                    If sCondPago <> "" Then
                                        sCondPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OCTG""", """PymntGroup""", """PymntGroup""='" & sCondPago & "'")
                                    End If
                                    If sCondPago = "" Then
                                        Select Case sTIC
                                            Case "C" : sCondPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CPV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S" : sCondPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CPC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    End If
                                    oOCRD.PayTermsGrpCode = sCondPago

                                    'Via de pago
                                    'Comprobamos que exista, si no existe cogemos el valor por defecto del configurador
                                    If sViaPago <> "" Then
                                        sViaPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OPYM""", """PayMethCod""", """PayMethCod""='" & sViaPago & "'")
                                        ' SboApp.StatusBar.SetText("(EXO) - La vía de pago indicada no existe - " & sViaPago & " -  Se tomará el valor por defecto de la configuración.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
                                    If sViaPago = "" Then
                                        Select Case sTIC
                                            Case "C" : sViaPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_VIAPV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S" : sViaPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_VIAPC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    End If
                                    oOCRD.BPPaymentMethods.PaymentMethodCode = sViaPago
                                    oOCRD.BPPaymentMethods.Add()
                                    oOCRD.PeymentMethodCode = sViaPago
                                    'oOCRD.Equalization = SAPbobsCOM.BoYesNoEnum.tNO ' Rerención
                                    If sImpuesto = "" Then
                                        'Indicamos el valor por defecto de la pantalla de configuración
                                        Dim sBPImpuestoCod As String = ""
                                        Select Case sTIC
                                            Case "C" : sBPImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S" : sBPImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAP""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                        If sBPImpuestoCod <> "" Then
                                            oOCRD.VatGroup = sBPImpuestoCod
                                        Else
                                            Dim sMensaje As String = ""
                                            sMensaje = "No se ha indicado el impuesto para el interlocutor en la ventana de configuración."
                                            SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            SboApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                    Else
                                        'Tenemos que hacer la conversión
                                        Select Case sTIC
                                            Case "C"
                                                If sImpuesto = "1" Then
                                                    oOCRD.VatGroup = "R3" '21%
                                                ElseIf sImpuesto = "2" Then
                                                    oOCRD.VatGroup = "R2" '10%
                                                ElseIf sImpuesto = "3" Then
                                                    oOCRD.VatGroup = "R1" '4%
                                                End If
                                            Case "S"
                                                If sImpuesto = "1" Then
                                                    oOCRD.VatGroup = "S3" '21%
                                                ElseIf sImpuesto = "2" Then
                                                    oOCRD.VatGroup = "S2" '10%
                                                ElseIf sImpuesto = "3" Then
                                                    oOCRD.VatGroup = "S1" '4%
                                                End If
                                        End Select
                                    End If
                                    If sPhone1 <> "" Then : oOCRD.Phone1 = sPhone1 : End If
                                    If sFax <> "" Then : oOCRD.Fax = sFax : End If
                                    If sMovil <> "" Then : oOCRD.Cellular = sMovil : End If
                                    'SboApp.StatusBar.SetText("(EXO) - Currency - " & sMoneda & ". Cliente " & sCliente & " Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    If sMoneda <> "" And Len(Trim(sMoneda)) <= 3 Then : oOCRD.Currency = sMoneda : Else : oOCRD.Currency = "EUR" : End If
                                    'Dirección
                                    oOCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                                    oOCRD.Addresses.AddressName = "Facturación"
                                    oOCRD.Addresses.Street = sBPStreet : oOCRD.Addresses.City = sBPCity : oOCRD.Addresses.ZipCode = sBPZipCode : oOCRD.Addresses.County = sBPCounty : oOCRD.Addresses.Country = sBusinessPartnerPais
                                    oOCRD.Addresses.Add()
                                    If oOCRD.Add() <> 0 Then
                                        Throw New Exception(Company.GetLastErrorCode & " / No se puede crear el interlocutor -" & sCliNombre & " - " & Company.GetLastErrorDescription)
                                    Else
                                        sCodigoSAP = Company.GetNewObjectKey()
                                        SboApp.StatusBar.SetText("(EXO) - Se ha creado el interlocutor - " & sCodigoSAP & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
#End Region
                                Else
#Region "Actualizamos"
                                    oOCRD.GetByKey(sExisteIC)
                                    'Componemos el Codigo para saber si es cliente o Proveedor
                                    'If sTIC = "" Then
                                    '    If Left(sCliente, 3) = "430" Then
                                    '        sTIC = "C"
                                    '    ElseIf Left(sCliente, 3) = "400" Then
                                    '        sTIC = "S"
                                    '    End If
                                    'End If
                                    If sCliNombre <> "" Then : oOCRD.CardName = sCliNombre : End If
                                    If sCliNomComercial <> "" Then : oOCRD.CardForeignName = sCliNomComercial : End If
                                    If sCodCliente <> "" Then : oOCRD.AdditionalID = sCodCliente : End If
                                    If sBusinessPartnerPais <> "" Then
                                        Select Case sBusinessPartnerPais.ToUpper
                                            Case "", "ESP" : sBusinessPartnerPais = "ES"
                                            Case Else : sBusinessPartnerPais = "ES"
                                        End Select
                                        oOCRD.Country = sBusinessPartnerPais
                                    End If
                                    If sLicTradNum <> "" Then : oOCRD.FederalTaxID = sLicTradNum : End If
                                    'Condición de pago
                                    If sCondPago <> "" Then
                                        sCondPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OCTG""", """PymntGroup""", """PymntGroup""='" & sCondPago & "'")
                                    End If
                                    If sCondPago = "" Then
                                        Select Case sTIC
                                            Case "C" : sCondPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CPV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S" : sCondPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CPC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    End If
                                    oOCRD.PayTermsGrpCode = sCondPago

                                    'Via de pago
                                    'Comprobamos que exista, si no existe cogemos el valor por defecto del configurador
                                    If sViaPago <> "" Then
                                        sViaPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OPYM""", """PayMethCod""", """PayMethCod""='" & sViaPago & "'")
                                        'SboApp.StatusBar.SetText("(EXO) - La vía de pago indicada no existe - " & sViaPago & " -  Se tomará el valor por defecto de la configuración.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
                                    If sViaPago = "" Then
                                        Select Case sTIC
                                            Case "C" : sViaPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_VIAPV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S" : sViaPago = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_VIAPC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    End If
                                    oOCRD.BPPaymentMethods.PaymentMethodCode = sViaPago
                                    oOCRD.BPPaymentMethods.Add()
                                    oOCRD.PeymentMethodCode = sViaPago
                                    'oOCRD.Equalization = SAPbobsCOM.BoYesNoEnum.tNO ' Retención
                                    If sImpuesto = "" Then
                                        'Indicamos el valor por defecto de la pantalla de configuración
                                        Dim sBPImpuestoCod As String = ""
                                        Select Case sTIC
                                            Case "C" : sBPImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "S" : sBPImpuestoCod = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAP""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                        If sBPImpuestoCod <> "" Then
                                            oOCRD.VatGroup = sBPImpuestoCod
                                        Else
                                            Dim sMensaje As String = ""
                                            sMensaje = "No se ha indicado el impuesto para el interlocutor en la ventana de configuración."
                                            SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            SboApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                    Else
                                        'Tenemos que hacer la conversión
                                        Select Case sTIC
                                            Case "C"
                                                If sImpuesto = "1" Then
                                                    oOCRD.VatGroup = "R3" '21%
                                                ElseIf sImpuesto = "2" Then
                                                    oOCRD.VatGroup = "R2" '10%
                                                ElseIf sImpuesto = "3" Then
                                                    oOCRD.VatGroup = "R1" '4%
                                                End If
                                            Case "S"
                                                If sImpuesto = "1" Then
                                                    oOCRD.VatGroup = "S3" '21%
                                                ElseIf sImpuesto = "2" Then
                                                    oOCRD.VatGroup = "S2" '10%
                                                ElseIf sImpuesto = "3" Then
                                                    oOCRD.VatGroup = "S1" '4%
                                                End If
                                        End Select
                                    End If
                                    If sPhone1 <> "" Then : oOCRD.Phone1 = sPhone1 : End If
                                    If sFax <> "" Then : oOCRD.Fax = sFax : End If
                                    If sMovil <> "" Then : oOCRD.Cellular = sMovil : End If
                                    If sMoneda <> "" And Len(Trim(sMoneda)) <= 3 Then : oOCRD.Currency = sMoneda : Else : oOCRD.Currency = "EUR" : End If
                                    'Buscamos el codigo de afcturación
                                    Dim sDirFactu As String = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """CRD1""", """Address""", """AdresType""='B' and ""CardCode""='" & sExisteIC & "'")
                                    If oOCRD.Update() <> 0 Then
                                        Throw New Exception(Company.GetLastErrorCode & " / No se puede modificar el interlocutor -" & sCliNombre & " - " & Company.GetLastErrorDescription)
                                    Else
                                        sCodigoSAP = Company.GetNewObjectKey()
                                        'Dirección
                                        'Actualizamos la dircción por update
                                        sSQL = "UPDATE " & objGlobal.conexionSAP.compañia.CompanyDB.ToString & ".""CRD1"" "
                                        sSQL &= " SET ""Street""='" & sBPStreet & " ',"
                                        sSQL &= " ""City""='" & sBPCity & "', "
                                        sSQL &= " ""ZipCode""='" & sBPZipCode & "', "
                                        sSQL &= " ""County""='" & sBPCounty & "', "
                                        sSQL &= " ""Country""='" & sBusinessPartnerPais & "' "
                                        sSQL &= " WHERE ""AdresType""='B' and ""CardCode""='" & sCodigoSAP & "' and ""Address""='" & sDirFactu & "' "
                                        oRsDir.DoQuery(sSQL)
                                        SboApp.StatusBar.SetText("(EXO) - Se ha modificado el interlocutor - " & sCodigoSAP & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    End If
#End Region
                                End If
                            End If
#End Region
                            iLin += 1
                        Loop While sCliente <> ""
                    End If
                Else
                    SboApp.StatusBar.SetText("(EXO) - Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    SboApp.MessageBox("Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.")
                End If
            Else
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("No se ha encontrado el fichero excel a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            pck.Dispose()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsDir, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD, Object))
        End Try
    End Sub

    Private Sub TratarFichero(ByVal sArchivo As String, ByVal sTipoArchivo As String, ByRef oForm As SAPbouiCOM.Form)
        Dim myStream As StreamReader = Nothing
        Dim Reader As XmlTextReader = New XmlTextReader(myStream)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sExiste As String = "" ' Para comprobar si existen los datos
        Dim sDelimitador As String = ""
        Try
            sSQL = "Select ""U_EXO_STXT"" FROM ""@EXO_CFCNF""  WHERE ""Code""='" & sTipoArchivo & "'"
            sDelimitador = objGlobal.SQL.sqlStringB1(sSQL)

            sSQL = "Select ""U_EXO_TEXP"" FROM ""@EXO_CFCNF""  WHERE ""Code""='" & sTipoArchivo & "'"
            sTipoArchivo = objGlobal.SQL.sqlStringB1(sSQL)

            Select Case sTipoArchivo
                Case "1"
#Region "TXT|CSV"
                    EXO_GLOBALES.TratarFichero_TXT(sArchivo, sDelimitador, oForm, Company, SboApp, objGlobal)
#End Region
                Case "2"
#Region "EXCEL"
                    TratarFichero_Excel(sArchivo, oForm)
#End Region
                Case Else
                    SboApp.StatusBar.SetText("(EXO) -El tipo de fichero a importar no está contemplado. Avise a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    SboApp.MessageBox("El tipo de fichero a importar no está contemplado. Avise a su Administrador.")
                    Exit Sub
            End Select
            SboApp.StatusBar.SetText("(EXO) - Se ha leido correctamente el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            oForm.Freeze(True)
            'SboApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'SboApp.MessageBox("Se ha leido correctamente el fichero. Fin del proceso")
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            myStream = Nothing
            Reader.Close()
            Reader = Nothing
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub

    Private Sub Copia_Seguridad(ByVal sArchivoOrigen As String, ByVal sArchivo As String)
        'Comprobamos el directorio de copia que exista
        Dim sPath As String = ""
        sPath = IO.Path.GetDirectoryName(sArchivo)
        If IO.Directory.Exists(sPath) = False Then
            IO.Directory.CreateDirectory(sPath)
        End If
        If IO.File.Exists(sArchivo) = True Then
            IO.File.Delete(sArchivo)
        End If
        'Subimos el archivo
        SboApp.StatusBar.SetText("(EXO) - Comienza la Copia de seguridad del fichero - " & sArchivoOrigen & " -.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        If objGlobal.conexionSAP.SBOApp.ClientType = BoClientType.ct_Browser Then
            Dim fs As FileStream = New FileStream(sArchivoOrigen, FileMode.Open, FileAccess.Read)
            Dim b(CInt(fs.Length() - 1)) As Byte
            fs.Read(b, 0, b.Length)
            fs.Close()
            Dim fs2 As New System.IO.FileStream(sArchivo, IO.FileMode.Create, IO.FileAccess.Write)
            fs2.Write(b, 0, b.Length)
            fs2.Close()
        Else
            My.Computer.FileSystem.CopyFile(sArchivoOrigen, sArchivo)
        End If
        SboApp.StatusBar.SetText("(EXO) - Copia de Seguridad realizada Correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
    Private Function CargaComboFormato(ByRef oForm As SAPbouiCOM.Form) As Boolean

        CargaComboFormato = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            If objGlobal.conexionSAP.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sSQL = "(Select '--' as ""Code"",' ' as ""Name"" FROM ""DUMMY"") "
                sSQL &= " UNION ALL "
                sSQL &= " (Select ""Code"",""Name"" FROM ""@EXO_CFCNF"" Order by ""Name"") "
            Else
                sSQL = "SELECT * FROM ( "
                sSQL &= " (Select ""Code"",""Name"" FROM ""@EXO_CFCNF"") "
                sSQL &= " UNION ALL "
                sSQL &= "(Select '--' as ""Code"",' ' as ""Name"" ) "
                sSQL &= " ) T  Order by ""Name"" "
            End If

            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                objGlobal.conexionSAP.refSBOApp.cargaCombo(CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).ValidValues, sSQL)
            Else
                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Por favor, antes de continuar, revise la parametrización.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
            CargaComboFormato = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
    Private Function ComprobarDOC(ByRef oForm As SAPbouiCOM.Form, ByVal sFra As String) As Boolean
        Dim bLineasSel As Boolean = False

        ComprobarDOC = False

        Try
            For i As Integer = 0 To oForm.DataSources.DataTables.Item(sFra).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sFra).GetValue("Sel", i).ToString = "Y" Then
                    bLineasSel = True
                    Exit For
                End If
            Next

            If bLineasSel = False Then
                SboApp.MessageBox("Debe seleccionar al menos una línea.")
                Exit Function
            End If

            ComprobarDOC = bLineasSel

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
