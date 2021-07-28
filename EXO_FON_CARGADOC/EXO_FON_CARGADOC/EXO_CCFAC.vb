Imports System.Xml
Imports SAPbouiCOM
Imports OfficeOpenXml
Imports System.IO
Public Class EXO_CCFAC
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef general As EXO_Generales.EXO_General, ByRef actualizar As Boolean)
        MyBase.New(general, actualizar)
        cargamenu()
        If actualizar Then
            cargaCampos()
        End If
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
    Private Sub cargaCampos()
        If objGlobal.conexionSAP.esAdministrador Then
            Dim oXML As String = ""
            Dim udoObj As EXO_Generales.EXO_UDO = Nothing
            'EXO_TMPDOC Cabecera
            oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOC.xml")
            objGlobal.conexionSAP.LoadBDFromXML(oXML)
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validado: UT_EXO_TMPDOC", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'EXO_TMPDOC Líneas
            oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCL.xml")
            objGlobal.conexionSAP.LoadBDFromXML(oXML)
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validado: UT_EXO_TMPDOCL", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'EXO_TMPDOC Líneas Lotes
            oXML = objGlobal.Functions.leerEmbebido(Me.GetType(), "UT_EXO_TMPDOCLT.xml")
            objGlobal.conexionSAP.LoadBDFromXML(oXML)
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Validado: UT_EXO_TMPDOCALT", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
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
                    Case "EXO-MnCFac"
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
            oFP.XmlData = objGlobal.conexionSAP.leerEmbebido(Me.GetType(), "EXO_CCFAC.srf")

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
                CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Select("FACCOMPRAS", BoSearchKey.psk_ByValue)
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
                        Case "EXO_CCFAC"
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
                        Case "EXO_CCFAC"
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
                        Case "EXO_CCFAC"
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
                        Case "EXO_CCFAC"
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
        Dim sArchivo As String = objGlobal.conexionSAP.pathHistorico & "\DOC_CARGADOS\" & objGlobal.conexionSAP.SBOApp.Company.DatabaseName & "\COMPRAS\FACTURAS\"
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
                    If SboApp.MessageBox("¿Está seguro que quiere generar los Documentos seleccionados?", 1, "Sí", "No") = 1 Then
                        If ComprobarDOC(oForm, "DT_DOC") = True Then
                            oForm.Items.Item("btn_Carga").Enabled = False
                            'Generamos facturas
                            SboApp.StatusBar.SetText("Creando Documentos ... Espere por favor.", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                            oForm.Freeze(True)
                            If EXO_GLOBALES.CrearDocumentos(oForm, "DT_DOC", "FACTURA", Company, SboApp, objGlobal) = False Then
                                Exit Function
                            End If
                            oForm.Freeze(False)
                            SboApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                            SboApp.MessageBox("Fin del Proceso" & ChrW(10) & ChrW(13) & "Por favor, revise el Log para ver las operaciones realizadas.")
                            oForm.Items.Item("btn_Carga").Enabled = True
                        End If
                    End If
                Case "btn_Fich"
                    Limpiar_Grid(oForm)
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
                            sNomFICH = IO.Path.GetFileName(sArchivoOrigen)
                            sArchivo = sArchivo & sNomFICH
                            'Hacemos copia de seguridad para tratarlo
                            Copia_Seguridad(sArchivoOrigen, sArchivo)
                            'Ahora abrimos el fichero para tratarlo
                            TratarFichero(sArchivo, CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, oForm)
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
    Private Sub Limpiar_Grid(ByRef oForm As SAPbouiCOM.Form)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oForm.Freeze(True)
            'Limpiamos grid
            'Borrar tablas temporales por usuario activo
            sSQL = "DELETE FROM ""@EXO_TMPDOC"" where ""U_EXO_USR""='" & objGlobal.usuario & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPDOCL"" where ""U_EXO_USR""='" & objGlobal.usuario & "'  "
            oRs.DoQuery(sSQL)
            sSQL = "DELETE FROM ""@EXO_TMPDOCLT"" where ""U_EXO_USR""='" & objGlobal.usuario & "'  "
            oRs.DoQuery(sSQL)
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""Code"",""U_EXO_MODO"" as ""Modo"", '     ' as ""Estado"",""U_EXO_TIPOF"" As ""Tipo"",'      ' as ""DocEntry"", ""U_EXO_Serie"" as ""Serie"",""U_EXO_DOCNUM"" as ""Nº Documento"","
            sSQL &= " ""U_EXO_REF"" as ""Referencia"", ""U_EXO_MONEDA"" as ""Moneda"", ""U_EXO_COMER"" as ""Comercial"", ""U_EXO_CLISAP"" as ""Interlocutor SAP"", ""U_EXO_ADDID"" as ""Interlocutor Ext."", "
            sSQL &= " ""U_EXO_FCONT"" as ""F. Contable"", ""U_EXO_FDOC"" as ""F. Documento"", ""U_EXO_FVTO"" as ""F. Vto"", ""U_EXO_TDTO"" as ""T. Dto."", ""U_EXO_DTO"" as ""Dto."",  "
            sSQL &= " ""U_EXO_CPAGO"" as ""Vía Pago"", ""U_EXO_GROUPNUM"" as ""Cond. Pago"", ""U_EXO_COMENT"" as ""Comentario"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPDOC"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.usuario & "' "
            sSQL &= " ORDER BY ""U_EXO_FDOC"", ""U_EXO_MODO"", ""U_EXO_TIPOF"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    '    Private Sub TratarFichero_Excel(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form)
    '        Dim sSQL As String = ""
    '        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        Dim oRCampos As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        Dim sCampo As String = ""

    '        Dim iDoc As Integer = 0 'Contador de Cabecera de documentos
    '        Dim sTFac As String = "" : Dim sTFacColumna As String = "" : Dim sTipoLineas As String = "" : Dim sTDoc As String = ""
    '        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCodCliente As String = "" : Dim sClienteColumna As String = "" : Dim sCodClienteColumna As String = ""
    '        Dim sSerie As String = "" : Dim sDocNum As String = "" : Dim sManual As String = "" : Dim sSerieColumna As String = "" : Dim sDocNumColumna As String = ""
    '        Dim sNumAtCard As String = "" : Dim sNumAtCardColumna As String = ""
    '        Dim sMoneda As String = "" : Dim sMonedaColumna As String = ""
    '        Dim sEmpleado As String = ""
    '        Dim sFContable As String = "" : Dim sFDocumento As String = "" : Dim sFVto As String = "" : Dim sFDocumentoColumna As String = ""
    '        Dim sTipoDto As String = "" : Dim sDto As String = ""
    '        Dim sPeyMethod As String = "" : Dim sCondPago As String = ""
    '        Dim sDirFac As String = "" : Dim sDirEnv As String = ""
    '        Dim sComent As String = "" : Dim sComentCab As String = "" : Dim sComentPie As String = ""
    '        Dim sCondicion As String = ""

    '        Dim sExiste As String = ""
    '        Dim iLinea As Integer = 0 : Dim sCodCampos As String = ""

    '        Dim pck As ExcelPackage = Nothing
    '        Dim iLin As Integer = 0
    '        Try
    '            ' miramos si existe el fichero y cargamos
    '            If File.Exists(sArchivo) Then
    '                Dim excel As New FileInfo(sArchivo)
    '                pck = New ExcelPackage(excel)
    '                Dim workbook = pck.Workbook
    '                Dim worksheet = workbook.Worksheets.First()
    '                sSQL = "SELECT ""U_EXO_FEXCEL"",""U_EXO_CSAP"",""U_EXO_TDOC"" FROM ""@EXO_CFCNF"" WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
    '                oRs.DoQuery(sSQL)
    '                If oRs.RecordCount > 0 Then
    '                    iLin = oRs.Fields.Item("U_EXO_FEXCEL").Value
    '                    sTDoc = oRs.Fields.Item("U_EXO_TDOC").Value
    '                    If sTDoc = "1" Then
    '                        sTDoc = "B"
    '                    Else
    '                        sTDoc = "F"
    '                    End If
    '                    sCodCampos = oRs.Fields.Item("U_EXO_CSAP").Value
    '                    sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPC"" WHERE ""Code""='" & sCodCampos & "'"
    '                    oRCampos.DoQuery(sSQL)
    '                    If oRCampos.RecordCount > 0 Then
    '#Region "Matrix de Cabecera"
    '                        Dim sCamposC(oRCampos.RecordCount, 3) As String
    '                        For I = 1 To oRCampos.RecordCount
    '                            sCamposC(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
    '                            sCamposC(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
    '                            sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='C' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
    '                            sCampo = EXO_GLOBALES.GetValueDB(Company, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
    '                            If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
    '                                ' Si es obligatorio y no se ha indicado para leer hay que dar un error
    '                                Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
    '                                sMensaje &= "Por favor, Revise la parametrización."
    '                                SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                SboApp.MessageBox(sMensaje)
    '                                Exit Sub
    '                            End If
    '                            sCamposC(I, 3) = sCampo
    '                            Select Case oRCampos.Fields.Item("U_EXO_COD").Value.ToString
    '                                Case "ObjType" : sTFacColumna = sCampo
    '                                Case "CardCode" : sClienteColumna = sCampo
    '                                Case "ADDID" : sCodClienteColumna = sCampo
    '                                Case "Series" : sSerieColumna = sCampo
    '                                Case "DocNum" : sDocNumColumna = sCampo
    '                                Case "NumAtCard" : sNumAtCardColumna = sCampo
    '                                Case "DocCurrency" : sMonedaColumna = sCampo
    '                                Case "TaxDate" : sFDocumentoColumna = sCampo
    '                            End Select
    '                            oRCampos.MoveNext()
    '                        Next
    '#End Region
    '                        Do
    '#Region "Cabecera"
    '                            If sTFac <> worksheet.Cells(sTFacColumna & iLin).Text Or sCliente <> worksheet.Cells(sClienteColumna & iLin).Text Or sCodCliente <> worksheet.Cells(sCodClienteColumna & iLin).Text _
    '                                Or sSerie <> worksheet.Cells(sSerieColumna & iLin).Text Or sDocNum <> worksheet.Cells(sDocNumColumna & iLin).Text Or sNumAtCard <> worksheet.Cells(sNumAtCardColumna & iLin).Text _
    '                                Or sMoneda <> worksheet.Cells(sMonedaColumna & iLin).Text Or sFDocumento <> worksheet.Cells(sFDocumentoColumna & iLin).Text Then
    '                                'Grabamos la cabecera
    '                                For C = 1 To sCamposC.GetUpperBound(0)
    '                                    Select Case sCamposC(C, 1)
    '                                        Case "ObjType"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    If worksheet.Cells("A" & iLin).Text <> "" Then
    '                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                        sMensaje &= "Por favor, Revise el documento Excel."
    '                                                        SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                        SboApp.MessageBox(sMensaje)
    '                                                        Exit Sub
    '                                                    Else
    '                                                        Exit Do
    '                                                    End If
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "DocType"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sTipoLineas = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sTipoLineas = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "CardCode"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "CardName"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "ADDID"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "NumAtCard"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sNumAtCard = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sNumAtCard = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "EXO_Manual"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sManual = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sManual = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "Series"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sSerie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sSerie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "DocNum"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDocNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDocNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "DocCurrency"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "SlpCode"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sEmpleado = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sEmpleado = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "DocDate"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sFContable = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sFContable = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "TaxDate"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sFDocumento = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sFDocumento = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "DocDueDate"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sFVto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sFVto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "EXO_TDTO"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sTipoDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sTipoDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "EXO_DTO"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "PeyMethod"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sPeyMethod = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sPeyMethod = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "GroupNum"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCondPago = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sCondPago = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "PayToCode"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDirFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDirFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "ShipToCode"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDirEnv = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sDirEnv = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "Comments"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sComent = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sComent = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "OpeningRemarks"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sComentCab = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sComentCab = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                        Case "ClosingRemarks"
    '                                            If sCamposC(C, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sComentPie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
    '                                                    sComentPie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
    '                                                End If
    '                                            End If
    '                                    End Select
    '                                Next
    '                                'Grabamos la cabecera
    '                                iLinea = 0
    '                                'Insertar en la tabla temporal la cabecera
    '                                If sTFac <> "" Then
    '                                    iDoc += 1
    '                                    sSQL = "insert into ""@EXO_TMPFAC"" values('" & iDoc.ToString & "','" & iDoc.ToString & "'," & iDoc.ToString & ",'N','',0," & objGlobal.conexionSAP.compañia.UserSignature
    '                                    sSQL &= ",'','',0,'',0,'','" & objGlobal.usuario & "',"
    '                                    sSQL &= "'" & sTDoc & "','" & sDocNum & "','" & sTFac & "','" & sManual & "','" & sSerie & "','" & sNumAtCard & "','" & sMoneda & "','','" & sEmpleado & "',"
    '                                    sSQL &= "'" & sCliente & "','" & sCodCliente & "','" & sFContable & "','" & sFDocumento & "','" & sFVto & "','" & sTipoDto & "',"
    '                                    sSQL &= EXO_GLOBALES.DblNumberToText(objGlobal, sDto.ToString) & ",'" & sPeyMethod & "','" & sDirFac & "','" & sDirEnv & "','" & sComent.Replace("'", "") & "','"
    '                                    sSQL &= sComentCab.Replace("'", "") & "','" & sComentPie.Replace("'", "") & "','" & sCondPago & "') "
    '                                    oRs.DoQuery(sSQL)
    '                                End If
    '                            End If
    '#End Region
    '                            'Ahora tratamos la línea
    '#Region "Líneas"
    '                            Dim sCuenta As String = "" : Dim sArt As String = "" : Dim sArtDes As String = ""
    '                            Dim sCantidad As String = "0.00" : Dim sprecio As String = "0.00" : Dim sDtoLin As String = "0.00" : Dim sTotalServicios As String = "0.00"
    '                            Dim sTextoAmpliado As String = "" : Dim sLinImpuestoCod As String = "" : Dim sLinRetCodigo As String = ""
    '                            sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPL"" WHERE ""Code""='" & sCodCampos & "'"
    '                            oRCampos.DoQuery(sSQL)
    '                            If oRCampos.RecordCount > 0 Then
    '#Region "Matrix de Líneas"
    '                                Dim sCamposL(oRCampos.RecordCount, 3) As String
    '                                For I = 1 To oRCampos.RecordCount
    '                                    sCamposL(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
    '                                    sCamposL(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
    '                                    sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='L' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
    '                                    sCampo = EXO_GLOBALES.GetValueDB(Company, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
    '                                    If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
    '                                        ' Si es obligatorio y no se ha indicado para leer hay que dar un error
    '                                        Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
    '                                        sMensaje &= "Por favor, Revise la parametrización."
    '                                        SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                        SboApp.MessageBox(sMensaje)
    '                                        Exit Sub
    '                                    End If
    '                                    sCamposL(I, 3) = sCampo
    '                                    oRCampos.MoveNext()
    '                                Next
    '#End Region
    '                                For L = 1 To sCamposL.GetUpperBound(0)
    '                                    Select Case sCamposL(L, 1)
    '                                        Case "AcctCode"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sCuenta = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sCuenta = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sCuenta = ""
    '                                                End If
    '                                            End If
    '                                        Case "ItemCode"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sArt = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sArt = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sArt = ""
    '                                                End If
    '                                            End If
    '                                        Case "Dscription"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sArtDes = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sArtDes = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sArtDes = ""
    '                                                End If
    '                                            End If
    '                                        Case "Quantity"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sCantidad = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sCantidad = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sCantidad = "0.00"
    '                                                End If
    '                                            End If
    '                                        Case "UnitPrice"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sprecio = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sprecio = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sprecio = "0.00"
    '                                                End If
    '                                            End If
    '                                        Case "DiscPrcnt"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sDtoLin = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sDtoLin = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sDtoLin = "0.00"
    '                                                End If
    '                                            End If
    '                                        Case "EXO_IMPSRV"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sTotalServicios = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sTotalServicios = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sTotalServicios = "0.00"
    '                                                End If
    '                                            End If
    '                                        Case "EXO_TextoLin"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sTextoAmpliado = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sTextoAmpliado = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sTextoAmpliado = ""
    '                                                End If
    '                                            End If
    '                                        Case "EXO_IMP"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sLinImpuestoCod = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sLinImpuestoCod = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sLinImpuestoCod = ""
    '                                                End If
    '                                            End If
    '                                        Case "EXO_RET"
    '                                            If sCamposL(L, 2) = "Y" Then
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sLinRetCodigo = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
    '                                                    sMensaje &= "Por favor, Revise el documento Excel."
    '                                                    SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                                    SboApp.MessageBox(sMensaje)
    '                                                    Exit Sub
    '                                                End If
    '                                            Else
    '                                                If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
    '                                                    sLinRetCodigo = worksheet.Cells(sCamposL(L, 3) & iLin).Text
    '                                                Else
    '                                                    sLinRetCodigo = ""
    '                                                End If
    '                                            End If
    '                                    End Select
    '                                Next
    '#Region "Comprobar datos línea"
    '                                'Comprobamos que exista la cuenta
    '                                If sCuenta <> "" Then
    '                                    sExiste = ""
    '                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OACT""", """AcctCode""", """AcctCode""='" & sCuenta & "'")
    '                                    If sExiste = "" Then
    '                                        SboApp.StatusBar.SetText("(EXO) - La Cuenta contable SAP  - " & sCuenta & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                        SboApp.MessageBox("La Cuenta contable SAP - " & sCuenta & " - no existe.")
    '                                        Exit Sub
    '                                    End If
    '                                End If
    '                                'Comprobamos que exista el artículo
    '                                If sTipoLineas = "I" Then
    '                                    sExiste = ""
    '                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OITM""", """ItemCode""", """ItemCode"" like '" & sArt & "'")
    '                                    If sExiste = "" Then
    '                                        SboApp.StatusBar.SetText("(EXO) - El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                        SboApp.MessageBox("El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.")
    '                                        Exit Sub
    '                                    End If
    '                                ElseIf sTipoLineas = "S" Then
    '                                    If sCuenta = "" Then
    '                                        ' No puede estar la cuenta vacía si es de tipo servicio
    '                                        Dim sMensaje As String = " La cuenta en la línea del servicio no puede estar vacía. Por favor, Revise los datos."
    '                                        SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                        SboApp.MessageBox(sMensaje)
    '                                        Exit Sub
    '                                    End If
    '                                End If
    '                                'Comprobamos que exista el impuesto si está relleno
    '                                If sLinImpuestoCod <> "" Then
    '                                    sExiste = ""
    '                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """OVTG""", """Code""", """Code""='" & sLinImpuestoCod & "'")
    '                                    If sExiste = "" Then
    '                                        SboApp.StatusBar.SetText("(EXO) - El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                        SboApp.MessageBox("El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.")
    '                                        Exit Sub
    '                                    End If
    '                                End If
    '                                'Comprobamos que exista la retención si está relleno
    '                                If sLinRetCodigo <> "" Then
    '                                    sExiste = EXO_GLOBALES.GetValueDB(objGlobal.conexionSAP.compañia, """CRD4""", """WTCode""", """CardCode""='" & sCliente & "' and ""WTCode""='" & sLinRetCodigo & "'")
    '                                    If sExiste = "" Then
    '                                        SboApp.StatusBar.SetText("(EXO) - El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                                        SboApp.MessageBox("El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.")
    '                                        Exit Sub
    '                                    End If
    '                                End If
    '#End Region
    '                                'Grabamos la línea
    '                                sSQL = "insert into ""@EXO_TMPFAL"" values('" & iDoc.ToString & "','" & iLinea & "','',0,'" & objGlobal.usuario & "',"
    '                                sSQL &= "'" & sCuenta & "','" & sArt & "','" & sArtDes & "'," & EXO_GLOBALES.DblNumberToText(objGlobal, sCantidad.ToString) & ","
    '                                sSQL &= EXO_GLOBALES.DblNumberToText(objGlobal, sprecio.ToString) & "," & EXO_GLOBALES.DblNumberToText(objGlobal, sDtoLin.ToString)
    '                                sSQL &= "," & EXO_GLOBALES.DblNumberToText(objGlobal, sTotalServicios.ToString) & ",'" & sLinImpuestoCod & "','" & sLinRetCodigo & "','"
    '                                sSQL &= sTextoAmpliado & "','" & sTipoLineas & "' ) "
    '                                oRs.DoQuery(sSQL)
    '                                iLin += 1 : iLinea += 1
    '                            End If
    '#End Region
    '                        Loop While sTFac <> ""
    '                    End If


    '                Else
    '                    SboApp.StatusBar.SetText("(EXO) - Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                    SboApp.MessageBox("Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.")
    '                End If
    '            Else
    '                objGlobal.conexionSAP.SBOApp.StatusBar.SetText("No se ha encontrado el fichero excel a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                Exit Sub
    '            End If


    '        Catch exCOM As System.Runtime.InteropServices.COMException
    '            Throw exCOM
    '        Catch ex As Exception
    '            Throw ex
    '        Finally
    '            pck.Dispose()
    '            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
    '            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
    '        End Try
    '    End Sub

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
                    EXO_GLOBALES.TratarFichero_Excel(sArchivo, oForm, Company, SboApp, objGlobal)
#End Region
                Case Else
                    SboApp.StatusBar.SetText("(EXO) -El tipo de fichero a importar no está contemplado. Avise a su Administrador.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    SboApp.MessageBox("El tipo de fichero a importar no está contemplado. Avise a su Administrador.")
                    Exit Sub
            End Select
            SboApp.StatusBar.SetText("(EXO) - Se ha leido correctamente el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

#Region "cargar Grid con los datos leidos"
            'Ahora cargamos el Grid con los datos guardados
            objGlobal.conexionSAP.SBOApp.StatusBar.SetText("Cargando Documentos en pantalla ... Espere por favor", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            sSQL = "SELECT 'Y' as ""Sel"",""Code"",""U_EXO_MODO"" as ""Modo"", '     ' as ""Estado"",""U_EXO_TIPOF"" As ""Tipo"",'      ' as ""DocEntry"",""U_EXO_Serie"" as ""Serie"",""U_EXO_DOCNUM"" as ""Nº Documento"","
            sSQL &= " ""U_EXO_REF"" as ""Referencia"", ""U_EXO_MONEDA"" as ""Moneda"", ""U_EXO_COMER"" as ""Comercial"", ""U_EXO_CLISAP"" as ""Interlocutor SAP"", ""U_EXO_ADDID"" as ""Interlocutor Ext."", "
            sSQL &= " ""U_EXO_FCONT"" as ""F. Contable"", ""U_EXO_FDOC"" as ""F. Documento"", ""U_EXO_FVTO"" as ""F. Vto"", ""U_EXO_TDTO"" as ""T. Dto."", ""U_EXO_DTO"" as ""Dto."",  "
            sSQL &= " ""U_EXO_CPAGO"" as ""Vía Pago"", ""U_EXO_GROUPNUM"" as ""Cond. Pago"", ""U_EXO_COMENT"" as ""Comentario"", "
            sSQL &= " CAST('' as varchar(254)) as ""Descripción Estado"" "
            sSQL &= " From ""@EXO_TMPDOC"" "
            sSQL &= " WHERE ""U_EXO_USR""='" & objGlobal.usuario & "' "
            sSQL &= " ORDER BY ""U_EXO_FDOC"",""U_EXO_MODO"", ""U_EXO_TIPOF"" "
            'Cargamos grid
            oForm.DataSources.DataTables.Item("DT_DOC").ExecuteQuery(sSQL)
            FormateaGrid(oForm)
#End Region
            oForm.Freeze(True)
            SboApp.StatusBar.SetText("Fin del proceso.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            SboApp.MessageBox("Se ha leido correctamente el fichero. Fin del proceso")
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
    Private Sub FormateaGrid(ByRef oform As SAPbouiCOM.Form)
        Dim oColumnTxt As SAPbouiCOM.EditTextColumn = Nothing
        Dim oColumnChk As SAPbouiCOM.CheckBoxColumn = Nothing
        Dim oColumnCb As SAPbouiCOM.ComboBoxColumn = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Try
            oform.Freeze(True)
            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            oColumnChk = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(0), SAPbouiCOM.CheckBoxColumn)
            oColumnChk.Editable = True
            oColumnChk.Width = 30

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(1).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(1), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.Width = 40

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oColumnCb = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(2), SAPbouiCOM.ComboBoxColumn)
            oColumnCb.ValidValues.Add("F", "Factura")
            oColumnCb.ValidValues.Add("B", "Borrador")
            oColumnCb.Editable = True
            oColumnCb.Width = 70
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(3), SAPbouiCOM.EditTextColumn)
            oColumnTxt.Editable = False
            oColumnTxt.Width = 50

            CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(4).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oColumnCb = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(4), SAPbouiCOM.ComboBoxColumn)
            oColumnCb.ValidValues.Add("13", "Factura de Ventas")
            oColumnCb.ValidValues.Add("14", "Abonos de Venta")
            oColumnCb.ValidValues.Add("18", "Factura de Compras")
            oColumnCb.ValidValues.Add("19", "Abono de Compras")
            oColumnCb.ValidValues.Add("22", "Pedido de Compras")
            oColumnCb.DisplayType = BoComboDisplayType.cdt_Description
            oColumnCb.Editable = False
            oColumnCb.Width = 100

            For i = 5 To 10
                If i <> 8 Then
                    CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                    oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                    If i <> 10 Then
                        oColumnTxt.Editable = False
                    End If
                End If
                If i = 5 Then
                    oColumnTxt.LinkedObjectType = "22"
                ElseIf i = 10 Then
                    'Comercial
                    oColumnTxt.ChooseFromListUID = "CFL_0"
                    oColumnTxt.ChooseFromListAlias = "SlpName"
                    oColumnTxt.Width = 150
                End If
            Next

            For i = 11 To 21
                CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oColumnTxt = CType(CType(oform.Items.Item("grd_DOC").Specific, SAPbouiCOM.Grid).Columns.Item(i), SAPbouiCOM.EditTextColumn)
                oColumnTxt.Editable = False
                Select Case i
                    Case 11, 12, 13, 14, 15 : oColumnTxt.Width = 70
                    Case 16, 17 : oColumnTxt.Width = 45
                    Case 21 : oColumnTxt.Width = 300
                End Select

                If i = 11 Then
                    oColumnTxt.LinkedObjectType = "2"
                End If
                Select Case i
                    Case 16, 17 : oColumnTxt.RightJustified = True
                End Select
            Next
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            oform.Freeze(False)
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
