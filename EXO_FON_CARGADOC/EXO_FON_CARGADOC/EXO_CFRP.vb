Imports System.Xml
Imports SAPbouiCOM

Public Class EXO_CFRP
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
            oXML = objGlobal.funciones.leerEmbebido(Me.GetType(), "UDO_EXO_CSAP.xml")
            objGlobal.refDi.comunes.LoadBDFromXML(oXML)
            objGlobal.SBOApp.StatusBar.SetText("Validado: UDO_EXO_CSAP", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            'Introducir los datos
            CargarDatos()
        End If
    End Sub
    Private Function CargarDatos() As Boolean
        CargarDatos = False
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsCCC As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sPeriodo As String = ""
        Dim sFPoder As String = ""
        Dim oDI_COM As EXO_DIAPI.EXO_UDOEntity = Nothing 'Instancia del UDO para Insertar datos
        Try
            oDI_COM = New EXO_DIAPI.EXO_UDOEntity(objGlobal.refDi.comunes, "EXO_CSAP") 'UDO de Campos de SAP
#Region "CAMPOSSAP"
            sSQL = "SELECT * FROM ""@EXO_CSAP"" WHERE ""Code""='CAMPOSSAP' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sCode As String = oRs.Fields.Item("Code").Value.ToString
                oDI_COM.GetByKey(sCode)
                'Comprobamos que existan campos en la tabla de la cabecera
                sSQL = "SELECT * FROM ""@EXO_CSAPC"" WHERE ""Code""='CAMPOSSAP' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Cabecera Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposCabecera(oDI_COM, "CAMPOSSAP")
                End If
                'Comprobamos que existan campos en las líneas
                sSQL = "SELECT * FROM ""@EXO_CSAPL"" WHERE ""Code""='CAMPOSSAP' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Líneas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposLíneas(oDI_COM, "CAMPOSSAP")
                End If
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tablas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "CAMPOSSAP"
                oDI_COM.SetValue("CodEntry") = "99"
                oDI_COM.SetValue("Name") = "Campos de SAP"
                CrearCamposCabecera(oDI_COM, "CAMPOSSAP")
                CrearCamposLíneas(oDI_COM, "CAMPOSSAP")
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir campos SAP. " & oDI_COM.GetLastError)
                End If
            End If
#End Region
#Region "CAMPOSSAPEXCEL"
            sSQL = "SELECT * FROM ""@EXO_CSAP"" WHERE ""Code""='CAMPOSSAPEXCEL' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sCode As String = oRs.Fields.Item("Code").Value.ToString
                oDI_COM.GetByKey(sCode)
                'Comprobamos que existan campos en la tabla de la cabecera
                sSQL = "SELECT * FROM ""@EXO_CSAPC"" WHERE ""Code""='CAMPOSSAPEXCEL' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Cabecera Campos SAP EXCEL...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposCabecera(oDI_COM, "CAMPOSSAPEXCEL")
                End If
                'Comprobamos que existan campos en las líneas
                sSQL = "SELECT * FROM ""@EXO_CSAPL"" WHERE ""Code""='CAMPOSSAPEXCEL' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Líneas Campos SAP EXCEL...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposLíneas(oDI_COM, "CAMPOSSAPEXCEL")
                End If
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tablas Campos SAP EXCEL...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "CAMPOSSAPEXCEL"
                oDI_COM.SetValue("CodEntry") = "98"
                oDI_COM.SetValue("Name") = "Campos de SAP EXCEL "
                CrearCamposCabecera(oDI_COM, "CAMPOSSAPEXCEL")
                CrearCamposLíneas(oDI_COM, "CAMPOSSAPEXCEL")
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir campos SAP EXCEL. " & oDI_COM.GetLastError)
                End If
            End If
#End Region
#Region "CAMPOSSAPFCSV"
            sSQL = "SELECT * FROM ""@EXO_CSAP"" WHERE ""Code""='CAMPOSSAPFCSV' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sCode As String = oRs.Fields.Item("Code").Value.ToString
                oDI_COM.GetByKey(sCode)
                'Comprobamos que existan campos en la tabla de la cabecera
                sSQL = "SELECT * FROM ""@EXO_CSAPC"" WHERE ""Code""='CAMPOSSAPFCSV' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Cabecera Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposCabecera(oDI_COM, "CAMPOSSAPFCSV")
                End If
                'Comprobamos que existan campos en las líneas
                sSQL = "SELECT * FROM ""@EXO_CSAPL"" WHERE ""Code""='CAMPOSSAPFCSV' "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount = 0 Then
                    objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tabla de Líneas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    CrearCamposLíneas(oDI_COM, "CAMPOSSAPFCSV")
                End If
                If oDI_COM.UDO_Update = False Then
                    Throw New Exception("(EXO) - " & oDI_COM.GetLastError)
                End If
            Else
                objGlobal.SBOApp.StatusBar.SetText("(EXO) - Rellenando Tablas Campos SAP...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                oDI_COM.GetNew()
                oDI_COM.SetValue("Code") = "CAMPOSSAPFCSV"
                oDI_COM.SetValue("CodEntry") = "97"
                oDI_COM.SetValue("Name") = "Campos de SAP para Fac. Ventas CSV"
                CrearCamposCabecera(oDI_COM, "CAMPOSSAPFCSV")
                CrearCamposLíneas(oDI_COM, "CAMPOSSAPFCSV")
                If oDI_COM.UDO_Add = False Then
                    Throw New Exception("(EXO) - Error al añadir campos SAP. " & oDI_COM.GetLastError)
                End If
            End If
#End Region

            objGlobal.SBOApp.StatusBar.SetText("(EXO) - Tablas Campos SAP cargadas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            CargarDatos = True

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDI_COM, Object))
        End Try
    End Function

    Private Function CrearCamposCabecera(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String)
        Try
            For i = 0 To 21
                oDI_COM.GetNewChild("EXO_CSAPC")
                Select Case i
                    Case 0
                        oDI_COM.SetValueChild("U_EXO_COD") = "ObjType"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Tipo Documento"
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 1
                        oDI_COM.SetValueChild("U_EXO_COD") = "CardCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Interlocutor"
                        If sCodigo = "CAMPOSSAPEXCEL" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 2
                        oDI_COM.SetValueChild("U_EXO_COD") = "CardName"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Nombre Interlocutor"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 3

                        oDI_COM.SetValueChild("U_EXO_COD") = "ADDID"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Externo Interlocutor"
                        If sCodigo = "CAMPOSSAPEXCEL" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 4
                        oDI_COM.SetValueChild("U_EXO_COD") = "NumAtCard"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Número de referencia"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 5
                        oDI_COM.SetValueChild("U_EXO_COD") = "Series"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Serie Factura"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 6
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocNum"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Nº de Documento"
                        If sCodigo = "CAMPOSSAPEXCEL" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 7
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocCurrency"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Moneda"
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 8
                        oDI_COM.SetValueChild("U_EXO_COD") = "SlpCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Empleado"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 9
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Contable"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                    Case 10
                        oDI_COM.SetValueChild("U_EXO_COD") = "TaxDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Documento"
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 11
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocDueDate"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Fecha Vencimiento"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 12
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TDTO"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Tipo Dto."
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                    Case 13
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_DTO"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descuento"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 14
                        oDI_COM.SetValueChild("U_EXO_COD") = "PeyMethod"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Vía de Pago"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 15
                        oDI_COM.SetValueChild("U_EXO_COD") = "GroupNum"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Condición de Pago"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 16
                        oDI_COM.SetValueChild("U_EXO_COD") = "PayToCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Dir. Facturación"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 17
                        oDI_COM.SetValueChild("U_EXO_COD") = "ShipToCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Dirección de entrega"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 18
                        oDI_COM.SetValueChild("U_EXO_COD") = "Comments"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Comentario"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 19
                        oDI_COM.SetValueChild("U_EXO_COD") = "OpeningRemarks"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Texto en Cabecera"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 20
                        oDI_COM.SetValueChild("U_EXO_COD") = "ClosingRemarks"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Texto en pie"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 21
                        oDI_COM.SetValueChild("U_EXO_COD") = "DocType" ' I --> Artículos o S --> Servicios
                        oDI_COM.SetValueChild("U_EXO_DES") = "Tipo de Doc."
                        If sCodigo = "CAMPOSSAP" Then
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        Else
                            oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                        End If
                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Private Function CrearCamposLíneas(ByRef oDI_COM As EXO_DIAPI.EXO_UDOEntity, ByVal sCodigo As String)
        Try
            For i = 0 To 15
                oDI_COM.GetNewChild("EXO_CSAPL")
                Select Case i
                    Case 0
                        oDI_COM.SetValueChild("U_EXO_COD") = "AcctCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Cta. Mayor"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 1
                        oDI_COM.SetValueChild("U_EXO_COD") = "ItemCode"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Artículo"

                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 2
                        oDI_COM.SetValueChild("U_EXO_COD") = "Dscription"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descripción Artículo"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 3
                        oDI_COM.SetValueChild("U_EXO_COD") = "Quantity"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Cantidad"

                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 4
                        oDI_COM.SetValueChild("U_EXO_COD") = "UnitPrice"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Precio Unidad"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 5
                        oDI_COM.SetValueChild("U_EXO_COD") = "DiscPrcnt"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Descuento %"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 6
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_IMPSRV"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Importe Servicio"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 7
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TextoLin"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Texto Ampliado de la línea"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 8
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_IMP"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Impuesto"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 9
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_RET"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Retención"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 10
                        oDI_COM.SetValueChild("U_EXO_COD") = "GrossBuyPr"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Precio Bruto"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 11
                        oDI_COM.SetValueChild("U_EXO_COD") = "EXO_REPARTO"
                        oDI_COM.SetValueChild("U_EXO_DES") = "Código Reparto"
                        oDI_COM.SetValueChild("U_EXO_OBL") = "N"
                    Case 12
                        If sCodigo = "CAMPOSSAPFCSV" Then
                            oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TransSantander"
                            oDI_COM.SetValueChild("U_EXO_DES") = "Id Transacción Santander"
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        End If
                    Case 13
                        If sCodigo = "CAMPOSSAPFCSV" Then
                            oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TransfY"
                            oDI_COM.SetValueChild("U_EXO_DES") = "Id Transacción fY"
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        End If
                    Case 14
                        If sCodigo = "CAMPOSSAPFCSV" Then
                            oDI_COM.SetValueChild("U_EXO_COD") = "EXO_TransMM"
                            oDI_COM.SetValueChild("U_EXO_DES") = "Id Transacción MM"
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        End If
                    Case 15
                        If sCodigo = "CAMPOSSAPFCSV" Then
                            oDI_COM.SetValueChild("U_EXO_COD") = "EXO_NUMMOVIL"
                            oDI_COM.SetValueChild("U_EXO_DES") = "Número de móvil"
                            oDI_COM.SetValueChild("U_EXO_OBL") = "Y"
                        End If
                End Select
            Next
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Public Overrides Function SBOApp_MenuEvent(infoEvento As MenuEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing

        Try
            If infoEvento.BeforeAction = True Then
            Else
                Select Case infoEvento.MenuUID
                    Case "EXO-MnCSAP"
                        'Cargamos UDO Campos SAP.
                        objGlobal.funcionesUI.cargaFormUdoBD("EXO_CSAP")
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
                        Case "UDO_FT_EXO_CSAP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CSAP"
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
                        Case "UDO_FT_EXO_CSAP"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE
                                    If EventHandler_FORM_VISIBLE(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select

                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "UDO_FT_EXO_CSAP"

                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

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
    Private Function EventHandler_FORM_VISIBLE(ByRef objGlobal As EXO_UIAPI.EXO_UIAPI, ByRef pVal As ItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        EventHandler_FORM_VISIBLE = False

        Try
            oForm = objGlobal.SBOApp.Forms.Item(pVal.FormUID)

            If oForm.Visible = True Then
                sSQL = "SELECT * FROM ""@EXO_CSAP"" "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    oForm.Mode = BoFormMode.fm_OK_MODE
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objGlobal.SBOApp.ActivateMenuItem("1290") ' Ir al primer registro
                    End If
                Else
                    oForm.Mode = BoFormMode.fm_ADD_MODE
                End If
            End If

            EventHandler_FORM_VISIBLE = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.Mostrar_Error(exCOM, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Catch ex As Exception
            objGlobal.Mostrar_Error(ex, EXO_UIAPI.EXO_UIAPI.EXO_TipoMensaje.Excepcion)
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
End Class

