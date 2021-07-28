Imports SAPbouiCOM
Imports SAPbobsCOM
Imports System.IO
Imports OfficeOpenXml

Public Class EXO_GLOBALES
#Region "Funciones formateos datos"
    Public Shared Function TextToDbl(ByRef oObjGlobal As EXO_Generales.EXO_General, ByVal sValor As String) As Double
        Dim cValor As Double = 0
        Dim sValorAux As String = "0"

        TextToDbl = 0

        Try
            sValorAux = sValor

            If oObjGlobal.conexionSAP.SBOApp.ClientType = BoClientType.ct_Desktop Then
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "." Then sValorAux = "0" & sValorAux

                    If oObjGlobal.conexionSAP.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.conexionSAP.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ""))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(".") > 0 AndAlso sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "").Replace(".", ","))
                        ElseIf sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", ","))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            Else
                If sValorAux <> "" Then
                    If Left(sValorAux, 1) = "," Then sValorAux = "0" & sValorAux

                    If oObjGlobal.conexionSAP.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.conexionSAP.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", ""))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    Else 'Decimales USA
                        If sValorAux.IndexOf(",") > 0 AndAlso sValorAux.IndexOf(".") > 0 Then
                            cValor = CDbl(sValorAux.Replace(".", "").Replace(",", "."))
                        ElseIf sValorAux.IndexOf(",") > 0 Then
                            cValor = CDbl(sValorAux.Replace(",", "."))
                        Else
                            cValor = CDbl(sValorAux)
                        End If
                    End If
                End If
            End If

            TextToDbl = cValor

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function DblNumberToText(ByRef oObjGlobal As EXO_Generales.EXO_General, ByVal sValor As String) As String
        Dim sNumberDouble As String = "0"

        DblNumberToText = "0"

        Try
            If sValor <> "" Then
                If oObjGlobal.conexionSAP.OADM.separadorMillarB1 = "." AndAlso oObjGlobal.conexionSAP.OADM.separadorDecimalB1 = "," Then 'Decimales ES
                    sNumberDouble = sValor
                Else 'Decimales USA
                    sNumberDouble = sValor.Replace(",", ".")
                End If
            End If

            DblNumberToText = sNumberDouble


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Shared Function FormateaString(ByVal dato As Object, ByVal tam As Integer) As String
        Dim retorno As String = String.Empty

        If dato IsNot Nothing Then
            retorno = dato.ToString
        End If

        If retorno.Length > tam Then
            retorno = retorno.Substring(0, tam)
        End If

        Return retorno.PadRight(tam, CChar(" "))
    End Function
    Public Shared Function FormateaNumero(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
    Public Shared Function FormateaNumeroSinPunto(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        End If
        Return retorno
    End Function
    Public Shared Function FormateaNumeroconSigno(ByVal dato As String, ByVal posiciones As Integer, ByVal decimales As Integer, ByVal obligatorio As Boolean) As String
        Dim retorno As String = String.Empty
        Dim val As Decimal
        Dim totalNum As Integer = posiciones
        Dim format As String = ""
        Dim bEsNegativo As Boolean = False
        If Left(dato, 1) = "-" Then
            dato = dato.Replace("-", "")
            bEsNegativo = True
            posiciones = posiciones - 1
            totalNum = posiciones
        End If
        Decimal.TryParse(dato.Replace(".", ","), val)
        If val = 0 AndAlso Not obligatorio Then
            retorno = New String(CChar(" "), posiciones)
        Else
            If decimales <= 0 Then
            Else
                format = "0"
                format = "0." & New String(CChar("0"), decimales)
            End If
            retorno = val.ToString(format).Replace(",", ".")
            retorno = retorno.Replace(".", "")

            If retorno.Length > totalNum Then
                retorno = retorno.Substring(retorno.Length - totalNum)
            End If
            retorno = retorno.PadLeft(totalNum, CChar("0"))
        End If
        If bEsNegativo = True Then
            retorno = "N" & retorno
        Else
            retorno = " " & retorno
        End If
        Return retorno
    End Function
#End Region

#Region "Objetos SAP"
    'Public Shared Function GenerarPagoCTA(ByRef oObjGlobal As EXO_Generales.EXO_General, oCompany As SAPbobsCOM.Company, ByVal SCodLiq As String, ByVal sModelo As String) As Boolean
    '    Dim OVPM As SAPbobsCOM.Payments = Nothing
    '    Dim sPago As String = ""
    '    Dim sError As String = ""
    '    Dim oRs As SAPbobsCOM.Recordset = Nothing
    '    Dim oRsCuenta As SAPbobsCOM.Recordset = Nothing
    '    Dim sSQL As String = ""
    '    Dim dFechaCobro As Date
    '    Dim sFecha As String = ""
    '    Dim dblImporte As Double = 0
    '    Dim sCuentaBanco As String = ""
    '    Dim sCCCExt As String = ""
    '    Dim sMoneda As String = "EUR"
    '    Dim sConcepto As String = ""
    '    GenerarPagoCTA = False
    '    Try
    '        oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        oRsCuenta = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
    '        Select Case sModelo
    '            Case "303"
    '                sSQL = "SELECT ""Code"", ""U_EXO_NRC"", ""U_EXO_FINI"", ""U_EXO_FFIN"", ""U_EXO_CIF"", ""U_EXO_NOMBRE""," _
    '                & " ""U_EXO_BIC"", ""U_EXO_CCC1"",""U_EXO_CCC2"",""U_EXO_CCC3"",""U_EXO_CCC4"",""U_EXO_CCC5""," _
    '                & " ""U_EXO_73"" ""DEVOLVER"",""U_EXO_73I"" ""INGRESAR"" " _
    '                & " FROM ""@EXO_M303"" WHERE ""Code"" ='" & SCodLiq & "' "
    '                oRs.DoQuery(sSQL)
    '            Case "111"
    '                sSQL = "SELECT ""Code"", ""U_EXO_NRC"",""U_EXO_FINI"", ""U_EXO_FFIN"", ""U_EXO_NIF"", ""U_EXO_NOMBRE""" _
    '                & " ,""U_EXO_BIC"", ""U_EXO_CCC1"",""U_EXO_CCC2"",""U_EXO_CCC3"",""U_EXO_CCC4"",""U_EXO_CCC5"" " _
    '                & " ,Case    When ""U_EXO_30"" >  0 Then  ""U_EXO_30"" End  As ""INGRESAR"" " _
    '                & ",Case    When ""U_EXO_30"" <  0 Then  ""U_EXO_30"" End  As ""DEVOLVER"" " _
    '                & "  FROM ""@EXO_M111""  WHERE ""Code"" ='" & SCodLiq & "'"
    '                oRs.DoQuery(sSQL)
    '            Case "115"
    '                sSQL = "SELECT ""Code"", ""U_EXO_NRC"" , ""U_EXO_FINI"", ""U_EXO_FFIN"", ""U_EXO_NIF"", ""U_EXO_NOMBRE"" " _
    '                & " ,""U_EXO_BIC"", ""U_EXO_CCC1"",""U_EXO_CCC2"",""U_EXO_CCC3"",""U_EXO_CCC4"",""U_EXO_CCC5"",COALESCE(""U_EXO_05"",0) ""U_EXO_05"" " _
    '                & " ,Case    When ""U_EXO_05"" >  0 Then  COALESCE(""U_EXO_05"",0) End  As ""INGRESAR"" " _
    '                & " ,Case    When ""U_EXO_05"" <  0 Then  COALESCE(""U_EXO_05"",0) End  As ""DEVOLVER"" " _
    '                & " FROM ""@EXO_M115"" WHERE ""Code"" ='" & SCodLiq & "'"
    '                oRs.DoQuery(sSQL)
    '            Case "123"
    '                sSQL = "SELECT ""Code"", ""U_EXO_NRC"",""U_EXO_FINI"", ""U_EXO_FFIN"", ""U_EXO_NIF"", ""U_EXO_NOMBRE"" " _
    '                & " ,""U_EXO_BIC"", ""U_EXO_CCC1"",""U_EXO_CCC2"",""U_EXO_CCC3"",""U_EXO_CCC4"",""U_EXO_CCC5"",COALESCE(""U_EXO_08"",0) ""U_EXO_08"" " _
    '                & " ,Case    When ""U_EXO_08"" >  0 Then  ""U_EXO_08"" End  As ""INGRESAR"" " _
    '                & " ,Case    When ""U_EXO_08"" <  0 Then  ""U_EXO_08"" End  As ""DEVOLVER"" " _
    '                & " FROM ""@EXO_M123"" WHERE ""Code"" ='" & SCodLiq & "'"
    '                oRs.DoQuery(sSQL)
    '        End Select
    '        If oRs.RecordCount > 0 Then

    '            OVPM = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments), SAPbobsCOM.Payments)
    '            sFecha = Date.Now.ToShortDateString.ToString
    '            sConcepto = "Modelo " & sModelo & " " & oRs.Fields.Item("Code").Value.ToString
    '            'dFechaCobro = CDate(Date.ParseExact(sFecha, "yyyyMMdd", Nothing).ToString("yyyy/MM/dd"))

    '            sSQL = "Select   ""IBAN"", ""BankCode"", ""Branch"",""ControlKey"",""Account"", ""GLAccount"", ""GLIntriAct"" FROM """ & oCompany.CompanyDB & """.""DSC1"" WHERE LEFT(""IBAN"",4) =  '" & oRs.Fields.Item("U_EXO_CCC1").Value.ToString & "' and  ""BankCode"" = '" & oRs.Fields.Item("U_EXO_CCC2").Value.ToString & "' and ""Branch"" = '" & oRs.Fields.Item("U_EXO_CCC3").Value.ToString & "' and ""ControlKey"" ='" & oRs.Fields.Item("U_EXO_CCC4").Value.ToString & "' and ""Account"" = '" & oRs.Fields.Item("U_EXO_CCC5").Value.ToString & "'   "
    '            oRsCuenta.DoQuery(sSQL)
    '            If oRsCuenta.RecordCount > 0 Then
    '                sCuentaBanco = oRsCuenta.Fields.Item("GLIntriAct").Value.ToString
    '                If sCuentaBanco = "" Then
    '                    sCuentaBanco = oRsCuenta.Fields.Item("GLAccount").Value.ToString
    '                End If

    '                sCCCExt = "475000"
    '            End If

    '            dblImporte = CDbl(oRs.Fields.Item("INGRESAR").Value.ToString)
    '            OVPM.DocType = SAPbobsCOM.BoRcptTypes.rAccount

    '            'OVPM.TransferSum = CDbl(oRs.Fields.Item("INGRESAR").Value.ToString.Replace(".", ","))
    '            'OVPM.TransferDate = CDate(sFecha)
    '            'OVPM.TransferReference = sConcepto
    '            Select Case sMoneda
    '                Case "E"
    '                    sMoneda = "EUR"
    '            End Select
    '            OVPM.DocCurrency = sMoneda
    '            OVPM.DueDate = CDate(sFecha)
    '            OVPM.DocDate = CDate(oRs.Fields.Item("U_EXO_FFIN").Value.ToString)
    '            OVPM.TransferAccount = sCuentaBanco
    '            OVPM.TransferReference = oRs.Fields.Item("Code").Value.ToString
    '            OVPM.TransferDate = CDate(dFechaCobro)
    '            OVPM.TransferRealAmount = CDbl(oRs.Fields.Item("INGRESAR").Value.ToString.Replace(".", ","))
    '            OVPM.TransferSum = CDbl(oRs.Fields.Item("INGRESAR").Value.ToString.Replace(".", ","))
    '            OVPM.AccountPayments.AccountCode = sCCCExt
    '            OVPM.AccountPayments.SumPaid = CDbl(oRs.Fields.Item("INGRESAR").Value.ToString.Replace(".", ","))
    '            OVPM.AccountPayments.Decription = sConcepto
    '            OVPM.AccountPayments.Add()

    '            If OVPM.Add() = 0 Then
    '                sSQL = "SELECT ""DocNum"" FROM """ & oCompany.CompanyDB & """.""OVPM"" WHERE ""DocEntry""=" & oCompany.GetNewObjectKey()
    '                oRs.DoQuery(sSQL)
    '                If oRs.RecordCount > 0 Then
    '                    sPago = oRs.Fields.Item("DocNum").Value.ToString
    '                End If
    '                EXO_GLOBALES.Actualizar_ImpE(oObjGlobal, oCompany, sModelo, SCodLiq, "3", "", CInt(oCompany.GetNewObjectKey()))

    '                'grabar
    '                oObjGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Se ha generado el pago " & sPago & " con el concepto " & sConcepto, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '                GenerarPagoCTA = True
    '            Else
    '                sError = oCompany.GetLastErrorDescription
    '                oObjGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - No puede generar el pago. " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '                GenerarPagoCTA = False
    '            End If
    '        End If
    '    Catch ex As Exception
    '        sError = ex.Message
    '        oObjGlobal.conexionSAP.SBOApp.StatusBar.SetText(sError, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        Throw ex
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCuenta, Object))
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(OVPM, Object))
    '    End Try

    'End Function
#End Region
#Region "SQL"
    Public Shared Function GetValueDB(oCompany As SAPbobsCOM.Company, ByRef sTable As String, ByRef sField As String, ByRef sCondition As String) As String
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)

        Try
            If sCondition = "" Then
                sSQL = "Select " & sField & " FROM " & sTable
            Else
                sSQL = "Select " & sField & " FROM " & sTable & " WHERE " & sCondition
            End If
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                sField = sField.Replace("""", "")
                GetValueDB = CType(oRs.Fields.Item(sField).Value, String)
            Else
                GetValueDB = ""
            End If

        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
#Region "Funciones Bancarias"
    Public Shared Function Validar_CCC(ByVal numeroCuenta As String) As Boolean
        Try
            ' Primero compruebo que la longitud del parámetro
            ' sea de 20 caracteres.
            '
            If (numeroCuenta.Length <> 20) Then Return False

            ' Extraigo el dígito de control.
            '
            Dim dc As String = numeroCuenta.Substring(8, 2)

            ' Del número de cuenta, elimino el dígito de control.
            '
            numeroCuenta = numeroCuenta.Remove(8, 2)

            ' Obtengo el dígito de control verdadero.
            '
            Dim dcTemp As String = GetDCCuentaBancaria(numeroCuenta)

            ' Devuelvo el resultado.
            '
            If dc = dcTemp Then Return True
        Catch ex As Exception
            Throw ex
        End Try
        Return False
    End Function
    Public Shared Function GetDCCuentaBancaria(ByVal numeroCuenta As String) As String

        Try

            '
            If (numeroCuenta.Length <> 18) Then
                Return Nothing
            Else
                Dim ch As Char
                For Each ch In numeroCuenta
                    If (Not Char.IsNumber(ch)) Then Return Nothing
                Next
            End If

            Dim cociente1, cociente2, resto1, resto2 As Integer
            Dim sucursal, cuenta, dc1, dc2 As String
            Dim suma1, suma2, n As Integer

            ' Matriz que contiene los pesos utilizados en el
            ' algoritmo de cálculo de los dos dígitos de control.
            '
            Dim pesos() As Integer = {6, 3, 7, 9, 10, 5, 8, 4, 2, 1}

            sucursal = numeroCuenta.Substring(0, 8)
            cuenta = numeroCuenta.Substring(8, 10)

            ' Obtengo el primer dígito de control que verificará
            ' los códigos de Entidad y Oficina.
            '
            For n = 7 To 0 Step -1
                suma1 = suma1 + Convert.ToInt32(sucursal.Substring(n, 1)) * pesos(7 - n)
            Next

            ' Calculamos el cociente de dividir el resultado
            ' de la suma entre 11.
            '
            cociente1 = suma1 \ 11 ' Nos da un resultado entero.

            ' Calculo el resto de la división. Para ello,
            ' en lugar de utilizar el operador Mod, utilizo
            ' la fórmula para obtener el resto de la división.
            '
            resto1 = suma1 - (11 * cociente1)

            dc1 = (11 - resto1).ToString

            Select Case dc1
                Case "11"
                    dc1 = "0"
                Case "10"
                    dc1 = "1"
            End Select

            ' Ahora obtengo el segundo dígíto, que verificará
            ' el número de cuenta de cliente.
            '
            For n = 9 To 0 Step -1
                suma2 = suma2 + Convert.ToInt32(cuenta.Substring(n, 1)) * pesos(9 - n)
            Next

            ' Calculamos el cociente de dividir el resultado
            ' de la suma entre 11.
            '
            cociente2 = suma2 \ 11 ' Nos da un resultado entero

            ' Calculo el resto de la división. Para ello,
            ' en lugar de utilizar el operador Mod, utilizo
            ' la fórmula para obtener el resto de la división.
            '
            resto2 = suma2 - (11 * cociente2)

            dc2 = (11 - resto2).ToString

            Select Case dc2
                Case "11"
                    dc2 = "0"
                Case "10"
                    dc2 = "1"
            End Select

            ' Devuelvo el dígito de control.
            '
            Return dc1 & dc2
        Catch ex As Exception
            Throw ex
        End Try
        Return ""
    End Function
#End Region
#Region "Comprobar CIF-NIF"
    Public Shared Function Comprobar_CIF_NIF(ByRef oObjGlobal As EXO_Generales.EXO_General, ByVal sValor As String) As Boolean
        Comprobar_CIF_NIF = False
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim sSQL As String = ""
        Try
            oRs = CType(oObjGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            'Validamos el CIF o NIF
            If oObjGlobal.conexionSAP.compañia.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                sSQL = "SELECT ""EXO_VALIDAR_NIF_CIF""(RTRIM(LTRIM('" & sValor & "'))) ""Es_CIFNIF_OK"" FROM DUMMY;"
            Else
                sSQL = "SELECT [dbo].[EXO_VALIDAR_NIF_CIF](RTRIM(LTRIM('" & sValor & "'))) ""Es_CIFNIF_OK"" "
            End If
            oRs.DoQuery(sSQL)

            If oRs.RecordCount > 0 Then
                If CInt(oRs.Fields.Item("Es_CIFNIF_OK").Value.ToString) = 0 Then
                    oObjGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - El CIF/NIF " & sValor & " no es válido.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oObjGlobal.conexionSAP.SBOApp.MessageBox("El CIF/NIF " & sValor & " no es válido.")
                    Exit Function
                End If
            Else
                Throw New Exception("No se ha encontrado función EXO_VALIDAR_NIF_CIF")
                Exit Function
            End If
            Comprobar_CIF_NIF = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region

#Region "Creación de datos"
    Public Shared Function CrearDocumentos(ByRef oForm As SAPbouiCOM.Form, ByVal sData As String, ByVal sTDoc As String, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_Generales.EXO_General) As Boolean
        CrearDocumentos = False
        Dim oDoc As SAPbobsCOM.Documents = Nothing
        Dim sExiste As String = "" ' Para comprobar si existen los datos

        Dim sErrorDes As String = "" : Dim sDocAdd As String = "" : Dim sMensaje As String = ""
        Dim sTipoFac As String = "" : Dim sModo As String = "" : Dim sTabla As String = ""

        Dim oRsCab As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLin As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsLote As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsArt As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsCliente As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsSerie As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRsSerieNumber As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim esprimeralinea As Boolean = True
        Dim esprimeraportes As Boolean = True
        Try
            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If
            'Company.StartTransaction()
            For i = 0 To oForm.DataSources.DataTables.Item(sData).Rows.Count - 1
                If oForm.DataSources.DataTables.Item(sData).GetValue("Sel", i).ToString = "Y" Then 'Sólo los registros que se han seleccionado
                    sSQL = "Select * FROM ""@EXO_TMPDOC"" Where ""Code""=" & oForm.DataSources.DataTables.Item(sData).GetValue("Code", i).ToString & " and ""U_EXO_USR""='" & objglobal.usuario & "' "
                    oRsCab.DoQuery(sSQL)
                    If oRsCab.RecordCount > 0 Then
#Region "Cabecera"
                        Dim dImpTotal As Double = 0.00
                        esprimeralinea = True
#Region "Tipo Documento"
                        sTipoFac = oRsCab.Fields.Item("U_EXO_TIPOF").Value.ToString
                        sModo = oForm.DataSources.DataTables.Item(sData).GetValue("Modo", i).ToString
                        If sModo = "F" Then
                            Select Case sTipoFac
                                Case "13" 'Factura de ventas
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices), SAPbobsCOM.Documents)
                                    sTabla = "OINV"
                                Case "14" 'Abono de ventas
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes), SAPbobsCOM.Documents)
                                    sTabla = "ORIN"
                                Case "18" 'Factura de compras
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices), SAPbobsCOM.Documents)
                                    sTabla = "OPCH"
                                Case "19" 'Abono de compras
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes), SAPbobsCOM.Documents)
                                    sTabla = "ORPC"
                                Case "22" 'Pedidos de compras
                                    oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders), SAPbobsCOM.Documents)
                                    sTabla = "OPOR"
                            End Select
                        Else
                            oDoc = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts), SAPbobsCOM.Documents)
                            sTabla = "ODRF"
                        End If
                        Select Case sTipoFac
                            Case "13" 'Factura de ventas
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                            Case "14" 'Abono de ventas
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oCreditNotes
                            Case "18" 'Factura de compras
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseInvoices
                            Case "19" 'Abono de compras
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes
                            Case "22" 'Pedido de compras
                                oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseOrders
                        End Select
#End Region
#Region " Serie o Num Documento"
                        If oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString <> "" Then
                            ''Si se crea en borrador, habrá que buscar el número para no dejarlo crear
                            'If sTabla = "ODRF" Then
                            '    Dim sEncuentraDocNUm As String = ""
                            '    sEncuentraDocNUm = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """ODRF""", """DocNum""", """DocNum""=" & oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString)
                            '    If sEncuentraDocNUm <> "" Then
                            '        'Como lo ha encontrado, no podemos dejar crearlo

                            '    End If
                            'End If
                            oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                            oDoc.DocNum = oForm.DataSources.DataTables.Item(sData).GetValue("Nº Documento", i).ToString
                        Else
                            'Buscamos la Serie
                            Dim sSerie As String = oForm.DataSources.DataTables.Item(sData).GetValue("Serie", i).ToString
                            sSQL = "SELECT ""Series"" "
                            sSQL += " FROM ""NNM1"" "
                            sSQL += " WHERE ""ObjectCode""=" & sTipoFac & " and ""SeriesName""='" & sSerie & "' "
                            oRsSerie = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
                            oRsSerie.DoQuery(sSQL)
                            If oRsSerie.RecordCount > 0 Then
                                Dim sSerieDoc As String = oRsSerie.Fields.Item("Series").Value.ToString
                                oDoc.Series = sSerieDoc
                            Else
                                objglobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado serie para el documento.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                Exit Function
                            End If
                        End If
#End Region
                        oDoc.CardCode = oRsCab.Fields.Item("U_EXO_CLISAP").Value.ToString
                        oDoc.NumAtCard = oForm.DataSources.DataTables.Item(sData).GetValue("Referencia", i).ToString
                        oDoc.DocCurrency = oRsCab.Fields.Item("U_EXO_MONEDA").Value.ToString
                        If oRsCab.Fields.Item("U_EXO_CTABAL").Value.ToString <> "" Then
                            oDoc.ControlAccount = oRsCab.Fields.Item("U_EXO_CTABAL").Value.ToString
                        End If
                        'Hay que buscar el comercial para asignarlo
                        If oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString <> "" Then
                            Dim sCodComercial = ""
                            sCodComercial = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OSLP""", """SlpCode""", """SlpName""='" & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & "'")
                            If sCodComercial <> "" Then
                                oDoc.SalesPersonCode = sCodComercial
                            Else
                                oSboApp.SboApp.StatusBar.SetText("(EXO) - El empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                If oSboApp.SboApp.MessageBox("El empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & " - no existe. ¿Desea Crearlo?""?", 1, "Sí", "No") = 1 Then
                                    EXO_GLOBALES.CrearEmpleado(oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString, oCompany, oSboApp)
                                Else
                                    oSboApp.SboApp.StatusBar.SetText("(EXO) - No se puede continuar si no se crea el empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    oSboApp.SboApp.MessageBox("No se puede continuar si no se crea el empleado del Dpto.  - " & oForm.DataSources.DataTables.Item(sData).GetValue("Comercial", i).ToString & ".")
                                    Exit Function
                                End If
                            End If
                        End If
#Region "Fechas"
                        oDoc.DocDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString)
                        Try
                            oDoc.TaxDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Documento", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Documento", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Documento", i).ToString)
                        Catch ex As Exception
                            oDoc.TaxDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString)
                        End Try

                        Dim sFechaVTO As String = ""
                        Try
                            sFechaVTO = oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString
                        Catch ex As Exception
                            sFechaVTO = "1900-01-01"
                        End Try
                        If Year(sFechaVTO) > 1950 Then
                            oDoc.DocDueDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Vto", i).ToString)
                        Else
                            'oDoc.DocDueDate = Year(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Month(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString) & "-" & Day(oForm.DataSources.DataTables.Item(sData).GetValue("F. Contable", i).ToString)
                        End If
#End Region
                        If oRsCab.Fields.Item("U_EXO_DIRFAC").Value.ToString <> "" Then : oDoc.PayToCode = oRsCab.Fields.Item("U_EXO_DIRFAC").Value.ToString : End If
                        If oRsCab.Fields.Item("U_EXO_DIRENT").Value.ToString <> "" Then : oDoc.ShipToCode = oRsCab.Fields.Item("U_EXO_DIRENT").Value.ToString : End If
#Region "condición y modo de pago"
                        If oRsCab.Fields.Item("U_EXO_CPAGO").Value.ToString <> "" Then
                            oDoc.PaymentMethod = oRsCab.Fields.Item("U_EXO_CPAGO").Value.ToString
                        End If
                        If oRsCab.Fields.Item("U_EXO_GROUPNUM").Value.ToString <> "" Then
                            Dim sGroupNum As Integer = -1
                            Try
                                sGroupNum = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCTG""", """GroupNum""", """PymntGroup""='" & oRsCab.Fields.Item("U_EXO_GROUPNUM").Value.ToString & "'")
                            Catch ex As Exception
                                sGroupNum = -1
                            End Try
                            If sGroupNum >= 0 Then
                                oDoc.PaymentGroupCode = sGroupNum
                            End If
                        End If
#End Region
#Region "Comentarios"
                        oDoc.Comments = oForm.DataSources.DataTables.Item(sData).GetValue("Comentario", i).ToString
                        oDoc.OpeningRemarks = oRsCab.Fields.Item("U_EXO_CCAB").Value.ToString
                        oDoc.ClosingRemarks = oRsCab.Fields.Item("U_EXO_CPIE").Value.ToString
#End Region
#End Region

#Region "Líneas"
                        'Buscamos las líneas del documento
                        sSQL = "Select * FROM ""@EXO_TMPDOCL"" Where ""Code""=" & oRsCab.Fields.Item("Code").Value.ToString & " and ""U_EXO_USR""='" & objglobal.usuario & "' "
                        oRsLin.DoQuery(sSQL)
                        For iLin = 1 To oRsLin.RecordCount
                            If esprimeralinea = False Then
                                oDoc.Lines.Add()
                            Else
#Region "Tipo Líneas"
                                Select Case oRsLin.Fields.Item("U_EXO_DOCTYPE").Value.ToString
                                    Case "S" : oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                                    Case "I" : oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items
                                End Select
#End Region
                            End If
                            esprimeralinea = False
#Region "Norma Reparto Coste"
                            Dim sReparto As String = oRsLin.Fields.Item("U_EXO_REPARTO").Value.ToString
                            If sReparto <> "" Then
                                oDoc.Lines.CostingCode = sReparto
                            Else
                                oDoc.Lines.CostingCode = "12040201"
                            End If

#End Region
                            If oRsLin.Fields.Item("U_EXO_DOCTYPE").Value.ToString = "I" Then
                                oDoc.Lines.ItemCode = oRsLin.Fields.Item("U_EXO_ART").Value
                                If Trim(oRsLin.Fields.Item("U_EXO_ARTDES").Value.ToString) <> "" Then
                                    oDoc.Lines.ItemDescription = oRsLin.Fields.Item("U_EXO_ARTDES").Value
                                End If
                                oDoc.Lines.Quantity = oRsLin.Fields.Item("U_EXO_CANT").Value
                                oDoc.Lines.UnitPrice = oRsLin.Fields.Item("U_EXO_PRECIO").Value
                                oDoc.Lines.GrossBuyPrice = oRsLin.Fields.Item("U_EXO_PRECIOBRUTO").Value
                                oDoc.Lines.DiscountPercent = oRsLin.Fields.Item("U_EXO_DTOLIN").Value
                                dImpTotal += (oRsLin.Fields.Item("U_EXO_CANT").Value * oRsLin.Fields.Item("U_EXO_PRECIO").Value) - ((oRsLin.Fields.Item("U_EXO_DTOLIN").Value * (oRsLin.Fields.Item("U_EXO_CANT").Value * oRsLin.Fields.Item("U_EXO_PRECIO").Value)) / 100)
                                'Buscamos series disponibles
                                sSQL = "select t0.""SysNumber"" ""SysNumber"" "
                                sSQL &= " FROM ""OSRN"" t0 INNER JOIN ""OSRQ"" t1 on t0.""ItemCode""=t1.""ItemCode"" and t0.""SysNumber""=t1.""SysNumber"" "
                                sSQL &= " WHERE t0.""ItemCode""='" & oRsLin.Fields.Item("U_EXO_ART").Value.ToString & "' and t1.""Quantity"">0 ORDER BY ""SysNumber"""
                                oRsSerieNumber.DoQuery(sSQL)
                                'Incluimos los Lotes
                                sSQL = "Select * FROM ""@EXO_TMPDOCLT"" Where ""Code""=" & oRsLin.Fields.Item("Code").Value.ToString & " And ""U_EXO_USR""='" & objglobal.usuario & "' "
                                sSQL &= " and ""U_EXO_LineId""=" & oRsLin.Fields.Item("LineId").Value.ToString
                                oRsLote.DoQuery(sSQL)
                                For iLote = 1 To oRsLote.RecordCount
                                    'tengo que buscar el artículo para saber si va por lote o serie
                                    Dim sLote As String = "" : Dim sSerie As String = ""
                                    sSQL = "SELECT ""ManSerNum"", ""ManBtchNum"" FROM ""OITM"" WHERE ""ItemCode""='" & oRsLin.Fields.Item("U_EXO_ART").Value & "'"
                                    oRsArt.DoQuery(sSQL)
                                    If oRsArt.RecordCount > 0 Then
                                        sSerie = oRsArt.Fields.Item("ManSerNum").Value.ToString
                                        sLote = oRsArt.Fields.Item("ManBtchNum").Value.ToString
                                    End If
                                    If sLote = "Y" Then
                                        'Creamos el lote de la línea del artículo
                                        oDoc.Lines.BatchNumbers.BatchNumber = oRsLote.Fields.Item("U_EXO_Lote").Value.ToString
                                        oDoc.Lines.BatchNumbers.Quantity = oRsLote.Fields.Item("U_EXO_CANT").Value.ToString
                                        oDoc.Lines.BatchNumbers.Add()
                                    ElseIf sSerie = "Y" Then
                                        'Creamos la serie de la línea del artículo
                                        Select Case sTipoFac
                                            Case "14", "18"
                                                oDoc.Lines.SerialNumbers.InternalSerialNumber = oRsLote.Fields.Item("U_EXO_Lote").Value.ToString
                                            Case "13", "19"
                                                If oRsSerieNumber.RecordCount = 0 Then
                                                    oSboApp.SboApp.StatusBar.SetText("(EXO) - No hay series disponibles para generar el documento.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    oSboApp.SboApp.MessageBox("No hay series disponibles para generar el documento.")
                                                    Exit Function
                                                End If
                                                'Tenemos que buscar el system serial que la cantidad sea superior a 0
                                                Dim iSerialNumber As Integer = 0
                                                iSerialNumber = oRsSerieNumber.Fields.Item("SysNumber").Value
                                                oDoc.Lines.SerialNumbers.SystemSerialNumber = iSerialNumber
                                                oRsSerieNumber.MoveNext()
                                        End Select
                                        oDoc.Lines.SerialNumbers.Quantity = oRsLote.Fields.Item("U_EXO_CANT").Value.ToString
                                        oDoc.Lines.SerialNumbers.Add()
                                    End If
                                    'oDoc.Lines.WarehouseCode = "01"
                                    oRsLote.MoveNext()
                                Next

                                oDoc.Lines.FreeText = oRsLin.Fields.Item("U_EXO_TXT").Value
                            ElseIf oRsLin.Fields.Item("U_EXO_DOCTYPE").Value.ToString = "S" Then
                                oDoc.Lines.AccountCode = oRsLin.Fields.Item("U_EXO_CTA").Value
                                oDoc.Lines.LineTotal = oRsLin.Fields.Item("U_EXO_IMPSRV").Value
                                dImpTotal += oRsLin.Fields.Item("U_EXO_IMPSRV").Value

                                oDoc.Lines.ItemDescription = oRsLin.Fields.Item("U_EXO_TXT").Value
                            End If
#Region "Impuesto y Retencion de línea"
                            If oRsLin.Fields.Item("U_EXO_Impuesto").Value <> "" Then
                                oDoc.Lines.VatGroup = oRsLin.Fields.Item("U_EXO_Impuesto").Value
                            End If

                            If oRsLin.Fields.Item("U_EXO_Retencion").Value = "" Then
                                oDoc.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tNO
                            Else
                                oDoc.Lines.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES
                                If oRsLin.Fields.Item("U_EXO_Retencion").Value <> "" Then
                                    oDoc.WithholdingTaxData.WTCode = oRsLin.Fields.Item("U_EXO_Retencion").Value
                                    oDoc.WithholdingTaxData.Add()
                                End If
                            End If
#End Region
                            oRsLin.MoveNext()
                        Next
#End Region
#Region "Dto en cabecera"
                        If oRsCab.Fields.Item("U_EXO_TDTO").Value.ToString = "%" Then
                            oDoc.DiscountPercent = oForm.DataSources.DataTables.Item(sData).GetValue("Dto.", i).ToString
                        Else
                            oDoc.DiscountPercent = (oForm.DataSources.DataTables.Item(sData).GetValue("Dto.", i).ToString * 100) / dImpTotal
                        End If
#End Region
                        'grabar el documento
                        If oDoc.Add() <> 0 Then 'Si ocurre un error en la grabación entra
                            sErrorDes = oCompany.GetLastErrorCode & " / " & oCompany.GetLastErrorDescription
                            oSboApp.StatusBar.SetText(sErrorDes, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "ERROR")
                            oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sErrorDes)
                            oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, "")
                        Else
                            esprimeralinea = True
                            esprimeraportes = True
                            sDocAdd = oCompany.GetNewObjectKey() 'Recoge el último documento creado
                            oForm.DataSources.DataTables.Item(sData).SetValue("DocEntry", i, sDocAdd)
                            'Buscamos el documento para crear un mensaje
                            sDocAdd = EXO_GLOBALES.GetValueDB(oCompany, """" & sTabla & """", """DocNum""", """DocEntry""=" & sDocAdd)
                            If sModo = "F" Then
                                sModo = ""
                            Else
                                sModo = " borrador "
                            End If
                            oForm.DataSources.DataTables.Item(sData).SetValue("Estado", i, "OK")
                            oForm.DataSources.DataTables.Item(sData).SetValue("Nº Documento", i, sDocAdd)
                            Select Case sTipoFac
                                Case "13" 'Factura de ventas
                                    sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de ventas Nº" & sDocAdd
                                Case "14" 'Abono de ventas
                                    sMensaje = "(EXO) - Ha sido creado el abono " & sModo & " de ventas Nº" & sDocAdd
                                Case "18" 'Factura de compras
                                    sMensaje = "(EXO) - Ha sido creada la factura " & sModo & " de compras Nº" & sDocAdd
                                Case "19" 'Abono de compras
                                    sMensaje = "(EXO) - Ha sido creado el abono " & sModo & " de compras Nº" & sDocAdd
                                Case "22" 'Pedido de compras
                                    sMensaje = "(EXO) - Ha sido creado el pedido " & sModo & " de compras Nº" & sDocAdd
                            End Select
                            oForm.DataSources.DataTables.Item(sData).SetValue("Descripción Estado", i, sMensaje)
                            oSboApp.StatusBar.SetText(sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        End If
                    End If
                End If
            Next

            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            'End If

            CrearDocumentos = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            'If Company.InTransaction = True Then
            '    Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            'End If

            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oDoc, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsCab, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLin, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsLote, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsSerie, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsSerieNumber, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRsArt, Object))
        End Try
    End Function
    Public Shared Sub CrearEmpleado(ByVal sEmpleado As String, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application)
        Dim oEmpleado As SAPbobsCOM.SalesPersons = Nothing
        Try
            oEmpleado = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesPersons), SAPbobsCOM.SalesPersons)
            oEmpleado.SalesEmployeeName = sEmpleado
            oEmpleado.Active = SAPbobsCOM.BoYesNoEnum.tYES
            If oEmpleado.Add() <> 0 Then
                Throw New Exception(oCompany.GetLastErrorCode & " / " & oCompany.GetLastErrorDescription)
            Else
                oSboApp.StatusBar.SetText("(EXO) - Se ha creado el Empleado " & sEmpleado & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oEmpleado, Object))
        End Try
    End Sub
    Public Shared Sub CrearInterlocutor(ByVal sSerieI As String, ByRef sCodigoSAP As String, ByVal sAddID As String, ByVal sBusinessPartnerAccion As String, ByVal sBusinessPartnerTipo As String, ByVal sBusinessPartnerRazonSocial As String,
                                  ByVal sBusinessPartnerPais As String, ByVal sBusinessPartnerIFiscal As String, ByVal sBusinessPartnerTIdentificacion As String, ByVal sBusinessPartnerTipoEmpresa As String,
                                  ByVal sBusinessPartnerTel1 As String, ByVal sBusinessPartnerTel2 As String, ByVal sBusinessPartnerFax As String, ByVal sBusinessPartnercorreo As String, ByVal sBusinessPartnerCtaBalance As String,
                                  ByVal iContacto As Integer, ByRef sContactoAccion() As String, ByRef sContactoCodigo() As String, ByRef sContactoNombre() As String, ByRef sContactoApe1() As String, ByRef sContactoApe2() As String, ByRef sContactoTel() As String,
                                  ByRef sContactoMovil() As String, ByRef sContactoCorreo() As String, ByRef sContactoPuesto() As String, ByVal iDir As Integer, ByRef sDirAccion() As String, ByRef sDirTipo() As String,
                                  ByRef sDirCodigo() As String, ByRef sDirCalle() As String, ByRef sDirNum() As String, ByRef sDirBloque() As String, ByRef sDirEdif() As String, ByRef sDirCiudad() As String,
                                  ByRef sDirCodPostal() As String, ByRef sDirProvincia() As String, ByRef sDirPais() As String, ByVal iBanco As Integer, ByRef sBPBancoAccion() As String, ByRef sBPBancoPais() As String,
                                  ByRef sBPBanco() As String, ByRef sBPBancoSucursal() As String, ByRef sBPBancoCuenta() As String, ByVal sBPCViaPago As String, ByVal iDiasPago As Integer, ByRef sBPDiasPago() As String,
                                  ByVal sBPGestionRE As String, ByVal sBPI347 As String, ByVal sBPI347A As String, ByVal sBPImpuestoCod As String, ByVal iRetDetalle As Integer, ByRef sBPRetCodigo() As String,
                                  ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application)
        Dim oOCRD As SAPbobsCOM.BusinessPartners = Nothing
        Dim sSQl As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oOCRD = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)
            If sBusinessPartnerAccion = "U" Then
                'Si es modificación habrá que buscar el Interlocutor
                If oOCRD.GetByKey(sCodigoSAP) = False Then
                    oSboApp.StatusBar.SetText("(EXO) - No se puede modificar el interlocutor - " & sCodigoSAP & " -  No se ha encontrado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oSboApp.MessageBox("No se puede modificar el interlocutor - " & sCodigoSAP & " -  No se ha encontrado.")
                    Exit Sub
                End If
            End If

            If sSerieI <> "" And sBusinessPartnerAccion = "A" Then
                oOCRD.Series = sSerieI
            Else
                If sCodigoSAP <> "" And sBusinessPartnerAccion = "A" Then
                    oOCRD.CardCode = sCodigoSAP
                End If
            End If
            If sAddID <> "" Then : oOCRD.AdditionalID = sAddID : End If
            If sBusinessPartnerTipo <> "" Then
                Select Case sBusinessPartnerTipo
                    Case "C" : oOCRD.CardType = SAPbobsCOM.BoCardTypes.cCustomer
                    Case "S" : oOCRD.CardType = SAPbobsCOM.BoCardTypes.cSupplier
                End Select
            End If
            If sBusinessPartnerRazonSocial <> "" Then : oOCRD.CardName = sBusinessPartnerRazonSocial : End If
            If sBusinessPartnerPais <> "" Then : oOCRD.Country = sBusinessPartnerPais : End If
            If sBusinessPartnerIFiscal <> "" Then : oOCRD.FederalTaxID = sBusinessPartnerIFiscal : End If
            Select Case sBusinessPartnerTIdentificacion
                Case "1" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntSpanishFiscalID
                Case "2" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntVATRegistrationNumber 'NIF comunitario
                Case "3" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntPassport 'Pasaporte
                Case "4" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntFiscalIDIssuedbytheResidenceCountry 'ID fiscal emitido por el país de residencia
                Case "5" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntCertificateofFiscalResidence 'Certificado de residencia fiscal
                Case "6" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntOtherDocument 'Otro documento
                Case Else : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntOtherDocument 'No registrado No hay posibilidad para este
            End Select

            If sBusinessPartnerTipoEmpresa <> "" Then
                Select Case sBusinessPartnerTipoEmpresa
                    Case "A" : oOCRD.CompanyPrivate = SAPbobsCOM.BoCardCompanyTypes.cPrivate
                    Case "C" : oOCRD.CompanyPrivate = SAPbobsCOM.BoCardCompanyTypes.cCompany
                    Case Else : oOCRD.CompanyPrivate = SAPbobsCOM.BoCardCompanyTypes.cGovernment
                End Select
            End If
            If sBusinessPartnerTel1 <> "" Then : oOCRD.Phone1 = sBusinessPartnerTel1 : End If
            If sBusinessPartnerTel2 <> "" Then : oOCRD.Phone2 = sBusinessPartnerTel2 : End If
            If sBusinessPartnerFax <> "" Then : oOCRD.Fax = sBusinessPartnerFax : End If
            If sBusinessPartnercorreo <> "" Then : oOCRD.MailAddress = sBusinessPartnercorreo : End If
            'Revisar si es este el campo
            If sBusinessPartnerCtaBalance <> "" Then : oOCRD.AccountRecivablePayables.AccountCode = sBusinessPartnerCtaBalance : End If

            For i = 1 To iContacto
                If sContactoAccion(i) = "U" Then
                    sSQl = "SELECT ""Name"" FROM ""OCPR"" Where ""CardCode""='" & sCodigoSAP & "' and ""Name""='" & sContactoCodigo(i) & "' "
                    oRs.DoQuery(sSQl)
                    If oRs.RecordCount > 0 Then
                        sSQl = "SELECT ""Name"" FROM ""OCPR"" Where ""CardCode""='" & sCodigoSAP & "' Order By ""CntctCode"" "
                        oRs.DoQuery(sSQl)
                        For linea = 0 To oRs.RecordCount - 1
                            If oRs.Fields.Item("Name").Value.ToString = sContactoCodigo(i) Then
                                oOCRD.ContactEmployees.SetCurrentLine(linea)
                            End If
                            oRs.MoveNext()
                        Next
                    Else
                        'No se ha encontrado el contacto, no se puede modificar.
                        oSboApp.StatusBar.SetText("(EXO) - No se puede modificar el contacto """ & sContactoCodigo(i) & """" & " del interlocutor - " & sCodigoSAP & " -  No se ha encontrado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oSboApp.MessageBox("No se puede modificar el contacto """ & sContactoCodigo(i) & """" & " del interlocutor - " & sCodigoSAP & " -  No se ha encontrado.")
                        Exit Sub
                    End If
                End If
                oOCRD.ContactEmployees.Name = sContactoCodigo(i)
                oOCRD.ContactEmployees.Active = SAPbobsCOM.BoYesNoEnum.tYES
                oOCRD.ContactEmployees.FirstName = sContactoNombre(i) : oOCRD.ContactEmployees.LastName = sContactoApe1(i) : oOCRD.ContactEmployees.MiddleName = sContactoApe2(i)
                oOCRD.ContactEmployees.Phone1 = sContactoTel(i) : oOCRD.ContactEmployees.MobilePhone = sContactoMovil(i) : oOCRD.ContactEmployees.E_Mail = sContactoCorreo(i)
                oOCRD.ContactEmployees.Position = sContactoPuesto(i)
                If sContactoAccion(i) = "A" Then
                    oOCRD.ContactEmployees.Add()
                End If
                oOCRD.ContactPerson = sContactoCodigo(i)
            Next

            For i = 1 To iDir
                If sDirAccion(i) = "U" Then
                    sSQl = "SELECT ""LineNum"" FROM ""CRD1"" Where ""CardCode""='" & sCodigoSAP & "' and ""Address""='" & sDirCodigo(i) & "' "
                    oRs.DoQuery(sSQl)
                    If oRs.RecordCount > 0 Then
                        sSQl = "SELECT ""LineNum"" FROM ""CRD1"" Where ""CardCode""='" & sCodigoSAP & "' Order By ""LineNum"" "
                        oRs.DoQuery(sSQl)
                        For linea = 0 To oRs.RecordCount - 1
                            If oRs.Fields.Item("LineNum").Value.ToString = sContactoCodigo(i) Then
                                oOCRD.Addresses.SetCurrentLine(linea)
                            End If
                            oRs.MoveNext()
                        Next
                    Else
                        'No se ha encontrado la dirección, no se puede modificar.
                        oSboApp.StatusBar.SetText("(EXO) - No se puede modificar la dirección """ & sDirCodigo(i) & """" & " del interlocutor - " & sCodigoSAP & " -  No se ha encontrado.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        oSboApp.MessageBox("No se puede modificar el contacto """ & sDirCodigo(i) & """" & " del interlocutor - " & sCodigoSAP & " -  No se ha encontrado.")
                        Exit Sub
                    End If
                End If
                If sDirTipo(i) = "B" Then
                    oOCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                Else
                    oOCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo
                End If
                oOCRD.Addresses.AddressName = sDirCodigo(i)
                oOCRD.Addresses.Street = sDirCalle(i) : oOCRD.Addresses.StreetNo = sDirNum(i) : oOCRD.Addresses.Block = sDirBloque(i) : oOCRD.Addresses.BuildingFloorRoom = sDirEdif(i)
                oOCRD.Addresses.City = sDirCiudad(i) : oOCRD.Addresses.ZipCode = sDirCodPostal(i) : oOCRD.Addresses.County = sDirProvincia(i) : oOCRD.Addresses.Country = sDirPais(i)
                If sDirAccion(i) = "A" Then
                    oOCRD.Addresses.Add()
                End If
            Next

            For i = 1 To iBanco
                If sBPBancoPais(i) <> "" Then : oOCRD.HouseBankCountry = sBPBancoPais(i) : End If
                If sBPBanco(i) <> "" Then : oOCRD.HouseBank = sBPBanco(i) : End If
                If sBPBancoSucursal(i) <> "" Then : oOCRD.HouseBankBranch = sBPBancoSucursal(i) : End If
                If sBPBancoCuenta(i) <> "" Then : oOCRD.HouseBankAccount = sBPBancoCuenta(i) : End If
            Next

            'Via de pago
            If sBPCViaPago <> "" Then
                If sBusinessPartnerAccion = "A" Then
                    oOCRD.BPPaymentMethods.PaymentMethodCode = sBPCViaPago
                    oOCRD.BPPaymentMethods.Add()
                Else
                    oOCRD.BPPaymentMethods.SetCurrentLine(0)
                    oOCRD.BPPaymentMethods.PaymentMethodCode = sBPCViaPago
                End If

                oOCRD.PeymentMethodCode = sBPCViaPago
            End If

            For i = 1 To iDiasPago
                If sBusinessPartnerAccion = "A" Then
                    oOCRD.BPPaymentDates.PaymentDate = sBPDiasPago(i)
                    oOCRD.BPPaymentDates.Add()
                Else
                    oOCRD.BPPaymentDates.SetCurrentLine(i - 1)
                    oOCRD.BPPaymentDates.PaymentDate = sBPDiasPago(i)
                End If

            Next

            'Recargo de Equivalencia como no sabemos los valores si escribe algo es que si tiene, si está ne blanco no tiene
            If sBPGestionRE <> "" Then
                oOCRD.Equalization = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oOCRD.Equalization = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            Select Case sBPI347
                Case "1"
                    Select Case sBusinessPartnerTipo
                        Case "C" : oOCRD.OperationCode347 = SAPbobsCOM.OperationCode347Enum.ocSalesOrServicesRevenues 'repasar dato con tipo de interlocutor
                        Case "S" : oOCRD.OperationCode347 = SAPbobsCOM.OperationCode347Enum.ocGoodsOrServiciesAcquisitions 'repasar dato con tipo de interlocutor
                    End Select

                Case "2" : oOCRD.OperationCode347 = SAPbobsCOM.OperationCode347Enum.ocPublicEntitiesAcquisitions
                Case "3"
                    Select Case sBusinessPartnerTipo
                        Case "C" : oOCRD.OperationCode347 = SAPbobsCOM.OperationCode347Enum.ocTravelAgenciesSales 'repasar dato con tipo de interlocutor
                        Case "S" : oOCRD.OperationCode347 = SAPbobsCOM.OperationCode347Enum.ocTravelAgenciesPurchases 'repasar dato con tipo de interlocutor
                    End Select
            End Select

            If sBPI347A <> "" Then
                oOCRD.InsuranceOperation347 = SAPbobsCOM.BoYesNoEnum.tYES
            Else
                oOCRD.InsuranceOperation347 = SAPbobsCOM.BoYesNoEnum.tNO
            End If

            If sBPImpuestoCod <> "" Then : oOCRD.VatGroup = sBPImpuestoCod : End If

            'Retenciones
            oOCRD.WTCode = sBPRetCodigo(1)

            For i = 1 To iRetDetalle
                If sBusinessPartnerAccion = "A" Then
                    oOCRD.BPWithholdingTax.WTCode = sBPRetCodigo(i)
                    oOCRD.BPWithholdingTax.Add()
                Else
                    oOCRD.BPWithholdingTax.SetCurrentLine(i - 1)
                    oOCRD.BPWithholdingTax.WTCode = sBPRetCodigo(i)
                End If
            Next


            If sBusinessPartnerAccion = "U" Then
                If oOCRD.Update() <> 0 Then
                    Throw New Exception(oCompany.GetLastErrorCode & " / No se puede actualizar el interlocutor -" & sBusinessPartnerRazonSocial & " - " & oCompany.GetLastErrorDescription)
                Else
                    oSboApp.StatusBar.SetText("(EXO) - Se ha modificado el interlocutor - " & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            ElseIf sBusinessPartnerAccion = "A" Then
                If oOCRD.Add() <> 0 Then
                    Throw New Exception(oCompany.GetLastErrorCode & " / No se puede crear el interlocutor -" & sBusinessPartnerRazonSocial & " - " & oCompany.GetLastErrorDescription)
                Else
                    sCodigoSAP = oCompany.GetNewObjectKey()
                    oSboApp.StatusBar.SetText("(EXO) - Se ha creado el interlocutor - " & sCodigoSAP & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    Public Shared Sub CrearInterlocutorSencillo(ByVal sFormato As String, ByRef sCodigoSAP As String, ByVal sSerieI As String, ByVal sBusinessPartnerTipo As String, ByVal sBusinessPartnerRazonSocial As String,
                                  ByVal sBusinessPartnerPais As String, ByVal sBusinessPartnerIFiscal As String, ByVal sAddID As String, ByVal sBusinessPartnerTIdentificacion As String,
                                  ByVal sBPCCPago As String, ByVal sBPCViaPago As String, ByVal sDirFac As String, ByVal sDirEnv As String, ByVal sDIR As String, ByVal sPob As String, ByVal sProv As String,
                                  ByVal sCPos As String, ByVal sCPais As String, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef oObjGlobal As EXO_Generales.EXO_General)
        Dim oOCRD As SAPbobsCOM.BusinessPartners = Nothing
        Dim sSQl As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Try
            oRs = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            oOCRD = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners), SAPbobsCOM.BusinessPartners)

            If sSerieI <> "" Then
                oOCRD.Series = sSerieI
            End If
            If sAddID <> "" Then : oOCRD.AdditionalID = sAddID : End If
            If sBusinessPartnerTipo <> "" Then
                Select Case sBusinessPartnerTipo
                    Case "C" : oOCRD.CardType = SAPbobsCOM.BoCardTypes.cCustomer
                    Case "S" : oOCRD.CardType = SAPbobsCOM.BoCardTypes.cSupplier
                End Select
            End If
            'Ponemos el grupo de alquileres
            Dim sGrupo As String = ""
            sGrupo = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """OCRG""", """GroupCode""", """GroupName""='ALQUILER'")
            oOCRD.GroupCode = sGrupo
            If sBusinessPartnerRazonSocial <> "" Then : oOCRD.CardName = sBusinessPartnerRazonSocial : End If
            If sBusinessPartnerPais <> "" Then : oOCRD.Country = sBusinessPartnerPais : End If
            If sBusinessPartnerIFiscal <> "" Then : oOCRD.FederalTaxID = sBusinessPartnerIFiscal : End If
            Select Case sBusinessPartnerTIdentificacion
                Case "1" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntSpanishFiscalID
                Case "2" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntVATRegistrationNumber 'NIF comunitario
                Case "3" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntPassport 'Pasaporte
                Case "4" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntFiscalIDIssuedbytheResidenceCountry 'ID fiscal emitido por el país de residencia
                Case "5" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntCertificateofFiscalResidence 'Certificado de residencia fiscal
                Case "6" : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntOtherDocument 'Otro documento
                Case Else : oOCRD.ResidenNumber = SAPbobsCOM.ResidenceNumberTypeEnum.rntOtherDocument 'No registrado No hay posibilidad para este
            End Select

            'Direcciones
            For i = 1 To 2
                If i = 1 Then
                    oOCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_BillTo
                    oOCRD.Addresses.AddressName = sDirFac
                Else
                    oOCRD.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo
                    oOCRD.Addresses.AddressName = sDirEnv
                End If
                oOCRD.Addresses.Street = sDIR : oOCRD.Addresses.City = sPob : oOCRD.Addresses.ZipCode = sCPos : oOCRD.Addresses.County = sProv : oOCRD.Addresses.Country = sCPais
                oOCRD.Addresses.Add()
            Next


            'Condición de pago
            If sBPCCPago <> "" Then
                sBPCCPago = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """OCTG""", """PymntGroup""", """PymntGroup""='" & sBPCCPago & "'")
                oSboApp.StatusBar.SetText("(EXO) - La condición de pago indicada no existe - " & sBPCCPago & " -  Se tomará el valor por defecto de la configuración.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            If sBPCCPago = "" Then
                Select Case sBusinessPartnerTipo
                    Case "C" : sBPCCPago = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CPV""", """Code""='" & sFormato & "'")
                    Case "S" : sBPCCPago = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CPC""", """Code""='" & sFormato & "'")
                End Select
            End If
            oOCRD.PayTermsGrpCode = sBPCCPago

            'Via de pago
            'Comprobamos que exista, si no existe cogemos el valor por defecto del configurador
            If sBPCViaPago <> "" Then
                sBPCViaPago = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """OPYM""", """PayMethCod""", """PayMethCod""='" & sBPCViaPago & "'")
                oSboApp.StatusBar.SetText("(EXO) - La vía de pago indicada no existe - " & sBPCViaPago & " -  Se tomará el valor por defecto de la configuración.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
            If sBPCViaPago = "" Then
                Select Case sBusinessPartnerTipo
                    Case "C" : sBPCViaPago = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_VIAPV""", """Code""='" & sFormato & "'")
                    Case "S" : sBPCViaPago = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_VIAPC""", """Code""='" & sFormato & "'")
                End Select
            End If
            oOCRD.BPPaymentMethods.PaymentMethodCode = sBPCViaPago
            oOCRD.BPPaymentMethods.Add()
            oOCRD.PeymentMethodCode = sBPCViaPago


            'Recargo de Equivalencia como no sabemos los valores si escribe algo es que si tiene, si está ne blanco no tiene
            'If sBPGestionRE <> "" Then
            '    oOCRD.Equalization = SAPbobsCOM.BoYesNoEnum.tYES
            'Else
            oOCRD.Equalization = SAPbobsCOM.BoYesNoEnum.tNO
            'End If

            'Indicamos el valor por defectoi de la pantalla de configuración
            Dim sBPImpuestoCod As String = ""
            Select Case sBusinessPartnerTipo
                Case "C" : sBPImpuestoCod = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAV""", """Code""='" & sFormato & "'")
                Case "S" : sBPImpuestoCod = EXO_GLOBALES.GetValueDB(oObjGlobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAP""", """Code""='" & sFormato & "'")
            End Select
            If sBPImpuestoCod <> "" Then
                oOCRD.VatGroup = sBPImpuestoCod
            Else
                Dim sMensaje As String = ""
                sMensaje = "No se ha indicado el impuesto para el interlocutor en la ventana de configuración."
                oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oSboApp.MessageBox(sMensaje)
                Exit Sub
            End If

            If oOCRD.Add() <> 0 Then
                Throw New Exception(oCompany.GetLastErrorCode & " / No se puede crear el interlocutor -" & sBusinessPartnerRazonSocial & " - " & oCompany.GetLastErrorDescription)
            Else
                sCodigoSAP = oCompany.GetNewObjectKey()
                oSboApp.StatusBar.SetText("(EXO) - Se ha creado el interlocutor - " & sCodigoSAP & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOCRD, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Sub
    'Private Sub CrearArticulo(ByVal sArt As String, ByVal sArtDes As String, ByVal sArtInventariable As String, ByVal sArtCompras As String, ByVal sArtVentas As String, ByVal sArtGrupoArticulos As String,
    '                          ByVal sArtEanCode As String, ByVal sArtLotes As String, ByVal sArtSeries As String, ByVal sProveedorHabitual As String, ByVal sUnidadCompra As String, ByVal sLongitudCompra As String,
    '                          ByVal sAlturaCompra As String, ByVal sAnchoCompra As String, ByVal sVolumenCompra As String, ByVal sPesoCompra As String, ByVal sUnidadVenta As String, ByVal sLongitudVenta As String,
    '                          ByVal sAlturaVenta As String, ByVal sAnchoVenta As String, ByVal sVolumenVenta As String, ByVal sPesoVenta As String, ByVal sUnidadAlm As String, ByVal sStocksPorAlm As String,
    '                          ByVal sCuentasPorAlm As String)
    '    Dim oOITM As SAPbobsCOM.Items = Nothing

    '    Try
    '        oOITM = CType(Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems), SAPbobsCOM.Items)
    '        oOITM.ItemCode = sArt
    '        oOITM.ItemName = sArtDes
    '        Select Case sArtInventariable
    '            Case "Y" : oOITM.InventoryItem = SAPbobsCOM.BoYesNoEnum.tYES
    '            Case Else : oOITM.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO
    '        End Select
    '        Select Case sArtCompras
    '            Case "Y" : oOITM.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tYES
    '            Case Else : oOITM.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO
    '        End Select
    '        Select Case sArtVentas
    '            Case "Y" : oOITM.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
    '            Case Else : oOITM.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO
    '        End Select
    '        If sArtGrupoArticulos <> "" Then
    '            oOITM.ItemsGroupCode = sArtGrupoArticulos
    '        End If
    '        oOITM.BarCode = sArtEanCode
    '        If sArtLotes = "Y" Then
    '            oOITM.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES
    '            oOITM.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tNO
    '        ElseIf sArtSeries = "Y" Then
    '            oOITM.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES
    '            oOITM.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tNO
    '        Else
    '            oOITM.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tNO
    '            oOITM.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tNO
    '        End If
    '        oOITM.ItemClass = SAPbobsCOM.ItemClassEnum.itcMaterial

    '        'Compras
    '        oOITM.PreferredVendors.BPCode = sProveedorHabitual
    '        oOITM.PreferredVendors.Add()
    '        If sUnidadCompra <> "" Then : oOITM.PurchaseUnit = sUnidadCompra : End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sLongitudCompra) <> 0 Then
    '            oOITM.PurchaseLengthUnit = 2 : oOITM.PurchaseUnitLength = EXO_GLOBALES.TextToDbl(objGlobal, sLongitudCompra)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sAlturaCompra) <> 0 Then
    '            oOITM.PurchaseHeightUnit = 2 : oOITM.PurchaseUnitHeight = EXO_GLOBALES.TextToDbl(objGlobal, sAlturaCompra)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sAnchoCompra) <> 0 Then
    '            oOITM.PurchaseWidthUnit = 2 : oOITM.PurchaseUnitWidth = EXO_GLOBALES.TextToDbl(objGlobal, sAnchoCompra)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sVolumenCompra) <> 0 Then
    '            oOITM.PurchaseVolumeUnit = 2 : oOITM.PurchaseUnitVolume = EXO_GLOBALES.TextToDbl(objGlobal, sVolumenCompra)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sPesoCompra) <> 0 Then
    '            oOITM.PurchaseWeightUnit = 2 : oOITM.PurchaseUnitWeight = EXO_GLOBALES.TextToDbl(objGlobal, sPesoCompra)
    '        End If
    '        'Ventas
    '        If sUnidadVenta <> "" Then : oOITM.SalesUnit = sUnidadVenta : End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sLongitudVenta) <> 0 Then
    '            oOITM.SalesLengthUnit = 2 : oOITM.SalesUnitLength = EXO_GLOBALES.TextToDbl(objGlobal, sLongitudVenta)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sAlturaVenta) <> 0 Then
    '            oOITM.SalesHeightUnit = 2 : oOITM.SalesUnitHeight = EXO_GLOBALES.TextToDbl(objGlobal, sAlturaVenta)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sAnchoVenta) <> 0 Then
    '            oOITM.SalesWidthUnit = 2 : oOITM.SalesUnitWidth = EXO_GLOBALES.TextToDbl(objGlobal, sAnchoVenta)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sVolumenVenta) <> 0 Then
    '            oOITM.SalesVolumeUnit = 2 : oOITM.SalesUnitVolume = EXO_GLOBALES.TextToDbl(objGlobal, sVolumenVenta)
    '        End If
    '        If EXO_GLOBALES.TextToDbl(objGlobal, sPesoVenta) <> 0 Then
    '            oOITM.SalesWeightUnit = 2 : oOITM.SalesUnitWeight = EXO_GLOBALES.TextToDbl(objGlobal, sPesoVenta)
    '        End If
    '        'Alamcen
    '        oOITM.InventoryUOM = sUnidadAlm
    '        If sStocksPorAlm = "Y" Then
    '            oOITM.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tYES
    '        Else
    '            oOITM.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tNO
    '        End If
    '        Select Case sCuentasPorAlm
    '            Case "A" : oOITM.GLMethod = SAPbobsCOM.BoGLMethods.glm_WH
    '            Case "F" : oOITM.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass
    '            Case Else : oOITM.GLMethod = SAPbobsCOM.BoGLMethods.glm_ItemClass
    '        End Select
    '        If oOITM.Add() <> 0 Then
    '            Throw New Exception(Company.GetLastErrorCode & " / El Artículo que se intenta crear es: " & sArt & ". " & Company.GetLastErrorDescription)
    '        Else
    '            SboApp.StatusBar.SetText("(EXO) - Se ha creado el artículo - " & sArt & " - " & ". Por favor, revise los datos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '        End If
    '    Catch exCOM As System.Runtime.InteropServices.COMException
    '        Throw exCOM
    '    Catch ex As Exception
    '        Throw ex
    '    Finally
    '        EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oOITM, Object))

    '    End Try
    'End Sub
#End Region
#Region "Leer datos"
    Public Shared Function Leer_Campo(ByVal sCampo As String, ByVal sColumna As String, ByVal sObligatorio As String, ByRef sVCampo() As String, ByRef oSboApp As Application) As String
        Leer_Campo = ""
        Dim sValor As String = ""
        Dim icampo As Integer = 0
        Try
            If sColumna <> "" Then
                icampo = CInt(sColumna)
            End If
            If sVCampo(icampo) <> "" Then
                sValor = sVCampo(icampo)
            Else
                If sObligatorio = "Y" Then
                    Mensaje_CampoObligatorio(sCampo, sColumna, oSboApp)
                End If
            End If
            Leer_Campo = sValor
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

        End Try
    End Function
    Public Shared Sub Mensaje_CampoObligatorio(ByVal sCampo As String, ByVal sColumna As String, ByRef oSboApp As Application)
        Dim sMensaje As String = "El campo """ & sCampo & """ es obligatorio y la columna """ & sColumna & """ está vacía." & ChrW(13) & ChrW(10)
        sMensaje &= "Por favor, Revise el documento."
        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        oSboApp.MessageBox(sMensaje)
    End Sub
#End Region
#Region "Tratar ficheros"
    Public Shared Sub TratarFichero_TXT(ByVal sArchivo As String, ByVal sDelimitador As String, ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_Generales.EXO_General)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRCampos As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCampo As String = ""

        Dim iDoc As Integer = 0 'Contador de Cabecera de documentos
        Dim sTFac As String = "" : Dim sTFacColumna As String = "" : Dim sTipoLineas As String = "" : Dim sTDoc As String = ""
        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCodCliente As String = "" : Dim sClienteColumna As String = "" : Dim sCodClienteColumna As String = ""
        Dim sSerie As String = "" : Dim sDocNum As String = "" : Dim sManual As String = "" : Dim sSerieColumna As String = "" : Dim sDocNumColumna As String = ""
        Dim sDIR As String = "" : Dim sPob As String = "" : Dim sProv As String = "" : Dim sCPos As String = ""
        Dim sNumAtCard As String = "" : Dim sNumAtCardColumna As String = ""
        Dim sMoneda As String = "" : Dim sMonedaColumna As String = ""
        Dim sEmpleado As String = ""
        Dim sFContable As String = "" : Dim sFDocumento As String = "" : Dim sFVto As String = "" : Dim sFDocumentoColumna As String = ""
        Dim sTipoDto As String = "" : Dim sDto As String = ""
        Dim sPeyMethod As String = "" : Dim sCondPago As String = ""
        Dim sDirFac As String = "" : Dim sDirEnv As String = ""
        Dim sComent As String = "" : Dim sComentCab As String = "" : Dim sComentPie As String = ""
        Dim sCondicion As String = ""

        Dim sExiste As String = ""
        Dim bCrearCli As Boolean = False
        Dim iLinea As Integer = 0 : Dim sCodCampos As String = ""

        Dim sMensaje As String = ""
        Dim sCamposC(1, 3) As String : Dim sCamposL(1, 3) As String

        ' Apuntador libre a archivo
        Dim Apunt As Integer = FreeFile()
        ' Variable donde guardamos cada línea de texto
        Dim Texto As String = ""
        Dim sValorCampo As String = ""

        Dim sDocumento As String = "" : Dim sRef As String = ""
        Try
            'Tengo que buscar en la tabla el último numero de documento
            iDoc = objglobal.SQL.sqlNumericaB1("SELECT isnull(MAX(cast(CODE as int)),0) FROM ""@EXO_TMPDOC"" ")
            ' miramos si existe el fichero y cargamos
            If File.Exists(sArchivo) Then
                Using MyReader As New Microsoft.VisualBasic.
                        FileIO.TextFieldParser(sArchivo, System.Text.Encoding.UTF7)
                    MyReader.TextFieldType = FileIO.FieldType.Delimited
                    Select Case sDelimitador
                        Case "1" : MyReader.SetDelimiters(vbTab)
                        Case "2" : MyReader.SetDelimiters(";")
                        Case "3" : MyReader.SetDelimiters(",")
                        Case "4" : MyReader.SetDelimiters("-")
                        Case Else : MyReader.SetDelimiters(vbTab)
                    End Select

                    Dim currentRow As String()
                    While Not MyReader.EndOfData
                        Try
                            currentRow = MyReader.ReadFields()
                            Dim currentField As String
                            Dim scampos(1) As String
                            Dim iCampo As Integer = 0
                            For Each currentField In currentRow
                                iCampo += 1
                                ReDim Preserve scampos(iCampo)
                                scampos(iCampo) = currentField
                                'SboApp.MessageBox(scampos(iCampo))
                            Next
                            'Buscamos campos para traducir 
                            sSQL = "SELECT ""U_EXO_FEXCEL"",""U_EXO_CSAP"",""U_EXO_TDOC"" FROM ""@EXO_CFCNF"" "
                            sSQL &= " WHERE ""Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'"
                            oRs.DoQuery(sSQL)
                            If oRs.RecordCount > 0 Then
                                sTDoc = oRs.Fields.Item("U_EXO_TDOC").Value
                                If sTDoc = "1" Then
                                    sTDoc = "B"
                                Else
                                    sTDoc = "F"
                                End If
                                sCodCampos = oRs.Fields.Item("U_EXO_CSAP").Value
                                sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPC"" WHERE ""Code""='" & sCodCampos & "'"
                                oRCampos.DoQuery(sSQL)
                                If oRCampos.RecordCount > 0 Then
                                    'oSboApp.StatusBar.SetText("(EXO) - Leyendo Estructura de Cabecera...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Matrix de Cabecera"
                                    ReDim sCamposC(oRCampos.RecordCount, 3)
                                    For I = 1 To oRCampos.RecordCount
                                        sCamposC(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                        sCamposC(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                                        sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='C' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                                        sCampo = EXO_GLOBALES.GetValueDB(oCompany, """@EXO_FCCFL""", """U_EXO_posTXT""", sCondicion)
                                        If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                            ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                            sMensaje = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en el fichero TXT y es obligatorio." & ChrW(13) & ChrW(10)
                                            sMensaje &= "Por favor, Revise la parametrización."
                                            oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                        sCamposC(I, 3) = sCampo
                                        oRCampos.MoveNext()
                                    Next
#End Region
                                    'oSboApp.StatusBar.SetText("(EXO) - Estructura de Cabecera leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If

                                Dim sCuenta As String = "" : Dim sArt As String = "" : Dim sArtDes As String = ""
                                Dim sCantidad As String = "0.00" : Dim sprecio As String = "0.00" : Dim sDtoLin As String = "0.00" : Dim sTotalServicios As String = "0.00" : Dim sPrecioBruto As String = "0.00"
                                Dim sTextoAmpliado As String = "" : Dim sLinImpuestoCod As String = "" : Dim sLinRetCodigo As String = ""
                                sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPL"" WHERE ""Code""='" & sCodCampos & "'"
                                oRCampos.DoQuery(sSQL)
                                If oRCampos.RecordCount > 0 Then
                                    'oSboApp.StatusBar.SetText("(EXO) - Leyendo Estructura de líneas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
#Region "Matrix de Líneas"
                                    ReDim sCamposL(oRCampos.RecordCount, 3)
                                    For I = 1 To oRCampos.RecordCount
                                        sCamposL(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                        sCamposL(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                                        sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='L' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                                        sCampo = EXO_GLOBALES.GetValueDB(oCompany, """@EXO_FCCFL""", """U_EXO_posTXT""", sCondicion)
                                        If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                            ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                            sMensaje = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en el TXT y es obligatorio." & ChrW(13) & ChrW(10)
                                            sMensaje &= "Por favor, Revise la parametrización."
                                            oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox(sMensaje)
                                            Exit Sub
                                        End If
                                        sCamposL(I, 3) = sCampo
                                        oRCampos.MoveNext()
                                    Next
#End Region
                                    'oSboApp.StatusBar.SetText("(EXO) - Estructura de Líneas leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                End If
#Region "Lectura cabecera"
                                oSboApp.StatusBar.SetText("(EXO) - Leyendo Valores de Cabecera...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                For C = 1 To sCamposC.GetUpperBound(0)
                                    Select Case sCamposC(C, 1)
                                        Case "ObjType"
                                            sTFac = EXO_GLOBALES.Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sTFac = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocType"
                                            If sCamposC(C, 3) <> "" Then
                                                sTipoLineas = sCamposC(C, 3)
                                            End If
                                        Case "CardCode"
                                            sCliente = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sCliente = "" Then
                                                Exit Sub
                                            End If
                                            'Por lo que se ha visto no se da el Código de SAP, Por lo que buscaremos por el código de SAP, 
                                            'Sino existe por el CIF.  
                                            'Buscamos por el CODIGO DE SAP
                                            sExiste = ""
                                            sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """CardCode""='" & sCliente & "'")
                                            If sExiste = "" Then
                                                oSboApp.StatusBar.SetText("(EXO) - El Interlocutor  - " & sCliente & " - no se encuentra al buscarlo por el código de SAP. Se buscará por Nº identificación fiscal.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                                sExiste = ""
                                                'Suponemos que es el CIF, por lo que miramos a ver si tiene el país delante del CIF, y sino, ponemos el País España
                                                If IsNumeric(Left(sCliente, 2)) Then
                                                    sCliente = "ES" & sCliente
                                                End If
                                                sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """LicTradNum""='" & sCliente & "'")
                                                If sExiste = "" Then
                                                    oSboApp.StatusBar.SetText("(EXO) - El Interlocutor  - " & sCliente & " - no existe al buscarlo por Nº identificación fiscal. Se buscará por el campo ID Número 2.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """AddID""='" & sCliente & "'")
                                                    If sExiste = "" Then
                                                        oSboApp.StatusBar.SetText("(EXO) - El Interlocutor  - " & sCliente & " - no existe al buscarlo por ID Número 2.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        bCrearCli = True
                                                    Else
                                                        sCliente = sExiste
                                                    End If
                                                Else
                                                    sCliente = sExiste
                                                End If
                                            End If
                                        Case "CardName"
                                            sCliNombre = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sCliNombre = "" Then
                                                Exit Sub
                                            End If
                                        Case "ADDID"
                                            sCodCliente = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sCodCliente = "" Then
                                                Exit Sub
                                            End If
                                        Case "NumAtCard"
                                            sNumAtCard = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sNumAtCard = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_Manual"
                                            sManual = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sManual = "" Then
                                                Exit Sub
                                            End If
                                        Case "Series"
                                            sSerie = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sSerie = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocNum"
                                            sDocNum = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sDocNum = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocCurrency"
                                            sMoneda = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sMoneda = "" Then
                                                Exit Sub
                                            End If
                                        Case "SlpCode"
                                            sEmpleado = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sEmpleado = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocDate"
                                            sFContable = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sFContable = "" Then
                                                Exit Sub
                                            End If
                                        Case "TaxDate"
                                            sFDocumento = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sFDocumento = "" Then
                                                Exit Sub
                                            End If
                                        Case "DocDueDate"
                                            sFVto = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sFVto = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_TDTO"
                                            sTipoDto = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sTipoDto = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_DTO"
                                            sDto = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sDto = "" Then
                                                Exit Sub
                                            End If
                                        Case "PeyMethod"
                                            sPeyMethod = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sPeyMethod = "" Then
                                                Exit Sub
                                            End If
                                        Case "GroupNum"
                                            sCondPago = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sCondPago = "" Then
                                                Exit Sub
                                            End If
                                        Case "PayToCode"
                                            sDirFac = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sDirFac = "" Then
                                                Exit Sub
                                            ElseIf sDirFac = "" Then
                                                sDirFac = "Facturación"
                                            End If
                                        Case "ShipToCode"
                                            sDirEnv = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sDirEnv = "" Then
                                                Exit Sub
                                            ElseIf sDirEnv = "" Then
                                                sDirEnv = "Entrega"
                                            End If
                                        Case "Comments"
                                            sComent = "Importado a través del fichero - " & sArchivo & " - "
                                            sComent &= Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sComent = "Importado a través del fichero - " & sArchivo & " - " Then
                                                Exit Sub
                                            End If
                                        Case "OpeningRemarks"
                                            sComentCab = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sComentCab = "" Then
                                                Exit Sub
                                            End If
                                        Case "ClosingRemarks"
                                            sComentPie = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sComentPie = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_DIR"
                                            sDIR = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sDIR = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_POB"
                                            sPob = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sPob = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_PRO"
                                            sProv = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sProv = "" Then
                                                Exit Sub
                                            End If
                                        Case "EXO_CPOS"
                                            sCPos = Leer_Campo(sCamposC(C, 1), sCamposC(C, 3), sCamposC(C, 2), scampos, oSboApp)
                                            If sCamposC(C, 2) = "Y" And sCPos = "" Then
                                                Exit Sub
                                            End If
                                    End Select
                                Next
                                oSboApp.StatusBar.SetText("(EXO) - Valores de Cabecera leída.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
#Region "Comprobar datos cabecera"
                                If sTFac = "" Then
                                    Select Case oForm.TypeEx
                                        Case "EXO_CVFAC" : sTFac = "13" ' En el caso de no estar indicado, se ha tomado como Factura de venta
                                        Case "EXO_CCFAC" : sTFac = "18" ' En el caso de no estar indicado, se ha tomado como Factura de compras
                                        Case Else : sTFac = "13" ' En el caso de no estar indicado, se ha tomado como Factura de venta
                                    End Select
                                End If
                                If sTipoLineas = "" Then : sTipoLineas = "I" : End If ' En el caso de no estar indicado, se ha tomado como que son líneas de servicio
                                'Comprobamos que se haya introducido manual o con una serie
                                If sDocNum = "" Then
                                    If sSerie = "" Then
                                        sMensaje &= "No se ha indicado ni Nº de documento ni serie. Se indica el Nº de serie por defecto."
                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        'SboApp.MessageBox(sMensaje)
                                        sManual = "N"
                                        'Nº de serie
                                        sSerie = Year(Now).ToString("0000")
                                    Else
                                        sManual = "N"
                                    End If
                                Else
                                    sManual = "Y"
                                End If
                                If sMoneda = "" Then : sMoneda = "EUR" : End If 'En el caso de no estar indicado, se ha tomado EUR por defecto
                                If sFContable = "" Then
                                    sMensaje &= "No se ha indicado una fecha Contable para el documento. No se puede continuar."
                                    oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    oSboApp.MessageBox(sMensaje)
                                    Exit Sub
                                Else
                                    'Ponemos formato para SQL
                                    sFContable = Year(sFContable) & "-" & Month(sFContable) & "-" & Day(sFContable)
                                    If sFDocumento = "" Then
                                        sMensaje &= "No se ha indicado una fecha de documento. Se actualizará con la fecha contable."
                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        sFDocumento = sFContable
                                    Else
                                        sFDocumento = Year(sFDocumento) & "-" & Month(sFDocumento) & "-" & Day(sFDocumento)
                                    End If
                                End If
                                If sTipoDto = "" Then : sTipoDto = "%" : End If ' Se toma si no tiene valor que el dto va en Porcentaje
                                If sDto = "" Then : sDto = "0.00" : End If ' Se toma por defecto dto valor a 0.00

                                If bCrearCli = True Then
                                    'Creamos el cliente con los datos que nos han dado
                                    'Busco y compruebo que exista la serie que han marcado en los parametros por defecto depediento si es de venta o de compra
                                    Dim sSerieIC As String = Nothing : Dim sTipoIC As String = ""
                                    Select Case sTFac
                                        Case "13", "14"
                                            sSerieIC = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_SERIEV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            sTipoIC = "C"
                                        Case "18", "19", "22"
                                            sSerieIC = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_SERIEC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            sTipoIC = "S"
                                    End Select
                                    If sSerieIC = "" Then
                                        sMensaje &= "No se ha indicado la serie para crear el interlocutor. No se puede continuar. Revise la parametrización."
                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Exit Sub
                                    Else
                                        EXO_GLOBALES.CrearInterlocutorSencillo(CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString, sCliente,
                                                                  sSerieIC, sTipoIC, sCliNombre, "ES", sCliente, sCodCliente, "1", sCondPago, sPeyMethod,
                                                                  sDirFac, sDirEnv, sDIR, sPob, sProv, sCPos, "ES", oCompany, oSboApp, objglobal)
                                    End If
                                End If
                                oSboApp.StatusBar.SetText("(EXO) - Datos de cabecera comprobados.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region
                                'Grabamos la cabecera
                                'Insertar en la tabla temporal la cabecera
                                'Antes de insertar, comprobamos las direcciones de entrega y de facturación para comprobar que si son las de defecto del desarrollo, debemos buscar las de por defecto del cliente
                                If sDirFac = "Facturación" Then
                                    sSQL = "SELECT ""BillToDef"" FROM ""OCRD"" WHERE ""CardCode""='" & sCliente & "' "
                                    Dim sDIrFacDef As String = objglobal.SQL.sqlStringB1(sSQL)
                                    If sDIrFacDef <> "" Then
                                        sDirFac = sDIrFacDef
                                    End If
                                End If
                                If sDirEnv = "Entrega" Then
                                    sSQL = "SELECT ""ShipToDef"" FROM ""OCRD"" WHERE ""CardCode""='" & sCliente & "' "
                                    Dim sDirEnvDef As String = objglobal.SQL.sqlStringB1(sSQL)
                                    If sDirEnvDef <> "" Then
                                        sDirEnv = sDirEnvDef
                                    End If
                                End If

                                If sTFac <> "" Then
                                    If sDocumento <> sDocNum Or sRef <> sNumAtCard Then
                                        sDocumento = sDocNum
                                        sRef = sNumAtCard
                                        iDoc += 1
                                        iLinea = 0
                                        sSQL = "insert into ""@EXO_TMPDOC"" values('" & iDoc.ToString & "','" & iDoc.ToString & "'," & iDoc.ToString & ",'N','',0," & objglobal.conexionSAP.compañia.UserSignature
                                        sSQL &= ",'','',0,'',0,'','" & objglobal.usuario & "',"
                                        sSQL &= "'" & sTDoc & "','" & sDocNum & "','" & sTFac & "','" & sManual & "','" & sSerie & "','" & sNumAtCard & "','" & sMoneda & "','','" & sEmpleado & "',"
                                        sSQL &= "'" & sCliente & "','" & sCodCliente & "','" & sFContable & "','" & sFDocumento & "','" & sFVto & "','" & sTipoDto & "',"
                                        sSQL &= EXO_GLOBALES.DblNumberToText(objglobal, sDto.ToString) & ",'" & sPeyMethod & "','" & sDirFac & "','" & sDirEnv & "','" & sComent.Replace("'", "") & "','"
                                        sSQL &= sComentCab.Replace("'", "") & "','" & sComentPie.Replace("'", "") & "','" & sCondPago & "') "
                                        oRs.DoQuery(sSQL)
                                    Else
                                        iLinea += 1
                                    End If

#Region "Lectura de Líneas"
                                    oSboApp.StatusBar.SetText("(EXO) - Leyendo Valores de Líneas...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                    'Ahora la Línea
                                    For L = 1 To sCamposL.GetUpperBound(0)
                                        Select Case sCamposL(L, 1)
                                            Case "AcctCode"
                                                sCuenta = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sCuenta = "" Then
                                                    Exit Sub
                                                End If
                                            Case "ItemCode"
                                                sArt = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sArt = "" Then
                                                    Exit Sub
                                                End If
                                            Case "Dscription"
                                                sArtDes = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sArtDes = "" Then
                                                    Exit Sub
                                                End If
                                            Case "Quantity"
                                                sCantidad = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sCantidad = "" Then
                                                    Exit Sub
                                                End If
                                            Case "UnitPrice"
                                                sprecio = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sprecio = "" Then
                                                    Exit Sub
                                                End If
                                            Case "DiscPrcnt"
                                                sDtoLin = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sDtoLin = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_IMPSRV"
                                                sTotalServicios = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sTotalServicios = "" Then
                                                    Exit Sub
                                                End If
                                                oSboApp.StatusBar.SetText("(EXO) - " & sTotalServicios, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            Case "EXO_TextoLin"
                                                sTextoAmpliado = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sTextoAmpliado = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_IMP"
                                                sLinImpuestoCod = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sLinImpuestoCod = "" Then
                                                    Exit Sub
                                                End If
                                            Case "EXO_RET"
                                                sLinRetCodigo = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sLinRetCodigo = "" Then
                                                    Exit Sub
                                                End If
                                            Case "GrossBuyPr"
                                                sPrecioBruto = Leer_Campo(sCamposL(L, 1), sCamposL(L, 3), sCamposL(L, 2), scampos, oSboApp)
                                                If sCamposL(L, 2) = "Y" And sPrecioBruto = "" Then
                                                    Exit Sub
                                                End If
                                        End Select
                                    Next
                                    oSboApp.StatusBar.SetText("(EXO) - Valores de líneas leídos.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
#End Region

#Region "Comprobar datos línea"
                                    'Comprobamos que exista la cuenta                                  
                                    If sCuenta <> "" Then
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OACT""", """AcctCode""", """AcctCode""='" & sCuenta & "'")
                                        If sExiste = "" Then
                                            oSboApp.StatusBar.SetText("(EXO) - La Cuenta contable SAP  - " & sCuenta & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox("La Cuenta contable SAP - " & sCuenta & " - no existe.")
                                            Exit Sub
                                        End If
                                    End If
                                    'Comprobamos que exista el artículo
                                    If sTipoLineas = "I" Then
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OITM""", """ItemCode""", """ItemCode"" like '" & sArt & "'")
                                        If sExiste = "" Then
                                            oSboApp.StatusBar.SetText("(EXO) - El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox("El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.")
                                            Exit Sub
                                        End If
                                    ElseIf sTipoLineas = "S" Then
                                        If sCuenta = "" Then
                                            ' No puede estar la cuenta vacía si es de tipo servicio
                                            sExiste = ""
                                            sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CSRV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            If sExiste = "" Then
                                                sMensaje = " La cuenta en la línea del servicio no puede estar vacía. Por favor, Revise los datos de la parametrización."
                                                oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                oSboApp.MessageBox(sMensaje)
                                                Exit Sub
                                            Else
                                                sCuenta = sExiste
                                            End If
                                        End If
                                    End If
                                    'Comprobamos que exista el impuesto si está relleno
                                    If sLinImpuestoCod = "" Then
                                        Select Case sTFac
                                            Case "13", "14" 'Ventas
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                            Case "18", "19", "22" 'Compras
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_IVAC""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                        End Select
                                    Else
                                        sLinImpuestoCod = sLinImpuestoCod.Replace(",", ".")
                                        Select Case sTFac
                                            Case "13", "14" 'Ventas
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OVTG""", """Code""", """Rate""='" & sLinImpuestoCod & "' and  LENGTH(""Code"")=2 and left(""Code"",1)='R' and ""Category""='O' ")
                                            Case "18", "19", "22" 'Compras
                                                sLinImpuestoCod = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OVTG""", """Code""", """Rate""='" & sLinImpuestoCod & "' and  LENGTH(""Code"")=2 and left(""Code"",1)='S' and ""Category""='I' ")
                                        End Select
                                    End If
                                    If sLinImpuestoCod <> "" Then
                                        sExiste = ""
                                        sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OVTG""", """Code""", """Code""='" & sLinImpuestoCod & "'")
                                        If sExiste = "" Then
                                            oSboApp.StatusBar.SetText("(EXO) - El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox("El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.")
                                            Exit Sub
                                        End If
                                    End If
                                    'Comprobamos que exista la retención si está relleno
                                    If sLinRetCodigo <> "" Then
                                        sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """CRD4""", """WTCode""", """CardCode""='" & sCliente & "' and ""WTCode""='" & sLinRetCodigo & "'")
                                        If sExiste = "" Then
                                            oSboApp.StatusBar.SetText("(EXO) - El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oSboApp.MessageBox("El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.")
                                            Exit Sub
                                        End If
                                    End If
#End Region
                                    'Grabamos la línea
                                    sSQL = "insert into ""@EXO_TMPDOCL"" values('" & iDoc.ToString & "','" & iLinea & "','',0,'" & objglobal.usuario & "',"
                                    sSQL &= "'" & sCuenta & "','" & sArt & "','" & sArtDes & "'," & EXO_GLOBALES.DblNumberToText(objglobal, sCantidad.ToString).Replace(",", ".") & ","
                                    sSQL &= EXO_GLOBALES.DblNumberToText(objglobal, sprecio.ToString) & "," & EXO_GLOBALES.DblNumberToText(objglobal, sDtoLin.ToString)
                                    sSQL &= "," & EXO_GLOBALES.DblNumberToText(objglobal, sTotalServicios.ToString).Replace(",", ".") & ",'" & sLinImpuestoCod & "','" & sLinRetCodigo & "','"
                                    sSQL &= sTextoAmpliado & "','" & sTipoLineas & "'," & sPrecioBruto & " ) "
                                    oRs.DoQuery(sSQL)
                                End If
                            End If
                        Catch ex As Microsoft.VisualBasic.
                            FileIO.MalformedLineException
                            objglobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Línea " & ex.Message & " no es válida y se omitirá.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            oSboApp.MessageBox("Línea " & ex.Message & " no es válida y se omitirá.")
                        End Try
                    End While
                End Using
            Else
                objglobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - No se ha encontrado el fichero txt a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            ' Cerramos el archivo
            FileClose(Apunt)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
        End Try
    End Sub
    Public Shared Sub TratarFichero_Excel(ByVal sArchivo As String, ByRef oForm As SAPbouiCOM.Form, ByRef oCompany As SAPbobsCOM.Company, ByRef oSboApp As Application, ByRef objglobal As EXO_Generales.EXO_General)
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oRCampos As SAPbobsCOM.Recordset = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sCampo As String = ""

        Dim iDoc As Integer = 0 'Contador de Cabecera de documentos
        Dim sTFac As String = "" : Dim sTFacColumna As String = "" : Dim sTipoLineas As String = "" : Dim sTDoc As String = ""
        Dim sCliente As String = "" : Dim sCliNombre As String = "" : Dim sCodCliente As String = "" : Dim sClienteColumna As String = "" : Dim sCodClienteColumna As String = ""
        Dim sSerie As String = "" : Dim sDocNum As String = "" : Dim sManual As String = "" : Dim sSerieColumna As String = "" : Dim sDocNumColumna As String = ""
        Dim sNumAtCard As String = "" : Dim sNumAtCardColumna As String = ""
        Dim sMoneda As String = "" : Dim sMonedaColumna As String = ""
        Dim sEmpleado As String = ""
        Dim sFContable As String = "" : Dim sFDocumento As String = "" : Dim sFVto As String = "" : Dim sFDocumentoColumna As String = ""
        Dim sTipoDto As String = "" : Dim sDto As String = ""
        Dim sPeyMethod As String = "" : Dim sCondPago As String = ""
        Dim sDirFac As String = "" : Dim sDirEnv As String = ""
        Dim sComent As String = "" : Dim sComentCab As String = "" : Dim sComentPie As String = ""
        Dim sCondicion As String = ""

        Dim sExiste As String = ""
        Dim iLinea As Integer = 0 : Dim sCodCampos As String = ""

        Dim pck As ExcelPackage = Nothing
        Dim iLin As Integer = 0
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
                    sTDoc = oRs.Fields.Item("U_EXO_TDOC").Value
                    If sTDoc = "1" Then
                        sTDoc = "B"
                    Else
                        sTDoc = "F"
                    End If
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
                            sCampo = EXO_GLOBALES.GetValueDB(oCompany, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
                            If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
                                sMensaje &= "Por favor, Revise la parametrización."
                                oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                oSboApp.MessageBox(sMensaje)
                                Exit Sub
                            End If
                            sCamposC(I, 3) = sCampo
                            Select Case oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                Case "ObjType" : sTFacColumna = sCampo
                                Case "CardCode" : sClienteColumna = sCampo
                                Case "ADDID" : sCodClienteColumna = sCampo
                                Case "Series" : sSerieColumna = sCampo
                                Case "DocNum" : sDocNumColumna = sCampo
                                Case "NumAtCard" : sNumAtCardColumna = sCampo
                                Case "DocCurrency" : sMonedaColumna = sCampo
                                Case "TaxDate" : sFDocumentoColumna = sCampo
                            End Select
                            oRCampos.MoveNext()
                        Next
#End Region
                        Do
#Region "Cabecera"
                            If sTFac <> worksheet.Cells(sTFacColumna & iLin).Text Or sCliente <> worksheet.Cells(sClienteColumna & iLin).Text Or sCodCliente <> worksheet.Cells(sCodClienteColumna & iLin).Text _
                                Or sSerie <> worksheet.Cells(sSerieColumna & iLin).Text Or sDocNum <> worksheet.Cells(sDocNumColumna & iLin).Text Or sNumAtCard <> worksheet.Cells(sNumAtCardColumna & iLin).Text _
                                Or sMoneda <> worksheet.Cells(sMonedaColumna & iLin).Text Or sFDocumento <> worksheet.Cells(sFDocumentoColumna & iLin).Text Then
                                'Grabamos la cabecera
                                For C = 1 To sCamposC.GetUpperBound(0)
                                    Select Case sCamposC(C, 1)
                                        Case "ObjType"
#Region "ObjType"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        If worksheet.Cells("A" & iLin).Text <> "" Then
                                                            Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                            sMensaje &= "Por favor, Revise el documento Excel."
                                                            oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                            oSboApp.MessageBox(sMensaje)
                                                            Exit Sub
                                                        Else
                                                            Exit Do
                                                        End If
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sTFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sTFac = ""
                                            End If
#End Region
                                        Case "DocType"
#Region "DocType"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sTipoLineas = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sTipoLineas = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sTipoLineas = ""
                                            End If
#End Region
                                        Case "CardCode"
#Region "CardCode"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        'SboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sCliente = ""
                                            End If

#End Region
                                        Case "CardName"
#Region "CardName"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCliNombre = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sCliNombre = ""
                                            End If

#End Region
                                        Case "ADDID"
#Region "ADDID"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        'SboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCodCliente = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sCodCliente = ""
                                            End If
#End Region
                                        Case "NumAtCard"
#Region "NumAtCard"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sNumAtCard = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sNumAtCard = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sNumAtCard = ""
                                            End If

#End Region
                                        Case "EXO_Manual"
#Region "EXO_Manual"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sManual = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sManual = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sManual = ""
                                            End If
#End Region
                                        Case "Series"
#Region "Series"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sSerie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sSerie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sSerie = ""
                                            End If
#End Region
                                        Case "DocNum"
#Region "DocNum"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDocNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDocNum = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sDocNum = ""
                                            End If
#End Region
                                        Case "DocCurrency"
#Region "DocCurrency"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sMoneda = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sMoneda = "EUR"
                                            End If
#End Region
                                        Case "SlpCode"
#Region "SlpCode"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sEmpleado = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sEmpleado = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sEmpleado = ""
                                            End If

#End Region
                                        Case "DocDate"
#Region "DocDate"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sFContable = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                        If sFContable <> "" Then
                                                            sFContable = Year(sFContable).ToString("0000") & Month(sFContable).ToString("00") & Day(sFContable).ToString("00")
                                                        End If
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sFContable = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                        If sFContable <> "" Then
                                                            sFContable = Year(sFContable).ToString("0000") & Month(sFContable).ToString("00") & Day(sFContable).ToString("00")
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                sFContable = ""
                                            End If
#End Region
                                        Case "TaxDate"
#Region "TaxDate"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sFDocumento = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                        If sFDocumento <> "" Then
                                                            sFDocumento = Year(sFDocumento).ToString("0000") & Month(sFDocumento).ToString("00") & Day(sFDocumento).ToString("00")
                                                        End If
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sFDocumento = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                        If sFDocumento <> "" Then
                                                            sFDocumento = Year(sFDocumento).ToString("0000") & Month(sFDocumento).ToString("00") & Day(sFDocumento).ToString("00")
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                sFDocumento = ""
                                            End If
#End Region
                                        Case "DocDueDate"
#Region "DocDueDate"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sFVto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                        If sFVto <> "" Then
                                                            sFVto = Year(sFVto).ToString("0000") & Month(sFVto).ToString("00") & Day(sFVto).ToString("00")
                                                        End If
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sFVto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                        If sFVto <> "" Then
                                                            sFVto = Year(sFVto).ToString("0000") & Month(sFVto).ToString("00") & Day(sFVto).ToString("00")
                                                        Else
                                                            sFVto = ""
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                sFVto = ""
                                            End If
#End Region
                                        Case "EXO_TDTO"
#Region "EXO_TDTO"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sTipoDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sTipoDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sTipoDto = ""
                                            End If
#End Region
                                        Case "EXO_DTO"
#Region "EXO_DTO"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDto = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sDto = "0.00"
                                            End If

#End Region
                                        Case "PeyMethod"
#Region "PeyMethod"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sPeyMethod = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sPeyMethod = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sPeyMethod = ""
                                            End If

#End Region
                                        Case "GroupNum"
#Region "GroupNum"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCondPago = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sCondPago = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sCondPago = ""
                                            End If
#End Region
                                        Case "PayToCode"
#Region "PayToCode"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDirFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDirFac = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sDirFac = ""
                                            End If

#End Region
                                        Case "ShipToCode"
#Region "ShipToCode"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDirEnv = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sDirEnv = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sDirEnv = ""
                                            End If
#End Region
                                        Case "Comments"
#Region "Comments"
                                            Dim iPosicion As Integer = InStr(sArchivo, "08.Historico")
                                            Dim sMiDir As String = Mid(sArchivo, iPosicion)
                                            sComent = "Importado a través del fichero - " & sMiDir & " - "
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sComent &= worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sComent &= worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            End If
#End Region
                                        Case "OpeningRemarks"
#Region "OpeningRemarks"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sComentCab = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sComentCab = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sComentCab = ""
                                            End If

#End Region
                                        Case "ClosingRemarks"
#Region "ClosingRemarks"
                                            If sCamposC(C, 3) <> "" Then
                                                If sCamposC(C, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sComentPie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposC(C, 1) & """ es obligatorio y la columna """ & sCamposC(C, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposC(C, 3) & iLin).Text <> "" Then
                                                        sComentPie = worksheet.Cells(sCamposC(C, 3) & iLin).Text
                                                    End If
                                                End If
                                            Else
                                                sComentPie = ""
                                            End If
#End Region
                                    End Select
                                Next
                                'Grabamos la cabecera
                                iLinea = 0
                                If sTFac = "" Then
                                    sTFac = Left(sCliente, 2)
                                End If
                                Dim sExisteIC As String = ""
                                Select Case Left(sCodCliente, 2)
                                    Case "43"
                                        sExisteIC = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """LicTradNum"" like '%" & sCliente & "%' and ""CardType""='C'")
                                        If sExisteIC = "" Then
                                            'sCliente = "C" & Mid(sCodCliente, 4)
                                            sExisteIC = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """AddID""='" & sCodCliente & "' and ""CardType""='C'")
                                            sCliente = sExisteIC
                                        Else
                                            sCliente = sExisteIC
                                        End If
                                    Case "40"
                                        sExisteIC = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """LicTradNum"" like '%" & sCliente & "%' and ""CardType""='S'")
                                        If sExisteIC = "" Then
                                            'sCliente = "P" & Mid(sCodCliente, 4)
                                            sExisteIC = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OCRD""", """CardCode""", """AddID""='" & sCodCliente & "' and ""CardType""='S'")
                                            sCliente = sExisteIC
                                        Else
                                            sCliente = sExisteIC
                                        End If
                                End Select
                                'Insertar en la tabla temporal la cabecera
                                If sTFac <> "" Then
                                    iDoc += 1
                                    sSQL = "insert into ""@EXO_TMPDOC"" values('" & iDoc.ToString & "','" & iDoc.ToString & "'," & iDoc.ToString & ",'N','',0," & objglobal.conexionSAP.compañia.UserSignature
                                            sSQL &= ",'','',0,'',0,'','" & objglobal.usuario & "',"
                                    sSQL &= "'" & sTDoc & "','" & sDocNum & "','" & sTFac & "','" & sManual & "','" & sSerie & "','" & sNumAtCard & "','" & sMoneda & "','','" & sEmpleado & "',"
                                    sSQL &= "'" & sCliente & "','" & sCodCliente & "','" & sFContable & "','" & sFDocumento & "','" & sFVto & "','" & sTipoDto & "',"
                                    sSQL &= EXO_GLOBALES.DblNumberToText(objglobal, sDto.ToString) & ",'" & sPeyMethod & "','" & sDirFac & "','" & sDirEnv & "','" & sComent.Replace("'", "") & "','"
                                    sSQL &= sComentCab.Replace("'", "") & "','" & sComentPie.Replace("'", "") & "','" & sCondPago & "') "
                                    oRs.DoQuery(sSQL)
                                End If
                            End If
#End Region
                            'Ahora tratamos la línea
#Region "Líneas"
                            Dim sCuenta As String = "" : Dim sArt As String = "" : Dim sArtDes As String = ""
                            Dim sCantidad As String = "0.00" : Dim sprecio As String = "0.00" : Dim sDtoLin As String = "0.00" : Dim sTotalServicios As String = "0.00" : Dim sPrecioBruto As String = "0.00"
                            Dim sTextoAmpliado As String = "" : Dim sLinImpuestoCod As String = "" : Dim sLinRetCodigo As String = "" : Dim sReparto As String = ""
                            sSQL = "SELECT ""U_EXO_COD"",""U_EXO_OBL"" FROM ""@EXO_CSAPL"" WHERE ""Code""='" & sCodCampos & "'"
                            oRCampos.DoQuery(sSQL)
                            If oRCampos.RecordCount > 0 Then
#Region "Matrix de Líneas"
                                Dim sCamposL(oRCampos.RecordCount, 3) As String
                                For I = 1 To oRCampos.RecordCount
                                    sCamposL(I, 1) = oRCampos.Fields.Item("U_EXO_COD").Value.ToString
                                    sCamposL(I, 2) = oRCampos.Fields.Item("U_EXO_OBL").Value.ToString
                                    sCondicion = """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "' and ""U_EXO_TIPO""='L' And ""U_EXO_COD""='" & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & "' "
                                    sCampo = EXO_GLOBALES.GetValueDB(oCompany, """@EXO_FCCFL""", """U_EXO_posExcel""", sCondicion)
                                    If sCampo = "" And oRCampos.Fields.Item("U_EXO_OBL").Value.ToString = "Y" Then
                                        ' Si es obligatorio y no se ha indicado para leer hay que dar un error
                                        Dim sMensaje As String = "El campo """ & oRCampos.Fields.Item("U_EXO_COD").Value.ToString & """ no está asignado en la hoja de Excel y es obligatorio." & ChrW(13) & ChrW(10)
                                        sMensaje &= "Por favor, Revise la parametrización."
                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oSboApp.MessageBox(sMensaje)
                                        Exit Sub
                                    End If
                                    sCamposL(I, 3) = sCampo
                                    oRCampos.MoveNext()
                                Next
#End Region
                                For L = 1 To sCamposL.GetUpperBound(0)
                                    Select Case sCamposL(L, 1)
                                        Case "AcctCode"
#Region "AcctCode"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sCuenta = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sCuenta = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sCuenta = ""
                                                    End If
                                                End If
                                            Else
                                                sCuenta = ""
                                            End If

#End Region
                                        Case "ItemCode"
#Region "ItemCode"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sArt = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sArt = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sArt = ""
                                                    End If
                                                End If
                                            Else
                                                sArt = ""
                                            End If
#End Region
                                        Case "Dscription"
#Region "Dscription"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sArtDes = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sArtDes = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sArtDes = ""
                                                    End If
                                                End If
                                            Else
                                                sArtDes = ""
                                            End If
#End Region
                                        Case "Quantity"
#Region "Quantity"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sCantidad = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sCantidad = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sCantidad = "0.00"
                                                    End If
                                                End If
                                            Else
                                                sCantidad = "0.00"
                                            End If

#End Region
                                        Case "UnitPrice"
#Region "UnitPrice"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sprecio = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sprecio = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sprecio = "0.00"
                                                    End If
                                                End If
                                            Else
                                                sprecio = "0.00"
                                            End If

#End Region
                                        Case "DiscPrcnt"
#Region "DiscPrcnt"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sDtoLin = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sDtoLin = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sDtoLin = "0.00"
                                                    End If
                                                End If
                                            Else
                                                sDtoLin = "0.00"
                                            End If

#End Region
                                        Case "EXO_IMPSRV"
#Region "EXO_IMPSRV"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        Dim sLet1 As String = sCamposL(L, 3) : Dim iLet1 As Integer = Asc(sLet1) + 1 : sLet1 = ChrW(iLet1)
                                                        Dim sLet2 As String = sCamposL(L, 3) : Dim iLet2 As Integer = Asc(sLet2) + 2 : sLet2 = ChrW(iLet2)
                                                        sTotalServicios = (CDbl(worksheet.Cells(sCamposL(L, 3) & iLin).Text) + CDbl(worksheet.Cells(sLet1 & iLin).Text) + CDbl(worksheet.Cells(sLet2 & iLin).Text)).ToString.Replace(",", ".")
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        Dim sLet1 As String = sCamposL(L, 3) : Dim iLet1 As Integer = Asc(sLet1) + 1 : sLet1 = ChrW(iLet1)
                                                        Dim sLet2 As String = sCamposL(L, 3) : Dim iLet2 As Integer = Asc(sLet2) + 2 : sLet2 = ChrW(iLet2)
                                                        sTotalServicios = (CDbl(worksheet.Cells(sCamposL(L, 3) & iLin).Text) + CDbl(worksheet.Cells(sLet1 & iLin).Text) + CDbl(worksheet.Cells(sLet2 & iLin).Text)).ToString.Replace(",", ".")
                                                    Else
                                                        sTotalServicios = "0.00"
                                                    End If
                                                End If
                                            Else
                                                sTotalServicios = "0.00"
                                            End If

#End Region
                                        Case "EXO_TextoLin"
#Region "EXO_TextoLin"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sTextoAmpliado = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sTextoAmpliado = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sTextoAmpliado = ""
                                                    End If
                                                End If
                                            Else
                                                sTextoAmpliado = ""
                                            End If

#End Region
                                        Case "EXO_IMP"
#Region "EXO_IMP"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sLinImpuestoCod = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sLinImpuestoCod = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sLinImpuestoCod = ""
                                                    End If
                                                End If
                                            Else
                                                sLinImpuestoCod = ""
                                            End If

#End Region
                                        Case "EXO_RET"
#Region "EXO_RET"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sLinRetCodigo = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sLinRetCodigo = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sLinRetCodigo = ""
                                                    End If
                                                End If
                                            Else
                                                sLinRetCodigo = ""
                                            End If

#End Region
                                        Case "GrossBuyPr"
#Region "GrossBuyPr"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sPrecioBruto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sPrecioBruto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sPrecioBruto = "0.00"
                                                    End If
                                                End If
                                            Else
                                                sPrecioBruto = "0.00"
                                            End If

#End Region
                                        Case "EXO_REPARTO"
#Region "EXO_REPARTO"
                                            If sCamposL(L, 3) <> "" Then
                                                If sCamposL(L, 2) = "Y" Then
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sReparto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        Dim sMensaje As String = "El campo """ & sCamposL(L, 1) & """ es obligatorio y la columna """ & sCamposL(L, 3) & """ está vacía." & ChrW(13) & ChrW(10)
                                                        sMensaje &= "Por favor, Revise el documento Excel."
                                                        oSboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                        oSboApp.MessageBox(sMensaje)
                                                        Exit Sub
                                                    End If
                                                Else
                                                    If worksheet.Cells(sCamposL(L, 3) & iLin).Text <> "" Then
                                                        sReparto = worksheet.Cells(sCamposL(L, 3) & iLin).Text
                                                    Else
                                                        sReparto = ""
                                                    End If
                                                End If
                                            Else
                                                sReparto = ""
                                            End If

#End Region
                                    End Select
                                Next
#Region "Comprobar datos línea"
                                'Comprobamos que exista la cuenta
                                If sCuenta <> "" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OACT""", """AcctCode""", """AcctCode""='" & sCuenta & "'")
                                    If sExiste = "" Then
                                        oSboApp.StatusBar.SetText("(EXO) - La Cuenta contable SAP  - " & sCuenta & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oSboApp.MessageBox("La Cuenta contable SAP - " & sCuenta & " - no existe.")
                                        Exit Sub
                                    End If
                                End If
                                If sTipoLineas = "" And sArt = "" Then
                                    sTipoLineas = "S" ' En el caso de no estar indicado, se ha tomado como que son líneas de servicio
                                ElseIf sArt <> "" Then
                                    sTipoLineas = "I" ' En el caso de no estar indicado, se ha tomado como que son líneas de artículos
                                End If
                                'Comprobamos que exista el artículo
                                If sTipoLineas = "I" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OITM""", """ItemCode""", """ItemCode"" like '" & sArt & "'")
                                    If sExiste = "" Then
                                        oSboApp.StatusBar.SetText("(EXO) - El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oSboApp.MessageBox("El Artículo - " & sArt & " - " & sArtDes & " no existe. Borrelo de la sección concepto para poderlo crear automáticamente.")
                                        Exit Sub
                                    End If
                                ElseIf sTipoLineas = "S" Then
                                    If sCuenta = "" Then
                                        ' No puede estar la cuenta vacía si es de tipo servicio
                                        'Dim sMensaje As String = " La cuenta esta vacía. Se coge los datos por defecto."
                                        'SboApp.StatusBar.SetText("(EXO) - " & sMensaje, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        sCuenta = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """@EXO_CFCNF""", """U_EXO_CSRV""", """Code""='" & CType(oForm.Items.Item("cb_Format").Specific, SAPbouiCOM.ComboBox).Selected.Value.ToString & "'")
                                    End If
                                End If
                                'Comprobamos que exista el impuesto si está relleno
                                If sLinImpuestoCod <> "" Then
                                    sExiste = ""
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """OVTG""", """Code""", """Code""='" & sLinImpuestoCod & "'")
                                    If sExiste = "" Then
                                        oSboApp.StatusBar.SetText("(EXO) - El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oSboApp.MessageBox("El Grupo Impositivo  - " & sLinImpuestoCod & " - no existe.")
                                        Exit Sub
                                    End If
                                End If
                                'Comprobamos que exista la retención si está relleno
                                If sLinRetCodigo <> "" Then
                                    sExiste = EXO_GLOBALES.GetValueDB(objglobal.conexionSAP.compañia, """CRD4""", """WTCode""", """CardCode""='" & sCliente & "' and ""WTCode""='" & sLinRetCodigo & "'")
                                    If sExiste = "" Then
                                        oSboApp.StatusBar.SetText("(EXO) - El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        oSboApp.MessageBox("El indicador de Retención  - " & sLinRetCodigo & " - no no está marcado para el interlocutor " & sCliente & ". Por favor, revise el interlocutor.")
                                        Exit Sub
                                    End If
                                End If
#End Region
                                'Grabamos la línea
                                If EXO_GLOBALES.DblNumberToText(objglobal, sTotalServicios.ToString) > 0 Then
                                    'Actualizamos la cabecera para que sea factura
                                    If Left(sCodCliente, 2) = "43" Then
                                        sSQL = "UPDATE ""@EXO_TMPDOC"" SET ""U_EXO_TIPOF""='13' Where ""Code""='" & iDoc.ToString & "' "
                                    Else
                                        sSQL = "UPDATE ""@EXO_TMPDOC"" SET ""U_EXO_TIPOF""='18' Where ""Code""='" & iDoc.ToString & "' "
                                    End If
                                    oRs.DoQuery(sSQL)
                                Else
                                    'Actualizamos la cabecera para que sea Abono
                                    If Left(sCodCliente, 2) = "43" Then
                                        sSQL = "UPDATE ""@EXO_TMPDOC"" SET ""U_EXO_TIPOF""='14' Where ""Code""='" & iDoc.ToString & "' "
                                    Else
                                        sSQL = "UPDATE ""@EXO_TMPDOC"" SET ""U_EXO_TIPOF""='19' Where ""Code""='" & iDoc.ToString & "' "
                                    End If
                                    oRs.DoQuery(sSQL)
                                    'Actualizamos el importe
                                    sTotalServicios = sTotalServicios.Replace("-", "")
                                End If
                                If sTipoLineas = "S" Then
                                    'Escribimos el texto para el servicio
                                    sTextoAmpliado = "FACTURA SERVICIO"
                                End If
                                sSQL = "insert into ""@EXO_TMPDOCL"" values('" & iDoc.ToString & "','" & iLinea & "','',0,'" & objglobal.usuario & "',"
                                sSQL &= "'" & sCuenta & "','" & sArt & "','" & sArtDes & "'," & EXO_GLOBALES.DblNumberToText(objglobal, sCantidad.ToString).Replace(",", ".") & ","
                                sSQL &= EXO_GLOBALES.DblNumberToText(objglobal, sprecio.ToString) & "," & EXO_GLOBALES.DblNumberToText(objglobal, sDtoLin.ToString)
                                sSQL &= "," & EXO_GLOBALES.DblNumberToText(objglobal, sTotalServicios.ToString).Replace(",", ".") & ",'" & sLinImpuestoCod & "','" & sLinRetCodigo & "','"
                                sSQL &= sTextoAmpliado & "','" & sTipoLineas & "'," & sPrecioBruto & ",'" & sReparto & "' ) "
                                oRs.DoQuery(sSQL)
                                iLin += 1 : iLinea += 1
                            End If
#End Region
                        Loop While sTFac <> ""
                    End If


                Else
                    oSboApp.StatusBar.SetText("(EXO) - Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    oSboApp.MessageBox("Error inesperado. No se ha encontrado la configuración de lectura del fichero de excel. No se puede cargar el fichero.")
                End If
            Else
                objglobal.conexionSAP.SBOApp.StatusBar.SetText("No se ha encontrado el fichero excel a cargar.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If


        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            pck.Dispose()
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRCampos, Object))
        End Try
    End Sub
#End Region
End Class

