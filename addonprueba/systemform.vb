Imports System.Windows.Forms
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Windows.Forms
Imports CrystalDecisions.ReportSource
Imports System.IO
Imports System.Text
Public Class systemform

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Private esFactura As Boolean
    Private IdForm As String
    Private IdItem As String
    Private IdEvent As Integer
    Private IdAction As Boolean = False
    Private oForm As SAPbouiCOM.Form
    Private txtRecibo As SAPbouiCOM.EditText
    Private cmbEstado As SAPbouiCOM.ComboBox
    Private txtCodFinca As SAPbouiCOM.EditText
    Private txtTipoOT As SAPbouiCOM.EditText
    Private txtFinca As SAPbouiCOM.EditText
    Private cmbCatado As SAPbouiCOM.ComboBox
    Private txtCodCafe As SAPbouiCOM.EditText
    Private txtCantidad As SAPbouiCOM.EditText
    Private txtTipo As SAPbouiCOM.EditText
    Private txtEscala As SAPbouiCOM.EditText
    Private txtEscalaRechazo As SAPbouiCOM.EditText
    Private cmbVerde As SAPbouiCOM.ComboBox
    Private cmbTueste As SAPbouiCOM.ComboBox
    Private txtComVerde As SAPbouiCOM.EditText
    Private txtComTueste As SAPbouiCOM.EditText
    Private chkTzSana As SAPbouiCOM.CheckBox
    Private chkTzFermentada As SAPbouiCOM.CheckBox
    Private chkTzMohosa As SAPbouiCOM.CheckBox
    Private chkTzFruti As SAPbouiCOM.CheckBox
    Private chkTzTerrosa As SAPbouiCOM.CheckBox
    Private chkTzFenolica As SAPbouiCOM.CheckBox
    Private chkTzAgria As SAPbouiCOM.CheckBox
    Private chkTzVinosa As SAPbouiCOM.CheckBox
    Private chkTzAspera As SAPbouiCOM.CheckBox
    Private chkTzSucia As SAPbouiCOM.CheckBox
    Private txtComTaza As SAPbouiCOM.EditText
    Private cmbPargo As SAPbouiCOM.ComboBox
    Private txtComPargo As SAPbouiCOM.EditText
    Private txtHumedad As SAPbouiCOM.EditText
    Private txtRendBruto As SAPbouiCOM.EditText
    Private txtRendNeto As SAPbouiCOM.EditText
    Private txtGrsLtrs As SAPbouiCOM.EditText
    Private txtRendNetoE As SAPbouiCOM.EditText
    Private txtRO As SAPbouiCOM.EditText
    Private txtESC As SAPbouiCOM.EditText
    Private txtBZ As SAPbouiCOM.EditText
    Private txtSZ As SAPbouiCOM.EditText
    Private txtScore As SAPbouiCOM.EditText
    Private txtEntrada As SAPbouiCOM.EditText
    Private txtTotal As SAPbouiCOM.EditText
    Private txtObservaciones1 As SAPbouiCOM.EditText
    Private txtObservaciones2 As SAPbouiCOM.EditText
    Private txtObservaciones3 As SAPbouiCOM.EditText
    Private tblOrdenTrabajo As SAPbouiCOM.Matrix
    Private txtCantZar13 As SAPbouiCOM.EditText
    Private txtCantZar14 As SAPbouiCOM.EditText
    Private txtCantZar15 As SAPbouiCOM.EditText
    Private txtCantZar16 As SAPbouiCOM.EditText
    Private txtCantZar17 As SAPbouiCOM.EditText
    Private txtCantZar18 As SAPbouiCOM.EditText
    Private txtPorcZar13 As SAPbouiCOM.EditText
    Private txtPorcZar14 As SAPbouiCOM.EditText
    Private txtPorcZar15 As SAPbouiCOM.EditText
    Private txtPorcZar16 As SAPbouiCOM.EditText
    Private txtPorcZar17 As SAPbouiCOM.EditText
    Private txtPorcZar18 As SAPbouiCOM.EditText
    Private txtFondo As SAPbouiCOM.EditText
    Private txtVarPrueba As SAPbouiCOM.EditText
    Private txtCantNegro As SAPbouiCOM.EditText
    Private txtDefcNegro As SAPbouiCOM.EditText
    Private txtCantFermentado As SAPbouiCOM.EditText
    Private txtDefcFermentado As SAPbouiCOM.EditText
    Private txtCantCerezaSeca As SAPbouiCOM.EditText
    Private txtDefcCerezaSeca As SAPbouiCOM.EditText
    Private txtCantDanioHongo As SAPbouiCOM.EditText
    Private txtDefcDanioHongo As SAPbouiCOM.EditText
    Private txtCantMatExtrania As SAPbouiCOM.EditText
    Private txtDefcMatExtrania As SAPbouiCOM.EditText
    Private txtCantNegroParc As SAPbouiCOM.EditText
    Private txtDefcNegroParc As SAPbouiCOM.EditText
    Private txtCantParcFermn As SAPbouiCOM.EditText
    Private txtDefcParcFermn As SAPbouiCOM.EditText
    Private txtCantPergamino As SAPbouiCOM.EditText
    Private txtDefcPergamino As SAPbouiCOM.EditText
    Private txtCantFlotador As SAPbouiCOM.EditText
    Private txtDefcFlotador As SAPbouiCOM.EditText
    Private txtCantInmaduro As SAPbouiCOM.EditText
    Private txtDefcInmaduro As SAPbouiCOM.EditText
    Private txtCantAveranado As SAPbouiCOM.EditText
    Private txtDefcAveranado As SAPbouiCOM.EditText
    Private txtCantConcha As SAPbouiCOM.EditText
    Private txtDefcConcha As SAPbouiCOM.EditText
    Private txtCantMordido As SAPbouiCOM.EditText
    Private txtDefcMordido As SAPbouiCOM.EditText
    Private txtCantCascaraSeca As SAPbouiCOM.EditText
    Private txtDefcCascaraSeca As SAPbouiCOM.EditText
    Private txtCantBrocado As SAPbouiCOM.EditText
    Private txtDefcBrocado As SAPbouiCOM.EditText
    Private cmbFragancia As SAPbouiCOM.ComboBox
    Private txtComFragancia As SAPbouiCOM.EditText
    Private cmbAcidez As SAPbouiCOM.ComboBox
    Private txtComAcidez As SAPbouiCOM.EditText
    Private cmbCuerpo As SAPbouiCOM.ComboBox
    Private txtComCuerpo As SAPbouiCOM.EditText
    Private cmbSabor As SAPbouiCOM.ComboBox
    Private txtComSabor As SAPbouiCOM.EditText
    Private txtComentarioFinal As SAPbouiCOM.EditText
    Private chkBuenoEmbarque As SAPbouiCOM.CheckBox
    Private btnProceso As SAPbouiCOM.Button
    Private btnImprimirCatacion As SAPbouiCOM.Button
    Private btnAceptar As SAPbouiCOM.Button
    Private instanciado As Boolean = False
    Public Shared lote As New List(Of String)
    Public Shared loteproduccionlist As New List(Of String)
    Public Shared linenumlist As New List(Of String)
    Public Shared tipocafeproduccionlist As New List(Of String)
    Public Shared escalaproduccionlist As New List(Of String)
    Public Shared boletacatacionproduccionlist As New List(Of String)
    Public Shared escalarechazoproduccionlist As New List(Of String)
    Public Shared rendimientonetoproduccionlist As New List(Of String)
    Public Shared rendimientobrutoproduccionlist As New List(Of String)
    Public Shared humedad As New List(Of String)
    Public Shared tipo As New List(Of String)
    Public Shared escala As New List(Of String)
    Public Shared codigoFincaList As New List(Of String)
    Public Shared codigoProductorList As New List(Of String)
    Public Shared nombreFincaList As New List(Of String)
    Public Shared nombreProductorList As New List(Of String)
    Public Shared numeroVueltasList As New List(Of String)
    Public Shared recuperableList As New List(Of String)


    Public Sub New()
        MyBase.New()
        Try

            SetApplication()
            Dim result As Integer
            Dim serrmsg As String = ""

            If Not SetConnectionContext() = 0 Then
                SBO_Application.MessageBox("Failed setting a connection to DI API")
                End ' Terminating the Add-On Application
            End If

            result = ConnectToCompany()
            If Not result = 0 Then
                SBO_Application.MessageBox(result & " Failed connecting to the company's Data Base")
                End ' Terminating the Add-On Application
            End If


            SBO_Application.StatusBar.SetText("Iniciando Addon Catacion...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Utilss.SBOApplication = SBO_Application
            Utilss.Company = oCompany

            SetFilters()
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show(ex.Message & vbNewLine & "SBO application not found")
            System.Windows.Forms.Application.Exit()
        End Try
    End Sub

    Private Sub SetApplication()


        '*******************************************************************
        '// Use an SboGuiApi object to establish connection
        '// with the SAP Business One application and return an
        '// initialized appliction object
        '*******************************************************************

        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi

            '// by following the steps specified above, the following
            '// statment should be suficient for either development or run mode

            sConnectionString = Utilss.ConnectionString  'Environment.GetCommandLineArgs.GetValue(1)

            '// connect to a running SBO Application

            SboGuiApi.Connect(sConnectionString)

            '// get an initialized application object


            SBO_Application = SboGuiApi.GetApplication()
            GC.KeepAlive(SBO_Application)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Ocurrio un error")
        End Try

    End Sub

    Private Function SetConnectionContext() As Integer

        Dim sCookie As String
        Dim sConnectionContext As String

        '// First initialize the Company object

        oCompany = New SAPbobsCOM.Company

        '// Acquire the connection context cookie from the DI API.
        sCookie = oCompany.GetContextCookie

        '// Retrieve the connection context string from the UI API using the
        '// acquired cookie.
        sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)

        '// before setting the SBO Login Context make sure the company is not
        '// connected

        If oCompany.Connected = True Then
            oCompany.Disconnect()
        End If

        '// Set the connection context information to the DI API.
        SetConnectionContext = oCompany.SetSboLoginContext(sConnectionContext)

    End Function

    Private Function ConnectToCompany() As Integer

        Try
            '// Make sure you're not already connected.
            If oCompany.Connected = True Then
                oCompany.Disconnect()
            End If

            'oCompany = SBO_Application.Company.GetDICompany

            '// Establish the connection to the company database.
            ConnectToCompany = oCompany.Connect
        Catch ex As Exception
            SBO_Application.MessageBox(ex.Message)
        End Try

    End Function

    Private Sub SetFilters()

        '// Create a new EventFilters object

        oFilters = New SAPbouiCOM.EventFilters



        '// add an event type to the container

        '// this method returns an EventFilter object

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK)

        'oFilter = oFilter.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE)

        oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

        'oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK)

        ' oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK)
        oFilter.AddEx("60006") 'Quotation Form


        'oFilter.AddEx("60004")
        'oFilters.GetAsXML()
        oFilter.AddEx("UDO_FT_CATACION")
        'oFilter.AddEx("139") 'Orders Form
        'oFilter.AddEx("133") 'Invoice Form
        'oFilter.AddEx("169") 'Main Menu
        SBO_Application.SetFilter(oFilters)

    End Sub

    Private Sub moSBOApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles SBO_Application.FormDataEvent
        Try
            If (BubbleEvent = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) And BusinessObjectInfo.FormTypeEx = "60006" And BusinessObjectInfo.BeforeAction = True Then

#Region "Determinar tipo de cambio"
                Dim RecSet As SAPbobsCOM.Recordset
                Dim sql As String = ""
                sql = "CALL GET_PROCEDURE(6,'','','','','')"
                RecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RecSet.DoQuery(sql)
                If RecSet.RecordCount = 0 Then
                    SBOApplication.MessageBox("No se ha ingresado el tipo de cambio")
                    BubbleEvent = False
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet)
                RecSet = Nothing
                GC.Collect()
#End Region

#Region "Determinar el codigo del tipo de cafe"
                Dim CodigoCafe As SAPbouiCOM.EditText
                oForm = SBO_Application.Forms.ActiveForm

                CodigoCafe = oForm.Items.Item("Item_3").Specific

                If CodigoCafe.Value = "" And cmbEstado.Value = "Ingreso" Then
                    SBOApplication.MessageBox("Error: No Ingreso CAFE")
                    BubbleEvent = False
                    Exit Sub
                Else
                End If
#End Region

#Region "Determinar si la orden de produccion ya esta asociada a una catacion"
                Dim recibo As Integer
                Dim TipoOT As String = String.Empty
                Dim txtOrdenProduccion As SAPbouiCOM.EditText

                oForm = SBO_Application.Forms.ActiveForm
                txtOrdenProduccion = oForm.Items.Item("Item_24").Specific

                Dim rsTipoOT As SAPbobsCOM.Recordset = Nothing
                rsTipoOT = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsTipoOT.DoQuery($"SELECT T0.""U_Tipo_OT"" FROM ""@CATACION"" T0 WHERE T0.""U_DocEntry""='{txtOrdenProduccion.Value.ToString()}'")

                If rsTipoOT.RecordCount > 0 Then
                    TipoOT = rsTipoOT.Fields.Item("U_Tipo_OT").Value.ToString()
                End If

                Dim RecSet2 As SAPbobsCOM.Recordset
                Dim sql2 As String = ""

                sql2 = "CALL GET_PROCEDURE(24,'" + txtOrdenProduccion.Value + "','','','','')"
                RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RecSet2.DoQuery(sql2)
                If RecSet2.RecordCount > 0 Then
                    recibo = RecSet2.Fields.Item("respuesta").Value

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                    RecSet2 = Nothing
                    GC.Collect()
                    If recibo > 0 And TipoOT <> "MEOEXP" Then
                        SBO_Application.MessageBox("Error al Guardar: El numero de recibo ya fue utilizado")
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
#End Region

            End If
        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="FormUID"></param>
    ''' <param name="pVal"></param>
    ''' <param name="BubbleEvent"></param>
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        'If (pVal.FormTypeEx = "60006") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) And (pVal.ItemUID = "Item_22") Then
        '    Try
        '        InstanciarObjetos(pVal)
        '    Catch ex As Exception
        '        Console.WriteLine(ex.ToString)
        '    End Try
        '    ActivarFormulario(False)
        ''End If
        'If (pVal.FormTypeEx = "60006") And (pVal.ItemUID <> "Item_25") And (pVal.ItemUID <> "138") And (pVal.ItemUID <> "Item_24") Then
        '    Exit Sub
        'End If

        Try


            If (pVal.FormTypeEx = "60006") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) And (pVal.ItemUID = "Item_25") Then

                'Dim estado = cmbEstado.Value
                'cmbEstado.Select()
                'OForm = SBO_Application.Forms.ActiveForm
                Try
                    InstanciarObjetos(pVal)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try

                Select Case (oCompany.UserName)
                    Case "manager"
                        ActivarFormulario(True)
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtTotal.Item.Enabled = True
                        'btnProceso.Item.Enabled = True
                    Case "CATACION1"
                        ActivarFormulario(False)
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtTotal.Item.Enabled = True
                        Select Case (cmbEstado.Value)
                            Case "Ingreso"
                                ActivarFormulario(True)
                                'txtEntrada.Item.Enabled = False
                                btnProceso.Item.Enabled = False
                                txtCodCafe.Item.Enabled = False
                                txtCantidad.Item.Enabled = False
                                cmbCatado.Item.Enabled = True
                            Case "Proceso"
                                cmbVerde.Item.Enabled = True
                                cmbTueste.Item.Enabled = True
                                txtComVerde.Item.Enabled = True
                                txtComTueste.Item.Enabled = True
                                txtHumedad.Item.Enabled = True
                                txtGrsLtrs.Item.Enabled = True
                                ActivarComentarios(True)
                                tblOrdenTrabajo.Item.Enabled = True
                                If (txtTipo.Value = "TRILLA") Then
                                    ActivarTaza(False)
                                Else
                                    ActivarTaza(True)
                                End If

                                ActivarZarandra(True)
                                ActivarDefectos(True)
                                btnProceso.Item.Enabled = False
                                txtComTaza.Item.Enabled = False
                                cmbCatado.Item.Enabled = True
                            Case "Exportacion"
                                cmbCatado.Item.Enabled = True
                                tblOrdenTrabajo.Item.Enabled = True
                                cmbVerde.Item.Enabled = True
                                cmbTueste.Item.Enabled = True
                                txtComVerde.Item.Enabled = True
                                txtComTueste.Item.Enabled = True
                                txtHumedad.Item.Enabled = True
                                txtGrsLtrs.Item.Enabled = True
                                ActivarComentarios(True)
                                ActivarTaza(False)
                                ActivarZarandra(True)
                                ActivarDefectos(True)
                                btnProceso.Item.Enabled = False
                                txtScore.Item.Enabled = False
                            Case Else
                                ActivarFormulario(False)
                        End Select
                    Case "CATACION2"
                        ActivarFormulario(False)
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtTotal.Item.Enabled = True
                        Select Case (cmbEstado.Value)
                            Case "Proceso"
                                txtEscala.Item.Enabled = True
                                txtEscalaRechazo.Item.Enabled = True
                                ActivarTaza(True)
                                ActivarComentarios(True)
                                ActivarCalificacion(True)
                                btnProceso.Item.Enabled = False
                                chkTzAgria.Item.Enabled = True
                                chkTzFenolica.Item.Enabled = True
                                chkTzFermentada.Item.Enabled = True
                                chkTzFruti.Item.Enabled = True
                                chkTzMohosa.Item.Enabled = True
                                chkTzSana.Item.Enabled = True
                                chkTzTerrosa.Item.Enabled = True
                                chkTzVinosa.Item.Enabled = True
                                txtTipo.Item.Enabled = True
                                txtRO.Item.Enabled = False
                                txtESC.Item.Enabled = False
                                txtBZ.Item.Enabled = False
                                txtSZ.Item.Enabled = False
                                'txtEntrada.Item.Enabled = False
                                cmbCatado.Item.Enabled = True
                            Case "Exportacion"
                                cmbCatado.Item.Enabled = True
                                tblOrdenTrabajo.Item.Enabled = True
                                ActivarFormularioFinal(True)
                                ActivarTaza(True)
                                ActivarComentarios(True)
                                btnProceso.Item.Enabled = False
                                txtTipo.Item.Enabled = True
                                txtEscala.Item.Enabled = True
                                txtEscalaRechazo.Item.Enabled = True
                                txtScore.Item.Enabled = True
                                'If (txtTipoOT.Value = "MEZCLA") Then
                                '    cmbCatado.Item.Enabled = True
                                'End If
                            Case "Oferta"
                                ActivarFormulario(True)
                                ActivarZarandra(False)
                                ActivarDefectos(False)
                                ActivarFormularioFinal(False)
                                tblOrdenTrabajo.Item.Enabled = False
                                btnProceso.Item.Enabled = False
                                cmbCatado.Item.Enabled = True
                            Case Else
                                ActivarFormulario(False)
                        End Select
                    Case "CATACION3"
                        ActivarFormulario(False)
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtTotal.Item.Enabled = True
                        Select Case (cmbEstado.Value)
                            Case "Proceso"
                                txtEscala.Item.Enabled = True
                                txtEscalaRechazo.Item.Enabled = True
                                ActivarTaza(True)
                                ActivarComentarios(True)
                                ActivarCalificacion(True)
                                btnProceso.Item.Enabled = False
                                cmbCatado.Item.Enabled = True
                            Case "Exportacion"
                                cmbCatado.Item.Enabled = True
                                tblOrdenTrabajo.Item.Enabled = True
                                ActivarFormularioFinal(True)
                                ActivarTaza(True)
                                ActivarComentarios(True)
                                btnProceso.Item.Enabled = False
                            Case "Oferta"
                                ActivarFormulario(True)
                                ActivarZarandra(False)
                                ActivarDefectos(False)
                                ActivarFormularioFinal(False)
                                tblOrdenTrabajo.Item.Enabled = False
                                btnProceso.Item.Enabled = False
                                cmbCatado.Item.Enabled = True
                            Case Else
                                ActivarFormulario(False)
                        End Select
                    Case "PRODUC1"

                        ActivarFormulario(False)
                        txtTotal.Item.Enabled = True
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtCodCafe.Item.Enabled = False
                        txtCantidad.Item.Enabled = True
                        txtObservaciones1.Item.Enabled = True
                        txtObservaciones2.Item.Enabled = True
                        txtObservaciones3.Item.Enabled = True
                        txtRendNetoE.Item.Enabled = False
                        tblOrdenTrabajo.Item.Enabled = True
                        cmbCatado.Item.Enabled = False
                        If (cmbCatado.Value = "SI") Then
                            btnProceso.Item.Enabled = True
                        End If
                    Case "PRODUC2"
                        ActivarFormulario(False)
                        txtTotal.Item.Enabled = True
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtCodCafe.Item.Enabled = False
                        txtCantidad.Item.Enabled = True
                        txtObservaciones1.Item.Enabled = True
                        txtObservaciones2.Item.Enabled = True
                        txtObservaciones3.Item.Enabled = True
                        txtRendNetoE.Item.Enabled = False
                        tblOrdenTrabajo.Item.Enabled = True
                        cmbCatado.Item.Enabled = False
                        If (cmbCatado.Value = "SI") Then
                            btnProceso.Item.Enabled = True
                        End If
                    Case "PRODUC3"
                        ActivarFormulario(False)
                        txtTotal.Item.Enabled = True
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtCodCafe.Item.Enabled = False
                        txtCantidad.Item.Enabled = True
                        txtObservaciones1.Item.Enabled = True
                        txtObservaciones2.Item.Enabled = True
                        txtObservaciones3.Item.Enabled = True
                        txtRendNetoE.Item.Enabled = False
                        tblOrdenTrabajo.Item.Enabled = True
                        cmbCatado.Item.Enabled = False
                        If (cmbCatado.Value = "SI") Then
                            btnProceso.Item.Enabled = True
                        End If
                    Case "PRODUC4"
                        ActivarFormulario(False)
                        txtRendBruto.Item.Enabled = True
                        txtRendNeto.Item.Enabled = True
                        txtTotal.Item.Enabled = True
                        txtCodCafe.Item.Enabled = False
                        txtCantidad.Item.Enabled = True
                        txtObservaciones1.Item.Enabled = True
                        txtObservaciones2.Item.Enabled = True
                        txtObservaciones3.Item.Enabled = True
                        txtRendNetoE.Item.Enabled = False
                        tblOrdenTrabajo.Item.Enabled = True
                        cmbCatado.Item.Enabled = False
                        If (cmbCatado.Value = "SI") Then
                            btnProceso.Item.Enabled = True
                        End If
                    Case Else
                        ActivarFormulario(False)
                End Select
                'BubbleEvent = False
                'oForm.Freeze(False)
                ' ActivarFormulario(True)

            End If

            If (pVal.FormTypeEx = "60006") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) And (pVal.ItemUID = "1") And (pVal.Before_Action = True) Then
                oForm = SBO_Application.Forms.ActiveForm
                Dim btnAceptar As SAPbouiCOM.Button = oForm.Items.Item("1").Specific
                If btnAceptar.Caption.ToString <> "Buscar" Then
                    Dim Cantidad As SAPbouiCOM.EditText = oForm.Items.Item("23_U_E").Specific
                    If Cantidad.Value = "0.0" Or String.IsNullOrEmpty(Cantidad.Value) Then
                        SBO_Application.StatusBar.SetText("Error: La cantidad no debe ser cero", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        Exit Sub
                    End If
                End If
            End If


            'If (pVal.FormTypeEx = "60006") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) And (pVal.ItemUID = "138") And (pVal.Before_Action = True) Then
            If (pVal.FormTypeEx = "60006") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) And (pVal.ItemUID = "138") And (pVal.Before_Action = True) Then
                Try
                    'Validacion boton crear

                    Dim BotonCrear As SAPbouiCOM.Button
                    oForm = SBO_Application.Forms.ActiveForm

                    BotonCrear = oForm.Items.Item("1").Specific

                    If BotonCrear.Caption.ToString = "Actualizar" Then
                        SBO_Application.StatusBar.SetText("Guarde sus cambios antes de procesar", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False
                        Exit Sub
                    End If

                    Dim RecSet As SAPbobsCOM.Recordset
                    Dim sql As String = ""
                    sql = "CALL GET_PROCEDURE(6,'','','','','')"
                    RecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    RecSet.DoQuery(sql)
                    If RecSet.RecordCount = 0 Then
                        SBOApplication.MessageBox("No se ha ingresado el tipo de cambio")
                        BubbleEvent = False
                    Else

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet)
                        RecSet = Nothing
                        GC.Collect()

                        Dim txtCodigoCafe As SAPbouiCOM.EditText
                        Dim txtCantidadCafe As SAPbouiCOM.EditText
                        Dim cmbEstadoCatacion As SAPbouiCOM.ComboBox
                        Dim cmbCatado As SAPbouiCOM.ComboBox
                        Dim cmbTipoOt As SAPbouiCOM.EditText

                        oForm = SBO_Application.Forms.ActiveForm

                        cmbCatado = oForm.Items.Item("Item_110").Specific
                        cmbEstadoCatacion = oForm.Items.Item("Item_25").Specific
                        txtCodigoCafe = oForm.Items.Item("Item_3").Specific
                        txtCantidadCafe = oForm.Items.Item("23_U_E").Specific
                        cmbTipoOt = oForm.Items.Item("Item_90").Specific

                        If cmbEstado.Value <> "Ingreso" And String.IsNullOrEmpty(cmbTipoOt.Value) Then
                            SBO_Application.MessageBox("Error: No se ha Ingresado el Tipo de Orden de Trabajo")
                            BubbleEvent = False
                            Exit Sub
                        End If

                        If String.IsNullOrEmpty(cmbEstadoCatacion.Value) Then
                            SBO_Application.MessageBox("Error: No se ha Ingresado el Estado")
                            BubbleEvent = False
                            Exit Sub
                        End If

                        If txtCodigoCafe.Value = "" And cmbEstadoCatacion.Value = "Ingreso" Then
                            SBO_Application.MessageBox("Error: No Ingreso CAFE")
                            BubbleEvent = False
                            Exit Sub
                        Else
                            If txtCantidadCafe.Value = "" Then
                                SBO_Application.MessageBox("Error: No Ingreso QUINTALES")
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim txtCardCode As SAPbouiCOM.EditText
                                'Dim txtFinca As SAPbouiCOM.EditText
                                'Dim txtQuintales As SAPbouiCOM.EditText
                                'Dim txtTipo As SAPbouiCOM.EditText
                                'Dim txtHumedad As SAPbouiCOM.EditText
                                'Dim txtBruto As SAPbouiCOM.EditText
                                'Dim txtNeto As SAPbouiCOM.EditText
                                'Dim txtOrden As SAPbouiCOM.EditText
                                Dim txtRecibo As SAPbouiCOM.EditText
                                'Dim cmbEstadoCatacion As SAPbouiCOM.ComboBox
                                Dim txtNumeroCatacion As SAPbouiCOM.EditText
                                oForm = SBO_Application.Forms.ActiveForm
                                Dim txtEntrada As SAPbouiCOM.EditText

                                txtEntrada = oForm.Items.Item("Item_78").Specific
                                txtNumeroCatacion = oForm.Items.Item("0_U_E").Specific
                                txtRecibo = oForm.Items.Item("Item_24").Specific 'No.Recibo
                                Dim DocEntryCatacion As Integer
                                DocEntryCatacion = txtNumeroCatacion.Value
                                cmbEstadoCatacion = oForm.Items.Item("Item_25").Specific 'tipo de estado
                                txtCardCode = oForm.Items.Item("Item_47").Specific
                                If cmbEstadoCatacion.Value.ToString = "Proceso" Then

                                    If cmbCatado.Value = "NO" Then
                                        SBO_Application.MessageBox("Error: La orden no se encuentra catada...")
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        Dim respuesta As Integer
                                        Dim RecSetexpo As SAPbobsCOM.Recordset
                                        Dim sqlres As String = ""
                                        sqlres = "CALL GET_PROCEDURE(19,'" + DocEntryCatacion.ToString + "','','','','')"
                                        RecSetexpo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        RecSetexpo.DoQuery(sqlres)
                                        If RecSetexpo.RecordCount > 0 Then
                                            respuesta = RecSetexpo.Fields.Item("respuesta").Value
                                        Else
                                            'MessageBox.Show("No existe la orden de produccion: " + recibo.Value.ToString)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSetexpo)
                                            RecSetexpo = Nothing
                                            GC.Collect()
                                            'Exit Sub
                                        End If
                                        'respuesta = 0
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSetexpo)
                                        RecSetexpo = Nothing
                                        GC.Collect()
                                        If respuesta = 1 Then
                                            SBO_Application.MessageBox("Ya existe un proceso generado con este formulario...")
                                        ElseIf respuesta = 0 Then
                                            Dim DocEntryProduccion As Integer

                                            oForm = SBO_Application.Forms.ActiveForm
                                            txtRecibo = oForm.Items.Item("Item_24").Specific

                                            Dim RecSet2 As SAPbobsCOM.Recordset
                                            Dim sql2 As String = ""
                                            sql2 = "CALL GET_PROCEDURE(9,'" + txtRecibo.Value + "','','','','')"
                                            RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            RecSet2.DoQuery(sql2)
                                            If RecSet2.RecordCount > 0 Then
                                                DocEntryProduccion = RecSet2.Fields.Item("DocEntry").Value
                                                CreaProduccion(DocEntryProduccion, DocEntryCatacion)
                                            Else
                                                SBO_Application.MessageBox("No existe la orden de produccion: " + txtRecibo.Value.ToString)
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                                                RecSet2 = Nothing
                                                GC.Collect()
                                                Exit Sub
                                            End If
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                                            RecSet2 = Nothing
                                            GC.Collect()
                                        End If
                                    End If

                                ElseIf cmbEstadoCatacion.Value.ToString = "Exportacion" Then
#Region "Exportacion"

                                    If cmbCatado.Value = "NO" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        Dim respuesta As Integer
                                        Dim RecSetexpo As SAPbobsCOM.Recordset
                                        Dim sqlres As String = ""
                                        sqlres = "CALL GET_PROCEDURE(18,'" + DocEntryCatacion.ToString + "','','','','')"
                                        RecSetexpo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        RecSetexpo.DoQuery(sqlres)
                                        If RecSetexpo.RecordCount > 0 Then
                                            respuesta = RecSetexpo.Fields.Item("respuesta").Value
                                        Else
                                            'MessageBox.Show("No existe la orden de produccion: " + recibo.Value.ToString)
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSetexpo)
                                            RecSetexpo = Nothing
                                            GC.Collect()
                                            'Exit Sub
                                        End If
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSetexpo)
                                        RecSetexpo = Nothing
                                        GC.Collect()
                                        If respuesta = 1 Then
                                            SBO_Application.MessageBox("Ya existe una exportacion generada con este formulario...")
                                        ElseIf respuesta = 0 Then
                                            'Dim Recibo As SAPbouiCOM.EditText
                                            Dim DocEntryProduccion As Integer
                                            oForm = SBO_Application.Forms.ActiveForm
                                            txtRecibo = oForm.Items.Item("Item_24").Specific
                                            Dim docentryform2 As SAPbouiCOM.EditText
                                            oForm = SBO_Application.Forms.ActiveForm
                                            docentryform2 = oForm.Items.Item("0_U_E").Specific
                                            txtRecibo = oForm.Items.Item("Item_24").Specific 'No.Recibo
                                            Dim RecSet2 As SAPbobsCOM.Recordset
                                            Dim sql2 As String = ""
                                            sql2 = "CALL GET_PROCEDURE(9,'" + txtRecibo.Value + "','','','','')"
                                            RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            RecSet2.DoQuery(sql2)
                                            If RecSet2.RecordCount > 0 Then
                                                DocEntryProduccion = RecSet2.Fields.Item("DocEntry").Value
                                                'CreaExportacion(DocEntryProduccion, docentryform2.Value)
                                                CreaProduccion(DocEntryProduccion, docentryform2.Value)
                                            Else
                                                SBO_Application.MessageBox("No existe la orden de produccion: " + txtRecibo.Value.ToString)
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                                                RecSet2 = Nothing
                                                GC.Collect()
                                                Exit Sub
                                            End If
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                                            RecSet2 = Nothing
                                            GC.Collect()
                                        End If
                                    End If
#End Region
                                ElseIf cmbEstadoCatacion.Value.ToString = "Ingreso" Then
#Region "Ingreso"
                                    If cmbCatado.Value = "NO" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        If Len(txtEntrada.Value.ToString) > 1 Then

                                            SBO_Application.MessageBox("Ya existe una entrada generada..")
                                        Else


                                            'Dim Recibo As SAPbouiCOM.EditText

                                            oForm = SBO_Application.Forms.ActiveForm
                                            txtRecibo = oForm.Items.Item("Item_24").Specific

                                            Dim docentryform2 As SAPbouiCOM.EditText
                                            oForm = SBO_Application.Forms.ActiveForm


                                            docentryform2 = oForm.Items.Item("0_U_E").Specific
                                            txtRecibo = oForm.Items.Item("Item_24").Specific 'No.Recibo


                                            Dim RecSet2 As SAPbobsCOM.Recordset
                                            Dim sql2 As String = ""

                                            sql2 = "CALL GET_PROCEDURE(5,'" + txtRecibo.Value + "','','','','')"
                                            RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            RecSet2.DoQuery(sql2)
                                            If RecSet2.RecordCount > 0 Then
                                                CreaEntradaMerca(txtRecibo.Value, DocEntryCatacion, txtCardCode.Value.ToString)
                                                SBO_Application.ActivateMenuItem("1289") ' //Previous data record
                                                SBO_Application.ActivateMenuItem("1288") '//Following data record
                                            End If
                                            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                                            RecSet2 = Nothing
                                            GC.Collect()
                                        End If
                                    End If
#End Region
                                End If

                            End If
                        End If
                    End If

                Catch ex As Exception
                    SBO_Application.MessageBox("Error" & ex.ToString)
                End Try
            End If


            If (pVal.FormTypeEx = "60006") And (pVal.EventType = SAPbouiCOM.BoEventTypes.et_LOST_FOCUS) And (pVal.ItemUID = "Item_24") And (pVal.Before_Action = False) Then
                Try


                    Dim recibo As Integer
                    Dim campo1sap As SAPbouiCOM.EditText
                    Dim TipoOrden As String = String.Empty
                    oForm = Nothing
                    oForm = SBO_Application.Forms.ActiveForm
                    campo1sap = oForm.Items.Item("Item_24").Specific

                    Dim rsTipoOrden As SAPbobsCOM.Recordset
                    rsTipoOrden = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rsTipoOrden.DoQuery($"SELECT T0.""U_Tipo_OT"" FROM ""@CATACION"" T0 WHERE T0.""U_DocEntry""='{campo1sap.Value}'")
                    If rsTipoOrden.RecordCount > 0 Then
                        TipoOrden = rsTipoOrden.Fields.Item("U_Tipo_OT").Value.ToString()
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsTipoOrden)
                    rsTipoOrden = Nothing
                    GC.Collect()

                    Dim RecSet2 As SAPbobsCOM.Recordset
                    Dim sql2 As String = ""

                    sql2 = "CALL GET_PROCEDURE(24,'" + campo1sap.Value + "','','','','')"
                    RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    RecSet2.DoQuery(sql2)
                    If RecSet2.RecordCount > 0 Then
                        recibo = RecSet2.Fields.Item("respuesta").Value

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                        RecSet2 = Nothing
                        GC.Collect()
                        If recibo > 0 And TipoOrden <> "MEOEXP" Then
                            SBO_Application.StatusBar.SetText("El numero de recibo ya fue utilizado", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End If
                    End If

                    Dim cmbEstadoB As SAPbouiCOM.ComboBox
                    cmbEstadoB = oForm.Items.Item("Item_25").Specific

                    Dim BotonCrear As SAPbouiCOM.Button
                    BotonCrear = oForm.Items.Item("1").Specific

                    Dim existe As SAPbouiCOM.EditText
                    existe = oForm.Items.Item("Item_47").Specific
                    If existe.Value <> "" And BotonCrear.Caption.ToString = "OK" Then
                        Exit Sub
                    Else

                        Dim RecSeti As SAPbobsCOM.Recordset
                        RecSeti = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                        Dim sqli As String = ""

                        sqli = "CALL GET_PROCEDURE(25,'" & campo1sap.Value & "','','','','')"
                        RecSeti.DoQuery(sqli)

                        If (RecSeti.RecordCount > 0) Then
                            Dim codFinca As SAPbouiCOM.EditText
                            Dim Finca As SAPbouiCOM.EditText
                            Dim codCafe As SAPbouiCOM.EditText
                            Dim cantidad As SAPbouiCOM.EditText
                            codFinca = oForm.Items.Item("Item_47").Specific
                            Finca = oForm.Items.Item("34_U_E").Specific
                            codCafe = oForm.Items.Item("Item_3").Specific
                            cantidad = oForm.Items.Item("23_U_E").Specific

                            codFinca.Value = RecSeti.Fields.Item("U_CardCode").Value
                            Finca.Value = RecSeti.Fields.Item("U_CardName").Value
                            codCafe.Value = RecSeti.Fields.Item("U_GroupCade").Value
                            cantidad.Value = RecSeti.Fields.Item("U_Pbascula").Value


                        Else
                            'Case "Proceso"

                            Try


                                'Dim tipoProceso As SAPbouiCOM.EditText
                                'oForm = SBO_Application.Forms.ActiveForm
                                'tipoProceso = oForm.Items.Item("Item_90").Specific




                                Dim RecSeti2 As SAPbobsCOM.Recordset
                                RecSeti2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim sqli2 As String = ""

                                sqli2 = "CALL GET_PROCEDURE(26,'" & campo1sap.Value & "','','','','')"

                                RecSeti2.DoQuery(sqli2)
                                Dim codFinca As SAPbouiCOM.EditText
                                Dim Finca As SAPbouiCOM.EditText
                                Dim codCafe As SAPbouiCOM.EditText
                                Dim cantidad As SAPbouiCOM.EditText
                                Dim tipoOT As SAPbouiCOM.EditText
                                'Dim cEstado As SAPbouiCOM.ComboBox


                                codFinca = oForm.Items.Item("Item_47").Specific
                                Finca = oForm.Items.Item("34_U_E").Specific
                                codCafe = oForm.Items.Item("Item_3").Specific
                                cantidad = oForm.Items.Item("23_U_E").Specific
                                tipoOT = oForm.Items.Item("Item_90").Specific
                                codFinca.Value = RecSeti2.Fields.Item("CardCode").Value
                                Finca.Value = RecSeti2.Fields.Item("CardName").Value

                                Select Case (cmbEstado.Value)
                                    Case "Oferta"
                                        codCafe.Value = "PERGAMINO"
                                    Case "Exportacion"
                                        codCafe.Value = "OROEXP"
                                    Case Else
                                        If (tipoOT.Value = "PERGO-PERG") Then
                                            codCafe.Value = "PERGAMINO"
                                        Else
                                            codCafe.Value = "ORO"
                                        End If
                                End Select

                                cantidad.Value = RecSeti2.Fields.Item("Cantidad").Value


                            Catch ex As Exception
                                Dim RecSeti3 As SAPbobsCOM.Recordset
                                RecSeti3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                Dim sqli3 As String = ""

                                sqli3 = "CALL GET_PROCEDURE(27,'" & campo1sap.Value & "','','','','')"

                                RecSeti3.DoQuery(sqli3)
                                Dim codFinca As SAPbouiCOM.EditText
                                Dim Finca As SAPbouiCOM.EditText
                                Dim codCafe As SAPbouiCOM.EditText
                                Dim cantidad As SAPbouiCOM.EditText
                                codFinca = oForm.Items.Item("Item_47").Specific
                                Finca = oForm.Items.Item("34_U_E").Specific
                                codCafe = oForm.Items.Item("Item_3").Specific
                                cantidad = oForm.Items.Item("23_U_E").Specific

                                codFinca.Value = RecSeti3.Fields.Item("CardCode").Value
                                Finca.Value = RecSeti3.Fields.Item("CardName").Value
                                codCafe.Value = RecSeti3.Fields.Item("GroupCade").Value
                                If cantidad.Value > 0 Then
                                Else
                                    cantidad.Value = RecSeti3.Fields.Item("Cantidad").Value
                                End If
                            End Try
                            'Case "Exportacion"
                            '    Dim RecSeti As SAPbobsCOM.Recordset
                            '    RecSeti = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                            '    Dim sqli As String = ""

                            '    sqli = "CALL GET_PROCEDURE(28,'" & campo1sap.Value & "','','','','')"

                            '    RecSeti.DoQuery(sqli)
                            '    Dim codFinca As SAPbouiCOM.EditText
                            '    Dim Finca As SAPbouiCOM.EditText
                            '    Dim codCafe As SAPbouiCOM.EditText
                            '    Dim cantidad As SAPbouiCOM.EditText
                            '    codFinca = oForm.Items.Item("Item_47").Specific
                            '    Finca = oForm.Items.Item("34_U_E").Specific
                            '    codCafe = oForm.Items.Item("Item_3").Specific
                            '    cantidad = oForm.Items.Item("23_U_E").Specific

                            '    codFinca.Value = RecSeti.Fields.Item("CardCode").Value
                            '    Finca.Value = RecSeti.Fields.Item("CardName").Value
                            '    codCafe.Value = RecSeti.Fields.Item("GroupCade").Value
                            '    cantidad.Value = RecSeti.Fields.Item("Cantidad").Value
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If


        Catch ex As Exception

        End Try

    End Sub

    Public Sub CreaProduccion(DocEntryProduccion, DocEntryCata)
        'Dim btnProceso2 As SAPbouiCOM.Button = oForm.Items.Item("138").Specific
        'btnProceso2.Item.Enabled = False
        Dim CantidadLineas As Integer
        Dim OrdenProduccion As SAPbobsCOM.ProductionOrders
        Dim oreturn As Integer = -1
        Dim RecSet4 As SAPbobsCOM.Recordset
        Dim sql4 As String = ""
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oCatacionData As SAPbobsCOM.GeneralData
        Dim oCafeInferiorData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oChild As SAPbobsCOM.GeneralData
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        Dim rsDistCafe As SAPbobsCOM.Recordset
        Dim rsSubProductos As SAPbobsCOM.Recordset
        Dim queryDistCafe As String = ""
        Dim BatchNumber As StringBuilder = Nothing
        Dim Finca As SAPbobsCOM.BusinessPartners
        Dim Productor As SAPbobsCOM.BusinessPartners
        Dim TipoOrden As String = String.Empty
        Dim Recibo As String = String.Empty
        Dim Finc_ As String = String.Empty
        Dim TipoCafe As String = String.Empty
        Dim Escala As String = String.Empty
        Dim EscalaRechazo As String = String.Empty
        Dim Lote As String = String.Empty
        Dim Quantity As Double = 0
        Dim ItemCode As String = String.Empty
        Dim OrderLinesCount As Integer = 0
        Try

#Region "Recuperar la OT Asociada"
            OrdenProduccion = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
            OrdenProduccion.GetByKey(DocEntryProduccion)
            OrderLinesCount = OrdenProduccion.Lines.Count
#End Region

            If (OrdenProduccion.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased) Then

#Region "Recuperar datos del udo catacion"
                oCompanyService = oCompany.GetCompanyService
                oGeneralService = oCompanyService.GetGeneralService("CATACION")
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralParams.SetProperty("DocEntry", DocEntryCata)
                oCatacionData = oGeneralService.GetByParams(oGeneralParams)
#End Region

#Region "Recuperar la distribucion de cafe inferior"
                queryDistCafe = "CALL RENDIMIENTOS3(" + DocEntryCata.ToString + ")"
                rsDistCafe = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsDistCafe.DoQuery(queryDistCafe)
                Dim FilasOT As List(Of DistribucionCafe) = New List(Of DistribucionCafe)
                Dim FilasUDO As List(Of DistribucionCafe) = New List(Of DistribucionCafe)
                Dim FilaOT As DistribucionCafe = Nothing
                Dim FilaUDO As DistribucionCafe = Nothing
                rsDistCafe.MoveFirst()
                Do While (Not rsDistCafe.EoF)
                    If rsDistCafe.Fields.Item("COMPONENTE").Value.ToString = "PTPRIMERA" Then
                        FilaOT = New DistribucionCafe()
                        FilaOT.CodigoArticulo = rsDistCafe.Fields.Item("Codigo").Value.ToString
                        FilaOT.Recibo = rsDistCafe.Fields.Item("Recibo").Value.ToString
                        FilaOT.CantidadQQ = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaOT.CodigoFinca = rsDistCafe.Fields.Item("U_CodFinca").Value.ToString
                        FilaOT.Cantidad = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaOT.DescripcionArticulo = rsDistCafe.Fields.Item("Codigo").Value.ToString
                        FilaOT.CantidadRequerida = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaOT.TipoCafe = rsDistCafe.Fields.Item("U_Tipo").Value.ToString
                        FilaOT.Escala = rsDistCafe.Fields.Item("U_Escala").Value.ToString
                        FilaOT.EscalaRechazo = rsDistCafe.Fields.Item("U_Escala_Rechazo").Value.ToString
                        FilaOT.CalidadCafe = rsDistCafe.Fields.Item("COMPONENTE").Value.ToString
                        FilaOT.Consumido = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaOT.Disponible = Math.Round(rsDistCafe.Fields.Item("Stock_Alm").Value, 2)
                        FilaOT.CodigoUnidadMedida = rsDistCafe.Fields.Item("Cod_Und_Med").Value.ToString
                        FilaOT.DescripcionUnidadMedida = rsDistCafe.Fields.Item("Nom_Und_Med").Value.ToString
                        FilaOT.CodigoAlmacen = rsDistCafe.Fields.Item("Almacen").Value.ToString
                        FilaOT.MetodoEmision = rsDistCafe.Fields.Item("Metodo_Emision").Value.ToString
                        FilaOT.CantidadPendiente = Math.Round(rsDistCafe.Fields.Item("Cantidad_Pendiente").Value, 2)
                        FilaOT.CantidadAlmacen = Math.Round(rsDistCafe.Fields.Item("Cantidad_Pendiente").Value, 2)
                        FilaOT.SectorRuta = rsDistCafe.Fields.Item("Sec-Ruta").Value.ToString
                        FilaOT.RendimientoNeto = Math.Round(rsDistCafe.Fields.Item("U_Rend_Neto").Value, 6)
                        FilaOT.RendimientoBruto = Math.Round(rsDistCafe.Fields.Item("U_Rend_Bruto").Value, 6)
                        FilaOT.CantidadSacos = rsDistCafe.Fields.Item("Cant_Sacos").Value.ToString
                        FilaOT.Vueltas = rsDistCafe.Fields.Item("U_NVuelta").Value.ToString
                        FilaOT.Recuperacion = rsDistCafe.Fields.Item("U_Recup").Value.ToString
                        FilasOT.Add(FilaOT)
                    ElseIf rsDistCafe.Fields.Item("COMPONENTE").Value.ToString = "INFERIORES" Then
                        FilaUDO = New DistribucionCafe()
                        FilaUDO.CodigoArticulo = rsDistCafe.Fields.Item("Codigo").Value.ToString
                        FilaUDO.Recibo = rsDistCafe.Fields.Item("Recibo").Value.ToString
                        FilaUDO.CantidadQQ = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaUDO.CodigoFinca = rsDistCafe.Fields.Item("U_CodFinca").Value.ToString
                        FilaUDO.Cantidad = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaUDO.DescripcionArticulo = rsDistCafe.Fields.Item("Codigo").Value.ToString
                        FilaUDO.CantidadRequerida = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaUDO.TipoCafe = rsDistCafe.Fields.Item("U_Tipo").Value.ToString
                        FilaUDO.Escala = rsDistCafe.Fields.Item("U_Escala").Value.ToString
                        FilaUDO.EscalaRechazo = rsDistCafe.Fields.Item("U_Escala_Rechazo").Value.ToString
                        FilaUDO.CalidadCafe = rsDistCafe.Fields.Item("COMPONENTE").Value.ToString
                        FilaUDO.Consumido = Math.Round(rsDistCafe.Fields.Item("Cantidad").Value, 2)
                        FilaUDO.Disponible = Math.Round(rsDistCafe.Fields.Item("Stock_Alm").Value, 2)
                        FilaUDO.CodigoUnidadMedida = rsDistCafe.Fields.Item("Cod_Und_Med").Value.ToString
                        FilaUDO.DescripcionUnidadMedida = rsDistCafe.Fields.Item("Nom_Und_Med").Value.ToString
                        FilaUDO.CodigoAlmacen = rsDistCafe.Fields.Item("Almacen").Value.ToString
                        FilaUDO.MetodoEmision = rsDistCafe.Fields.Item("Metodo_Emision").Value.ToString
                        FilaUDO.CantidadPendiente = Math.Round(rsDistCafe.Fields.Item("Cantidad_Pendiente").Value, 2)
                        FilaUDO.CantidadAlmacen = Math.Round(rsDistCafe.Fields.Item("Cantidad_Pendiente").Value, 2)
                        FilaUDO.SectorRuta = rsDistCafe.Fields.Item("Sec-Ruta").Value.ToString
                        FilaUDO.RendimientoNeto = Math.Round(rsDistCafe.Fields.Item("U_Rend_Neto").Value, 6)
                        FilaUDO.RendimientoBruto = Math.Round(rsDistCafe.Fields.Item("U_Rend_Bruto").Value, 6)
                        FilaUDO.CantidadSacos = rsDistCafe.Fields.Item("Cant_Sacos").Value.ToString
                        FilaUDO.Vueltas = rsDistCafe.Fields.Item("U_NVuelta").Value.ToString
                        FilaUDO.Recuperacion = rsDistCafe.Fields.Item("U_Recup").Value.ToString
                        FilasUDO.Add(FilaUDO)
                    End If
                    rsDistCafe.MoveNext()
                Loop
#End Region

#Region "Recuperar SubProductos"
                If Not oCatacionData.GetProperty("U_Tipo_OT").ToString().Contains("MEZCLA") Then
                    If Not oCatacionData.GetProperty("U_Tipo_OT").ToString().Contains("TRASIEGO") Then
                        Dim querySubProductos As String = $"SELECT T0.* FROM ""@OT_SUB_PRODUCTOS"" T0 WHERE T0.""DocEntry""={DocEntryCata} AND T0.""U_Tipo"" IS NOT NULL AND T0.""U_Cantidad"">0"
                        rsSubProductos = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        rsSubProductos.DoQuery(querySubProductos)
                        If (rsSubProductos.RecordCount > 0) Then
                            rsSubProductos.MoveFirst()
                            Do While (Not rsSubProductos.EoF)
                                If Not String.IsNullOrEmpty(rsSubProductos.Fields.Item("U_Codigo").Value.ToString()) And Not String.IsNullOrEmpty(rsSubProductos.Fields.Item("U_Cantidad").Value.ToString) Then
                                    FilaOT = New DistribucionCafe()
                                    FilaOT.CodigoArticulo = rsSubProductos.Fields.Item("U_Codigo").Value.ToString
                                    FilaOT.Recibo = oCatacionData.GetProperty("U_DocEntry").ToString.Trim()
                                    FilaOT.CantidadQQ = -Math.Round(rsSubProductos.Fields.Item("U_Cantidad").Value, 2)
                                    FilaOT.CodigoFinca = "F00440"
                                    FilaOT.CantidadRequerida = -Math.Round(rsSubProductos.Fields.Item("U_Cantidad").Value, 2)
                                    FilaOT.DescripcionArticulo = rsSubProductos.Fields.Item("U_Codigo").Value.ToString
                                    FilaOT.Cantidad = -Math.Round(rsSubProductos.Fields.Item("U_Cantidad").Value, 2)
                                    FilaOT.TipoCafe = rsSubProductos.Fields.Item("U_Tipo").Value.ToString
                                    FilaOT.Escala = IIf(Not String.IsNullOrEmpty(rsSubProductos.Fields.Item("U_Escala").Value.ToString), rsSubProductos.Fields.Item("U_Escala").Value.ToString, oCatacionData.GetProperty("U_Escala").ToString)
                                    FilaOT.EscalaRechazo = IIf(Not String.IsNullOrEmpty(rsSubProductos.Fields.Item("U_Escala_Rechazo").Value.ToString), rsSubProductos.Fields.Item("U_Escala_Rechazo").Value.ToString, oCatacionData.GetProperty("U_Escala_Rechazo").ToString)
                                    FilaOT.CalidadCafe = GetCalidadCafe(FilaOT.TipoCafe)
                                    FilaOT.Consumido = 0
                                    FilaOT.Disponible = 0
                                    FilaOT.CodigoUnidadMedida = String.Empty
                                    FilaOT.DescripcionUnidadMedida = String.Empty
                                    FilaOT.CodigoAlmacen = "07"
                                    FilaOT.MetodoEmision = String.Empty
                                    FilaOT.CantidadPendiente = 0
                                    FilaOT.CantidadAlmacen = 0
                                    FilaOT.SectorRuta = String.Empty
                                    FilaOT.RendimientoNeto = Double.Parse(oCatacionData.GetProperty("U_Rend_Neto").ToString)
                                    FilaOT.RendimientoBruto = Double.Parse(oCatacionData.GetProperty("U_Rend_Bruto").ToString)
                                    FilaOT.CantidadSacos = String.Empty
                                    FilaOT.Vueltas = rsSubProductos.Fields.Item("U_NVuelta").Value.ToString
                                    FilaOT.Recuperacion = rsSubProductos.Fields.Item("U_prueba1").Value.ToString
                                    FilasOT.Add(FilaOT)
                                End If
                                rsSubProductos.MoveNext()
                            Loop
                        End If
                    End If
                End If

#End Region

#Region "Primera linea que se agrega a la OT"

                    sql4 = "CALL GET_PROCEDURE(11,'" + DocEntryProduccion.ToString + "','','','','')"
                    RecSet4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    RecSet4.DoQuery(sql4)
                    CantidadLineas = RecSet4.RecordCount
                    If RecSet4.RecordCount > 0 Then

                        OrdenProduccion.Lines.UserFields.Fields.Item("U_Tipo").Value = RecSet4.Fields.Item("COMPONENTE").Value
                        OrdenProduccion.Lines.UserFields.Fields.Item("U_Recibo").Value = RecSet4.Fields.Item("Recibo").Value
                        OrdenProduccion.Lines.ItemNo = RecSet4.Fields.Item("Codigo").Value
                        OrdenProduccion.Lines.Warehouse = RecSet4.Fields.Item("Almacen").Value
                        OrdenProduccion.Lines.UserFields.Fields.Item("U_Cantidad").Value = RecSet4.Fields.Item("Cantidad").Value
                        OrdenProduccion.Lines.BaseQuantity = Math.Round(RecSet4.Fields.Item("Cantidad").Value, 2)
                        OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value = RecSet4.Fields.Item("U_Escala").Value
                        OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Neto").Value = RecSet4.Fields.Item("U_Rend_Neto").Value
                        OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Bruto").Value = RecSet4.Fields.Item("U_Rend_Bruto").Value
                        OrdenProduccion.Lines.Add()
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet4)
                    RecSet4 = Nothing
                    GC.Collect()
#End Region

#Region "Segundo bloque de lineas que se insertan en la OT"
                    If FilasOT.Count > 0 Then
                        Dim y As Integer = 0
                        For Each otrow As DistribucionCafe In FilasOT

                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Tipo").Value = otrow.CalidadCafe
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Recibo").Value = otrow.Recibo
                            OrdenProduccion.Lines.ItemNo = otrow.CodigoArticulo
                            OrdenProduccion.Lines.Warehouse = IIf(String.IsNullOrEmpty(otrow.CodigoAlmacen), "07", otrow.CodigoAlmacen)
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Cantidad").Value = Math.Round(otrow.Cantidad, 2)
                            OrdenProduccion.Lines.BaseQuantity = Math.Round(otrow.Cantidad, 2)
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value = otrow.Escala
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Neto").Value = otrow.RendimientoNeto
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Bruto").Value = otrow.RendimientoBruto
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_NombreProductor").Value = otrow.CodigoFinca
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_TipoCafe").Value = otrow.TipoCafe
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_EscalaRechazo").Value = otrow.EscalaRechazo
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_NVuelta").Value = otrow.Vueltas
                            OrdenProduccion.Lines.UserFields.Fields.Item("U_Recup").Value = otrow.Recuperacion
                            OrdenProduccion.Lines.Add()
                            y += 1
                        Next
                    End If

#End Region

                    Dim ret As Integer = OrdenProduccion.Update
                    If ret <> 0 Then
                        SBO_Application.MessageBox(oCompany.GetLastErrorDescription)
                    Else
                        Dim ReciboProduccion As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

#Region "Vaciado de listas"

                        loteproduccionlist.Clear()
                        linenumlist.Clear()
                        tipocafeproduccionlist.Clear()
                        escalaproduccionlist.Clear()
                        rendimientobrutoproduccionlist.Clear()
                        rendimientonetoproduccionlist.Clear()
                        escalarechazoproduccionlist.Clear()
                        boletacatacionproduccionlist.Clear()
                        codigoFincaList.Clear()
                        codigoProductorList.Clear()
                        nombreFincaList.Clear()
                        nombreProductorList.Clear()
                        numeroVueltasList.Clear()
                        recuperableList.Clear()
                        humedad.Clear()
#End Region
                        Dim cont As Integer = CantidadLineas
                        For i As Integer = OrderLinesCount To OrdenProduccion.Lines.Count - 1
                            OrdenProduccion.Lines.SetCurrentLine(i)
                            ReciboProduccion.Lines.BaseEntry = OrdenProduccion.AbsoluteEntry
                            ReciboProduccion.Lines.BaseLine = OrdenProduccion.Lines.LineNumber
                            ReciboProduccion.Lines.BaseType = 202
                            If Math.Sign(OrdenProduccion.Lines.BaseQuantity) < 0 Then
                                Quantity = Math.Round(OrdenProduccion.Lines.BaseQuantity * -1, 2)
                            Else
                                Quantity = Math.Round(OrdenProduccion.Lines.BaseQuantity, 2)
                            End If

                            ReciboProduccion.Lines.Quantity = Quantity
                            ReciboProduccion.Lines.WarehouseCode = "07"
                            BatchNumber = New StringBuilder()
                            Select Case oCatacionData.GetProperty("U_Estado").ToString 'oCatacionData.GetProperty("U_Tipo_OT").ToString
                                Case "Exportacion"
                                    Lote = String.Concat(OrdenProduccion.Lines.UserFields.Fields.Item("U_Recibo").Value.ToString, "-01")
                                    ReciboProduccion.Lines.BatchNumbers.BatchNumber = Lote
                                    ReciboProduccion.Lines.BatchNumbers.Quantity = Quantity
                                    ReciboProduccion.Lines.BatchNumbers.Add()
                                Case Else

                                    If (Not String.IsNullOrEmpty(oCatacionData.GetProperty("U_Tipo_OT").ToString)) Then
                                        TipoOrden = oCatacionData.GetProperty("U_Tipo_OT").ToString.Substring(0, 3)
                                        BatchNumber.Append(TipoOrden)
                                    End If

                                    If (Not String.IsNullOrEmpty(oCatacionData.GetProperty("U_DocEntry").ToString)) Then
                                        Recibo = oCatacionData.GetProperty("U_DocEntry").ToString
                                        BatchNumber.Append(String.Concat("-", Recibo))
                                    End If

                                    If (Not String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_NombreProductor").Value.ToString)) Then
                                        Finc_ = OrdenProduccion.Lines.UserFields.Fields.Item("U_NombreProductor").Value.ToString
                                        BatchNumber.Append(String.Concat("-", Finc_))
                                    End If

                                    If (Not String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_TipoCafe").Value.ToString)) Then
                                        TipoCafe = OrdenProduccion.Lines.UserFields.Fields.Item("U_TipoCafe").Value.ToString
                                        If TipoCafe.Length <= 2 Then
                                            BatchNumber.Append(String.Concat("-", TipoCafe))
                                        ElseIf TipoCafe.Length > 2 Then
                                            BatchNumber.Append(String.Concat("-", TipoCafe.Substring(0, 3)))
                                        End If

                                    End If

                                    If (Not String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value.ToString)) Then
                                        Escala = OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value.ToString
                                        BatchNumber.Append(String.Concat("-", Escala))
                                    Else
                                        BatchNumber.Append("-0")
                                    End If

                                    If (Not String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_EscalaRechazo").Value.ToString)) Then
                                        EscalaRechazo = OrdenProduccion.Lines.UserFields.Fields.Item("U_EscalaRechazo").Value.ToString
                                        BatchNumber.Append(String.Concat("-", EscalaRechazo))
                                    Else
                                        BatchNumber.Append("-0")
                                    End If

                                    BatchNumber.Append(String.Concat("-", (i + 1).ToString.PadLeft(2, "0")))
                                    Lote = BatchNumber.ToString

                                    ReciboProduccion.Lines.BatchNumbers.BatchNumber = Lote
                                    ReciboProduccion.Lines.BatchNumbers.Quantity = Quantity
                                    ReciboProduccion.Lines.BatchNumbers.Add()
                            End Select
                            ReciboProduccion.Lines.Add()

#Region "Llenar las colecciones"

                            loteproduccionlist.Add(Lote) 'llena listas para lotes
                            linenumlist.Add(OrdenProduccion.Lines.LineNumber.ToString)
                            tipocafeproduccionlist.Add(OrdenProduccion.Lines.UserFields.Fields.Item("U_TipoCafe").Value.ToString)
                            escalaproduccionlist.Add(IIf(String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value.ToString), "0", OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value.ToString))
                            humedad.Add(oCatacionData.GetProperty("U_Humedad").ToString)
                            rendimientobrutoproduccionlist.Add(OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Bruto").Value.ToString)
                            rendimientonetoproduccionlist.Add(OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Neto").Value.ToString)
                            escalarechazoproduccionlist.Add(IIf(String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_EscalaRechazo").Value.ToString), "0", OrdenProduccion.Lines.UserFields.Fields.Item("U_EscalaRechazo").Value.ToString))
                            boletacatacionproduccionlist.Add(OrdenProduccion.Lines.UserFields.Fields.Item("U_Recibo").Value.ToString)
                            Finca = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                            Finca.GetByKey(OrdenProduccion.Lines.UserFields.Fields.Item("U_NombreProductor").Value.ToString)
                            codigoFincaList.Add(Finca.CardCode)
                            nombreFincaList.Add(Finca.CardName)
                            Productor = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                            Productor.GetByKey(Finca.FatherCard)
                            codigoProductorList.Add(Productor.CardCode)
                            nombreProductorList.Add(Productor.CardName)
                            numeroVueltasList.Add(IIf(String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_NVuelta").Value.ToString), "0", OrdenProduccion.Lines.UserFields.Fields.Item("U_NVuelta").Value.ToString))
                            recuperableList.Add(IIf(String.IsNullOrEmpty(OrdenProduccion.Lines.UserFields.Fields.Item("U_Recup").Value.ToString), "0", OrdenProduccion.Lines.UserFields.Fields.Item("U_Recup").Value.ToString))
#End Region
                            cont = cont + 1

                        Next
                        oreturn = ReciboProduccion.Add()
                        If oreturn <> 0 Then
                            SBO_Application.MessageBox("Error, debe de generar la emision de produccion o verificar el error: " & oCompany.GetLastErrorDescription)
                        Else
                            For cont = 0 To loteproduccionlist.Count - 1
                                Dim docentry As Integer
                                Dim sqlabs = "CALL GET_PROCEDURE(133,'" + loteproduccionlist.Item(cont).ToString + "','','','','')"
                                Dim RecSet2 As SAPbobsCOM.Recordset
                                RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                RecSet2.DoQuery(sqlabs)
                                If RecSet2.RecordCount > 0 Then
                                    docentry = RecSet2.Fields.Item("AbsEntry").Value
                                End If
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                                RecSet2 = Nothing
                                GC.Collect()

                                Dim oBatch As SAPbobsCOM.BatchNumberDetailsService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BatchNumberDetailsService)
                                Dim bNumDetailParams As SAPbobsCOM.BatchNumberDetailParams = oBatch.GetDataInterface(SAPbobsCOM.BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams)
                                bNumDetailParams.DocEntry = docentry
                                Dim oBacthDetail As SAPbobsCOM.BatchNumberDetail = oBatch.[Get](bNumDetailParams)
                                oBacthDetail.UserFields.Item("U_Escala").Value = escalaproduccionlist.Item(cont)
                                oBacthDetail.UserFields.Item("U_Humedad").Value = humedad.Item(cont)
                                oBacthDetail.UserFields.Item("U_TipoCafe").Value = tipocafeproduccionlist.Item(cont)
                                oBacthDetail.UserFields.Item("U_Rend_Neto").Value = rendimientonetoproduccionlist.Item(cont)
                                oBacthDetail.UserFields.Item("U_Rend_Bruto").Value = rendimientobrutoproduccionlist.Item(cont)
                                oBacthDetail.UserFields.Item("U_Escala_Rechazo").Value = escalarechazoproduccionlist.Item(cont)
                                oBacthDetail.UserFields.Item("U_Boleta_Catacion").Value = boletacatacionproduccionlist.Item(cont)
                                oBacthDetail.UserFields.Item("U_CodFinca").Value = codigoFincaList.Item(cont)
                                oBacthDetail.UserFields.Item("U_CodProductor").Value = codigoProductorList.Item(cont)
                                oBacthDetail.UserFields.Item("U_Nfinca").Value = nombreFincaList.Item(cont)
                                oBacthDetail.UserFields.Item("U_Nproductor").Value = nombreProductorList.Item(cont)
                                oBacthDetail.UserFields.Item("U_NVuelta").Value = numeroVueltasList.Item(cont)
                                oBacthDetail.UserFields.Item("U_Recuperable").Value = recuperableList.Item(cont)
                                oBatch.Update(oBacthDetail)
                            Next

                            Dim RecUpdate As SAPbobsCOM.Recordset
                            Dim sqlupdate As String = ""
                            If (oCatacionData.GetProperty("U_Estado").ToString = "Exportacion") Then
                                sqlupdate = "CALL GET_PROCEDURE(20,'" + DocEntryCata.ToString + "','','','','')"
                                RecUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                RecUpdate.DoQuery(sqlupdate)
                            End If


                            sqlupdate = "CALL GET_PROCEDURE(21,'" + DocEntryCata.ToString + "','','','','')"
                            RecUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            RecUpdate.DoQuery(sqlupdate)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecUpdate)
                            RecUpdate = Nothing
                            GC.Collect()

#Region "Insertar la distribucion de cafe inferior en el udo Café Inferior"
                            If FilasUDO.Count > 0 Then

                                oGeneralService = oCompanyService.GetGeneralService("Café Inferior")
                                oCafeInferiorData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)

                                Select Case OrdenProduccion.ProductionOrderType
                                    Case SAPbobsCOM.BoProductionOrderTypeEnum.bopotDisassembly
                                        oCafeInferiorData.SetProperty("U_Tipo", "Desmontar")
                                    Case SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial
                                        oCafeInferiorData.SetProperty("U_Tipo", "Especial")
                                    Case SAPbobsCOM.BoProductionOrderTypeEnum.bopotStandard
                                        oCafeInferiorData.SetProperty("U_Tipo", "Estandar")
                                End Select

                                Select Case OrdenProduccion.ProductionOrderStatus
                                    Case SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
                                        oCafeInferiorData.SetProperty("U_Estado", "Liberado")
                                    Case SAPbobsCOM.BoProductionOrderStatusEnum.boposClosed
                                        oCafeInferiorData.SetProperty("U_Estado", "Cerrado")
                                    Case SAPbobsCOM.BoProductionOrderStatusEnum.boposCancelled
                                        oCafeInferiorData.SetProperty("U_Estado", "Cancelado")
                                    Case SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned
                                        oCafeInferiorData.SetProperty("U_Estado", "Planificado")
                                End Select

                                oCafeInferiorData.SetProperty("U_TOrden", OrdenProduccion.ItemNo)
                                oCafeInferiorData.SetProperty("U_DesProductor", OrdenProduccion.ItemNo)
                                oCafeInferiorData.SetProperty("U_Cantidad", OrdenProduccion.PlannedQuantity.ToString)
                                oCafeInferiorData.SetProperty("U_Almacen", OrdenProduccion.Warehouse)
                                oCafeInferiorData.SetProperty("U_Prioridad", OrdenProduccion.Priority.ToString)

                                Dim oSeriesService As SAPbobsCOM.SeriesService
                                Dim oSeries As SAPbobsCOM.Series
                                Dim oSeriesParams As SAPbobsCOM.SeriesParams
                                oSeriesService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)
                                oSeriesParams = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeriesParams)
                                oSeriesParams.Series = OrdenProduccion.Series
                                oSeries = oSeriesService.GetSeries(oSeriesParams)
                                oCafeInferiorData.SetProperty("U_Cosecha", oSeries.Name)
                                oCafeInferiorData.SetProperty("U_OFabricacion", OrdenProduccion.DocumentNumber.ToString)
                                oCafeInferiorData.SetProperty("U_FFabricacion", OrdenProduccion.PostingDate.ToString)
                                oCafeInferiorData.SetProperty("U_FInicio", OrdenProduccion.StartDate.ToString)
                                oCafeInferiorData.SetProperty("U_FFinalizacion", OrdenProduccion.DueDate.ToString)

                                oCafeInferiorData.SetProperty("U_Usuario", oCompany.UserName)

                                Select Case OrdenProduccion.ProductionOrderOrigin
                                    Case SAPbobsCOM.BoProductionOrderOriginEnum.bopooManual
                                        oCafeInferiorData.SetProperty("U_Origen", "Manual")
                                    Case SAPbobsCOM.BoProductionOrderOriginEnum.bopooMRP
                                        oCafeInferiorData.SetProperty("U_Origen", "MRP")
                                    Case SAPbobsCOM.BoProductionOrderOriginEnum.bopooSalesOrder
                                        oCafeInferiorData.SetProperty("U_Origen", "Orden de venta")
                                End Select

                                If Not String.IsNullOrEmpty(OrdenProduccion.ProductionOrderOriginNumber.ToString) Then
                                    oCafeInferiorData.SetProperty("U_Pedido", OrdenProduccion.ProductionOrderOriginNumber.ToString)
                                End If

                                oCafeInferiorData.SetProperty("U_Cliente", OrdenProduccion.CustomerCode)
                                oCafeInferiorData.SetProperty("U_Reparto", OrdenProduccion.DistributionRule)
                                oCafeInferiorData.SetProperty("U_Proyecto", OrdenProduccion.Project)
                                oChildren = oCafeInferiorData.Child("CINFERIOR")
                                For Each udorow As DistribucionCafe In FilasUDO
                                    oChild = oChildren.Add
                                    oChild.SetProperty("U_ItemCode", udorow.CodigoArticulo)
                                    oChild.SetProperty("U_Recibo", udorow.Recibo)
                                    oChild.SetProperty("U_CantQQ", Math.Round(udorow.Cantidad, 2))
                                    oChild.SetProperty("U_CodFinca", udorow.CodigoFinca) 'aqui va el codigo de finca del parametro y no el que viene de la distribucion de cafe.
                                    oChild.SetProperty("U_CantR", Math.Round(udorow.Cantidad, 2))
                                    oChild.SetProperty("U_Desc", udorow.CodigoArticulo)
                                    oChild.SetProperty("U_Cantidad", Math.Round(udorow.Cantidad, 2))
                                    oChild.SetProperty("U_TipoCafe", udorow.TipoCafe)
                                    oChild.SetProperty("U_Escala", udorow.Escala)
                                    oChild.SetProperty("U_EscalaR", udorow.EscalaRechazo)
                                    oChild.SetProperty("U_Tipo", udorow.CalidadCafe)
                                    oChild.SetProperty("U_Consumido", Math.Round(udorow.Consumido, 2))
                                    oChild.SetProperty("U_Disponible", Math.Round(udorow.Disponible, 2))
                                    oChild.SetProperty("U_CMedida", udorow.CodigoUnidadMedida)
                                    oChild.SetProperty("U_NMedida", udorow.DescripcionUnidadMedida)
                                    oChild.SetProperty("U_Almacen", udorow.CodigoAlmacen)
                                    oChild.SetProperty("U_MEmision", udorow.MetodoEmision)
                                    oChild.SetProperty("U_CPendiente", Math.Round(udorow.CantidadPendiente, 2))
                                    oChild.SetProperty("U_CAlmacen", Math.Round(udorow.CantidadAlmacen, 2))
                                    oChild.SetProperty("U_SRuta", udorow.SectorRuta)
                                    oChild.SetProperty("U_Rend_Neto", Math.Round(udorow.RendimientoNeto, 6))
                                    oChild.SetProperty("U_Rend_Bruto", Math.Round(udorow.RendimientoBruto, 6))
                                    oChild.SetProperty("U_CSacos", udorow.CantidadSacos)
                                    oChild.SetProperty("U_Vueltas", udorow.Vueltas)
                                    oChild.SetProperty("U_Recup", udorow.Recuperacion)
                                Next
                                oGeneralService.Add(oCafeInferiorData)
                            End If
#End Region
                        ' If SBO_Application Is Nothing Then

                        'SetApplication()
                        ' End If
                        System.Windows.Forms.MessageBox.Show("Produccion Generada Exitosamente..")
                        ' SBO_Application.MessageBox("Produccion Generada Exitosamente..")

                        'SBO_Application.StatusBar.SetText("Produccion Generada Exitosamente..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                    End If


                    End If
                Else
                    SBO_Application.StatusBar.SetText("La Orden de produccion debe estar en estatus liberada.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If

        Catch ex As Exception

            SBO_Application.StatusBar.SetText("Verifique orden de trabajo antes de intentar procesar e intente nuevamente - " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
        'btnProceso2.Item.Enabled = True
    End Sub

    Public Sub CreaExportacion(DocEntryProduccion, DocEntryCata)


        Dim contadorlineas As Integer
        Dim OrdenProduccion As SAPbobsCOM.ProductionOrders
        Dim oreturn As Integer = -1
        Dim loteproduccion As String
        OrdenProduccion = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
        OrdenProduccion.GetByKey(DocEntryProduccion)
        Dim objStreamWriter As StreamWriter
        objStreamWriter = New StreamWriter(Application.StartupPath + "\Log_Exportacion" & DocEntryProduccion.ToString & "_" & DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") & ".txt")
        objStreamWriter.WriteLine("DocEntryproduccion: " & DocEntryProduccion.ToString)
        Dim RecSet4 As SAPbobsCOM.Recordset
        Dim sql4 As String = ""
        sql4 = "CALL GET_PROCEDURE(11,'" + DocEntryProduccion.ToString + "','','','','')"
        objStreamWriter.WriteLine("Sql update: " & sql4)
        RecSet4 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RecSet4.DoQuery(sql4)
        contadorlineas = RecSet4.RecordCount
        If RecSet4.RecordCount > 0 Then
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Tipo").Value = RecSet4.Fields.Item("COMPONENTE").Value
            objStreamWriter.WriteLine("Componente: " & RecSet4.Fields.Item("COMPONENTE").Value)
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Recibo").Value = RecSet4.Fields.Item("Recibo").Value

            OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Bruto").Value = RecSet4.Fields.Item("U_Rend_Bruto").Value
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Neto").Value = RecSet4.Fields.Item("U_Rend_Neto").Value
            If RecSet4.Fields.Item("U_Escala").Value = "" Then
                OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value = "0"
            Else
                OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value = RecSet4.Fields.Item("U_Escala").Value
            End If



            objStreamWriter.WriteLine("Recibo: " & RecSet4.Fields.Item("Recibo").Value)
            OrdenProduccion.Lines.ItemNo = RecSet4.Fields.Item("Codigo").Value
            objStreamWriter.WriteLine("ItemNo: " & RecSet4.Fields.Item("Codigo").Value)
            OrdenProduccion.Lines.Warehouse = RecSet4.Fields.Item("Almacen").Value
            objStreamWriter.WriteLine("Almacen: " & RecSet4.Fields.Item("Almacen").Value)
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Cantidad").Value = Math.Round(RecSet4.Fields.Item("Cantidad").Value, 2)
            objStreamWriter.WriteLine("Cantidad: " & RecSet4.Fields.Item("Cantidad").Value)
            OrdenProduccion.Lines.BaseQuantity = Math.Round(RecSet4.Fields.Item("Cantidad").Value, 2)
            objStreamWriter.WriteLine("BaseQuantity: " & RecSet4.Fields.Item("Cantidad").Value)
            OrdenProduccion.Lines.Add()
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet4)
        RecSet4 = Nothing
        GC.Collect()

        Dim RecSet3 As SAPbobsCOM.Recordset
        Dim sql3 As String = ""
        sql3 = "CALL GET_PROCEDURE(17,'" + DocEntryCata.ToString + "','','','','')"
        objStreamWriter.WriteLine("TraeLineas: " & sql3)
        RecSet3 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        RecSet3.DoQuery(sql3)
        Do While Not RecSet3.EoF
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Tipo").Value = RecSet3.Fields.Item("COMPONENTE").Value
            objStreamWriter.WriteLine("Componente: " & RecSet3.Fields.Item("COMPONENTE").Value)
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Recibo").Value = RecSet3.Fields.Item("Recibo").Value
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Bruto").Value = RecSet3.Fields.Item("U_Rend_Bruto").Value
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Rend_Neto").Value = RecSet3.Fields.Item("U_Rend_Neto").Value
            OrdenProduccion.Lines.UserFields.Fields.Item("U_TipoCafe").Value = RecSet3.Fields.Item("U_Tipo").Value
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Escala").Value = RecSet3.Fields.Item("U_Escala").Value
            objStreamWriter.WriteLine("Recibo: " & RecSet3.Fields.Item("Recibo").Value)
            OrdenProduccion.Lines.ItemNo = RecSet3.Fields.Item("Codigo").Value
            objStreamWriter.WriteLine("ItemNo: " & RecSet3.Fields.Item("Codigo").Value)
            OrdenProduccion.Lines.Warehouse = "07"
            OrdenProduccion.Lines.UserFields.Fields.Item("U_Cantidad").Value = Math.Round(RecSet3.Fields.Item("Cantidad").Value, 2)
            objStreamWriter.WriteLine("Cantidad: " & RecSet3.Fields.Item("Cantidad").Value)
            OrdenProduccion.Lines.BaseQuantity = Math.Round(RecSet3.Fields.Item("Cantidad").Value, 2)
            objStreamWriter.WriteLine("BaseQuantity: " & RecSet3.Fields.Item("Cantidad").Value)
            OrdenProduccion.Lines.Add()
            RecSet3.MoveNext()
        Loop
        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet3)
        RecSet3 = Nothing
        GC.Collect()



        'wo.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased
        'wo.Update()
        oreturn = OrdenProduccion.Update()
        objStreamWriter.WriteLine("Actualizo")
        If oreturn <> 0 Then
            SBO_Application.MessageBox(oCompany.GetLastErrorDescription)
            objStreamWriter.WriteLine("Error: " & oCompany.GetLastErrorDescription)
            objStreamWriter.Close()
        Else
            Dim oreceipt As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)


            Dim RecSet5 As SAPbobsCOM.Recordset
            Dim sql5 As String = ""
            'sql5 = "CALL GET_PROCEDURE(14,'" + DocEntryCata.ToString + "','','','','')"
            sql5 = "CALL RENDIMIENTOS3 ('" + DocEntryCata.ToString + "')"
            objStreamWriter.WriteLine("Trae datos para Recibo: " & sql5)
            RecSet5 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RecSet5.DoQuery(sql5)
            Dim cont As Integer = contadorlineas
            loteproduccionlist.Clear()
            tipocafeproduccionlist.Clear()
            escalaproduccionlist.Clear()
            rendimientobrutoproduccionlist.Clear()
            rendimientonetoproduccionlist.Clear()
            escalarechazoproduccionlist.Clear()
            boletacatacionproduccionlist.Clear()

            Do While Not RecSet5.EoF

                objStreamWriter.WriteLine("Line: " & cont)
                oreceipt.Lines.BaseEntry = DocEntryProduccion
                oreceipt.Lines.BaseLine = cont
                objStreamWriter.WriteLine("Docentryprodcuccion: " & DocEntryProduccion)
                oreceipt.Lines.BaseType = 202
                oreceipt.Lines.Quantity = Math.Round(RecSet5.Fields.Item("Cantidad").Value, 2)
                objStreamWriter.WriteLine("Quantity: " & RecSet5.Fields.Item("Cantidad").Value)
                oreceipt.Lines.WarehouseCode = "07"
                'oreceipt.Lines.COGSCostingCode = "SL4"
                'oreceipt.Lines.TransactionType = SAPbobsCOM.BoTransactionTypeEnum.botrntReject
                loteproduccion = RecSet5.Fields.Item("Lote").Value.ToString
                loteproduccionlist.Add(loteproduccion) 'llena listas para lotes
                tipocafeproduccionlist.Add(RecSet5.Fields.Item("U_Tipo").Value.ToString)
                escalaproduccionlist.Add(RecSet5.Fields.Item("U_Escala").Value.ToString)
                rendimientobrutoproduccionlist.Add(RecSet5.Fields.Item("U_Rend_Bruto").Value.ToString)
                rendimientonetoproduccionlist.Add(RecSet5.Fields.Item("U_Rend_Neto").Value.ToString)
                escalarechazoproduccionlist.Add(RecSet5.Fields.Item("U_Escala_Rechazo").Value.ToString)
                boletacatacionproduccionlist.Add(RecSet5.Fields.Item("Recibo").Value.ToString)
                objStreamWriter.WriteLine("Lote: " & RecSet5.Fields.Item("Lote").Value.ToString)
                oreceipt.Lines.BatchNumbers.BatchNumber = RecSet5.Fields.Item("Lote").Value
                objStreamWriter.WriteLine("Batchnumber: " & RecSet5.Fields.Item("Lote").Value.ToString)
                oreceipt.Lines.BatchNumbers.Quantity = Math.Round(RecSet5.Fields.Item("Cantidad").Value, 2)
                objStreamWriter.WriteLine("Quantity: " & RecSet5.Fields.Item("Cantidad").Value.ToString)
                'oreceipt.Lines.BatchNumbers.Location = "01-A-07"
                'oreceipt.Lines.BatchNumbers.UserFields.Fields.Item("U_Ancho").Value = ancho.Item(cont)
                oreceipt.Lines.BatchNumbers.Add()


                oreceipt.Lines.Add()
                cont = cont + 1
                RecSet5.MoveNext()
            Loop
            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet5)
            RecSet5 = Nothing
            GC.Collect()


            oreturn = oreceipt.Add


            If oreturn <> 0 Then
                objStreamWriter.WriteLine("error: " & oCompany.GetLastErrorDescription)
                SBO_Application.MessageBox("Error, debe de generar la emision de produccion o verificar el error: " & oCompany.GetLastErrorDescription)
                objStreamWriter.Close()
            Else
                objStreamWriter.WriteLine("Finalizo correctamente")
                objStreamWriter.Close()

                'Dim cont As Integer = 0
                For cont = 0 To loteproduccionlist.Count - 1
                    Dim docentry As Integer
                    Dim sqlabs = "CALL GET_PROCEDURE(133,'" + loteproduccionlist.Item(cont) + "','','','','')"
                    Dim RecSet2 As SAPbobsCOM.Recordset
                    RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    RecSet2.DoQuery(sqlabs)
                    If RecSet2.RecordCount > 0 Then
                        docentry = RecSet2.Fields.Item("AbsEntry").Value
                    End If
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                    RecSet2 = Nothing
                    GC.Collect()

                    Dim oCompanyService As SAPbobsCOM.CompanyService
                    oCompanyService = oCompany.GetCompanyService()
                    Dim bNumService As SAPbobsCOM.BatchNumberDetailsService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BatchNumberDetailsService)
                    Dim bNumDetailParams As SAPbobsCOM.BatchNumberDetailParams = bNumService.GetDataInterface(SAPbobsCOM.BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams)
                    bNumDetailParams.DocEntry = docentry
                    Dim bNumDetail As SAPbobsCOM.BatchNumberDetail = bNumService.[Get](bNumDetailParams)
                    bNumDetail.UserFields.Item("U_Escala").Value = escalaproduccionlist.Item(cont)
                    bNumDetail.UserFields.Item("U_TipoCafe").Value = tipocafeproduccionlist.Item(cont)
                    bNumDetail.UserFields.Item("U_Rend_Neto").Value = rendimientonetoproduccionlist.Item(cont)
                    bNumDetail.UserFields.Item("U_Rend_Bruto").Value = rendimientobrutoproduccionlist.Item(cont)
                    bNumDetail.UserFields.Item("U_Escala_Rechazo").Value = escalarechazoproduccionlist.Item(cont)
                    bNumDetail.UserFields.Item("U_Boleta_Catacion").Value = boletacatacionproduccionlist.Item(cont)
                    bNumService.Update(bNumDetail)

                Next
                Dim RecUpdate As SAPbobsCOM.Recordset
                Dim sqlupdate As String = ""
                sqlupdate = "CALL GET_PROCEDURE(20,'" + DocEntryCata.ToString + "','','','','')"
                RecUpdate = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RecUpdate.DoQuery(sqlupdate)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecUpdate)
                RecUpdate = Nothing
                GC.Collect()
                SBO_Application.StatusBar.SetText("Exportacion Generada Exitosamente..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                SBO_Application.MessageBox("Exportacion Generada Exitosamente..")
            End If


        End If
    End Sub

    Public Sub CreaEntradaMerca(ReciboCata, DocEntryCata, cardcode)
        Dim oReturn As Integer = -1
        Dim oError As Integer = 0
        Dim errMsg As String = ""
        Dim Item1 As String = ""
        Dim Item2 As String = ""
        Dim Item3 As String = ""
        Dim Quan1 As Integer
        Dim Quan2 As Integer
        Dim Quan3 As Integer
        'Dim cardcode As String
        'Dim itemcode1 As String
        'Dim itemcode2 As String
        'Dim itemcode3 As String
        'Dim itemprimero As String
        Dim humedad As String = String.Empty
        Dim tipocafe As String = String.Empty
        Dim rendimientobruto As Double
        Dim rendimientoneto As Double
        Dim escala As String = String.Empty
        Dim escalar As String = String.Empty
        Try
#Region "cardcode"
            Dim objStreamWriter As StreamWriter
            'Pass the file path and the file name to the StreamWriter constructor.
            objStreamWriter = New StreamWriter(Application.StartupPath + "\Log_INGRESO" & ReciboCata.ToString & "_" & DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") & ".txt")
            Dim RecSet2 As SAPbobsCOM.Recordset
            Dim sql2 As String = ""

            sql2 = "CALL GET_PROCEDURE(3,'" + ReciboCata.ToString + "','','','','')"
            objStreamWriter.WriteLine("LLena datos: " & sql2)


            RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RecSet2.DoQuery(sql2)
            Dim oWIn As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

            If RecSet2.RecordCount > 0 Then

                oWIn.DocDate = Date.Now.ToShortDateString
                oWIn.DocDueDate = Date.Now.AddDays(30).ToShortDateString
                oWIn.TaxDate = Date.Now.ToShortDateString
                oWIn.CardCode = cardcode
                objStreamWriter.WriteLine("Cardcode: " & cardcode)
                oWIn.Lines.ItemCode = RecSet2.Fields.Item("U_Tipo_Cafe").Value
                objStreamWriter.WriteLine("Itemcode: " & RecSet2.Fields.Item("U_Tipo_Cafe").Value)
                oWIn.Lines.Quantity = RecSet2.Fields.Item("U_Quintales").Value
                objStreamWriter.WriteLine("Quantity" & RecSet2.Fields.Item("U_Quintales").Value)
                oWIn.Lines.WarehouseCode = RecSet2.Fields.Item("WhsCode").Value
                objStreamWriter.WriteLine("WhsCode" & RecSet2.Fields.Item("WhsCode").Value)
                oWIn.Lines.AccountCode = RecSet2.Fields.Item("AccountCode").Value
                objStreamWriter.WriteLine("Account" & RecSet2.Fields.Item("AccountCode").Value)
                'oWIn.Lines.WarehouseCode = "01"
                oWIn.Comments = "Creado desde Addon"
                oWIn.JournalMemo = "Creado desde Addon"
                oWIn.Lines.BatchNumbers.BatchNumber = ReciboCata.ToString
                objStreamWriter.WriteLine("Lote" & ReciboCata.ToString)
                oWIn.Lines.BatchNumbers.Quantity = RecSet2.Fields.Item("U_Quintales").Value
                objStreamWriter.WriteLine("Quantity" & RecSet2.Fields.Item("U_Quintales").Value)
                'oWIn.Lines.BatchNumbers.RequiredDate = Date.Now.ToShortDateString
                humedad = RecSet2.Fields.Item("Humedad").Value.ToString
                tipocafe = RecSet2.Fields.Item("U_Tipo").Value.ToString
                escala = RecSet2.Fields.Item("Escala").Value.ToString
                escalar = RecSet2.Fields.Item("EscalaR").Value.ToString
                rendimientobruto = Convert.ToDouble(RecSet2.Fields.Item("RendimientoBruto").Value.ToString)
                rendimientoneto = Convert.ToDouble(RecSet2.Fields.Item("RendimientoNeto").Value.ToString)
                'oWIn.Lines.BatchNumbers.UserFields.Fields.Item("U_BoletaCatacion").Value = DocEntryCata
                'oWIn.Lines.BatchNumbers.UserFields.Fields.Item("U_Rendimiento_Neto").Value = humedad
                'oWIn.Lines.BatchNumbers.UserFields.Fields.Item("U_Humedad").Value = bruto
                'oWIn.Lines.BatchNumbers.UserFields.Fields.Item("U_Humedad").Value = neto
                'oWIn.Lines.BatchNumbers.Add()
                oWIn.Lines.Add()
                ''''''
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
            RecSet2 = Nothing
            GC.Collect()
#End Region

            Dim RecSet As SAPbobsCOM.Recordset
            Dim sql As String = ""

            sql = "CALL GET_PROCEDURE(1,'" + ReciboCata + "','','','','')"
            objStreamWriter.WriteLine("Para items: " & sql)
            RecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            RecSet.DoQuery(sql)
            If RecSet.RecordCount > 0 Then
                Item1 = RecSet.Fields.Item("Item1").Value
                Item2 = RecSet.Fields.Item("Item2").Value
                Item3 = RecSet.Fields.Item("Item3").Value
                Quan1 = RecSet.Fields.Item("Quan1").Value
                Quan2 = RecSet.Fields.Item("Quan2").Value
                Quan3 = RecSet.Fields.Item("Quan3").Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet)
            RecSet = Nothing
            GC.Collect()
            objStreamWriter.WriteLine("item1" & Item1)
            If Item1 <> "" Then
                oWIn.Lines.ItemCode = Item1
                oWIn.Lines.Quantity = Quan1
                oWIn.Lines.WarehouseCode = "07"
                oWIn.Lines.AccountCode = "110403157"

                oWIn.Lines.BatchNumbers.BatchNumber = ReciboCata
                oWIn.Lines.BatchNumbers.Quantity = Quan1
                ' oWIn.Lines.BatchNumbers.RequiredDate = Date.Now.ToShortDateString
                oWIn.Lines.BatchNumbers.Add()
                oWIn.Lines.Add()
            End If
            objStreamWriter.WriteLine("item2" & Item2)
            If Item2 <> "" Then
                oWIn.Lines.ItemCode = Item2
                oWIn.Lines.Quantity = Quan2
                oWIn.Lines.WarehouseCode = "07"
                oWIn.Lines.AccountCode = "110403157"

                oWIn.Lines.BatchNumbers.BatchNumber = ReciboCata
                oWIn.Lines.BatchNumbers.Quantity = Quan2
                'oWIn.Lines.BatchNumbers.RequiredDate = Date.Now.ToShortDateString
                oWIn.Lines.BatchNumbers.Add()
                oWIn.Lines.Add()
            End If
            objStreamWriter.WriteLine("item3" & Item3)
            If Item3 <> "" Then
                oWIn.Lines.ItemCode = Item3
                oWIn.Lines.Quantity = Quan3
                oWIn.Lines.WarehouseCode = "07"
                oWIn.Lines.AccountCode = "110403157"

                oWIn.Lines.BatchNumbers.BatchNumber = ReciboCata
                oWIn.Lines.BatchNumbers.Quantity = Quan3
                'oWIn.Lines.BatchNumbers.RequiredDate = Date.Now.ToShortDateString
                oWIn.Lines.BatchNumbers.Add()
                oWIn.Lines.Add()
            End If
            oReturn = oWIn.Add()
            If oReturn <> 0 Then
                objStreamWriter.WriteLine(oCompany.GetLastErrorDescription)
                SBO_Application.MessageBox("Genere Entrada de Mercaderia Manualmente y Actualize #de entrada en catacion, error: " & oCompany.GetLastErrorDescription)
            Else
                Dim entry As Integer
                entry = oCompany.GetNewObjectKey()
                Dim docnum As String = String.Empty
                Dim RecSets As SAPbobsCOM.Recordset
                Dim sqls As String = ""
                sqls = "CALL GET_PROCEDURE(7,'" + entry.ToString + "','','','','')"
                objStreamWriter.WriteLine("para actualizar cata: " & sqls)
                RecSets = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RecSets.DoQuery(sqls)
                If RecSets.RecordCount > 0 Then
                    docnum = RecSets.Fields.Item("DocNum").Value
                    objStreamWriter.WriteLine(docnum)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSets)
                RecSets = Nothing
                GC.Collect()

                sql = "CALL GET_PROCEDURE(4,'" + ReciboCata.ToString + "','" + docnum.ToString + "','','','')"
                objStreamWriter.WriteLine("Update entrada merc: " & sql)
                RecSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RecSet.DoQuery(sql)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet)
                RecSet = Nothing
                GC.Collect()


                Dim docentry As Integer
                Dim sqlabs = "CALL GET_PROCEDURE(13,'" + ReciboCata + "','','','','')"
                objStreamWriter.WriteLine("Trae absentry" & sqlabs)
                RecSet2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                RecSet2.DoQuery(sqlabs)
                Do While Not RecSet2.EoF
                    docentry = RecSet2.Fields.Item("AbsEntry").Value
                    objStreamWriter.WriteLine("absentry: " & docentry.ToString)
                    Dim oCompanyService As SAPbobsCOM.CompanyService
                    oCompanyService = oCompany.GetCompanyService()
                    Dim bNumService As SAPbobsCOM.BatchNumberDetailsService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BatchNumberDetailsService)
                    Dim bNumDetailParams As SAPbobsCOM.BatchNumberDetailParams = bNumService.GetDataInterface(SAPbobsCOM.BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams)
                    bNumDetailParams.DocEntry = docentry
                    Dim bNumDetail As SAPbobsCOM.BatchNumberDetail = bNumService.[Get](bNumDetailParams)

                    bNumDetail.UserFields.Item("U_Escala_Rechazo").Value = escalar
                    bNumDetail.UserFields.Item("U_Escala").Value = escala
                    bNumDetail.UserFields.Item("U_Humedad").Value = humedad
                    bNumDetail.UserFields.Item("U_TipoCafe").Value = tipocafe
                    bNumDetail.UserFields.Item("U_Rend_Bruto").Value = rendimientobruto
                    bNumDetail.UserFields.Item("U_Rend_Neto").Value = rendimientoneto
                    bNumDetail.UserFields.Item("U_Boleta_Catacion").Value = DocEntryCata.ToString()
                    bNumService.Update(bNumDetail)
                    RecSet2.MoveNext()
                Loop
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecSet2)
                RecSet2 = Nothing
                GC.Collect()


                SBO_Application.MessageBox("Entrada Generada Exitosamente..")
            End If
            objStreamWriter.WriteLine("Finalizo correctamente")

            objStreamWriter.Close()
        Catch ex As Exception

            SBO_Application.MessageBox("Error: " + ex.Message.ToString)
        End Try
    End Sub

    Private Sub InstanciarObjetos(ByVal pVal As SAPbouiCOM.ItemEvent)
        If (pVal.FormTypeEx = "60006") Then

            instanciado = True

            oForm = SBO_Application.Forms.ActiveForm
            txtRecibo = oForm.Items.Item("Item_24").Specific
            cmbEstado = oForm.Items.Item("Item_25").Specific
            txtCodFinca = oForm.Items.Item("Item_47").Specific
            txtTipoOT = oForm.Items.Item("Item_90").Specific
            txtFinca = oForm.Items.Item("34_U_E").Specific
            cmbCatado = oForm.Items.Item("Item_110").Specific
            txtCodCafe = oForm.Items.Item("Item_3").Specific
            txtCantidad = oForm.Items.Item("23_U_E").Specific
            txtTipo = oForm.Items.Item("Item_6").Specific
            txtEscala = oForm.Items.Item("Item_8").Specific
            txtEscalaRechazo = oForm.Items.Item("Item_80").Specific
            cmbVerde = oForm.Items.Item("24_U_Cb").Specific
            cmbTueste = oForm.Items.Item("25_U_Cb").Specific
            txtComVerde = oForm.Items.Item("Item_5").Specific
            txtComTueste = oForm.Items.Item("Item_21").Specific
            chkTzSana = oForm.Items.Item("Item_9").Specific
            chkTzFermentada = oForm.Items.Item("Item_7").Specific
            chkTzMohosa = oForm.Items.Item("Item_10").Specific
            chkTzFruti = oForm.Items.Item("Item_11").Specific
            chkTzTerrosa = oForm.Items.Item("Item_12").Specific
            chkTzFenolica = oForm.Items.Item("Item_13").Specific
            chkTzAgria = oForm.Items.Item("Item_14").Specific
            chkTzVinosa = oForm.Items.Item("Item_15").Specific
            chkTzAspera = oForm.Items.Item("Item_141").Specific
            chkTzSucia = oForm.Items.Item("Item_146").Specific
            txtComTaza = oForm.Items.Item("Item_17").Specific
            cmbPargo = oForm.Items.Item("28_U_Cb").Specific
            txtComPargo = oForm.Items.Item("Item_19").Specific
            txtHumedad = oForm.Items.Item("30_U_E").Specific
            txtRendBruto = oForm.Items.Item("31_U_E").Specific
            txtRendNeto = oForm.Items.Item("32_U_E").Specific
            txtGrsLtrs = oForm.Items.Item("Item_40").Specific
            txtRendNetoE = oForm.Items.Item("Item_44").Specific
            txtRO = oForm.Items.Item("Item_115").Specific
            txtESC = oForm.Items.Item("Item_117").Specific
            txtBZ = oForm.Items.Item("Item_119").Specific
            txtSZ = oForm.Items.Item("Item_121").Specific
            txtScore = oForm.Items.Item("Item_75").Specific
            txtEntrada = oForm.Items.Item("Item_78").Specific
            txtTotal = oForm.Items.Item("Item_113").Specific
            txtObservaciones1 = oForm.Items.Item("33_U_E").Specific
            txtObservaciones2 = oForm.Items.Item("Item_106").Specific
            txtObservaciones3 = oForm.Items.Item("Item_107").Specific

            btnProceso = oForm.Items.Item("138").Specific
            btnImprimirCatacion = oForm.Items.Item("Item_0").Specific
            btnAceptar = oForm.Items.Item("1").Specific

            tblOrdenTrabajo = oForm.Items.Item("Item_86").Specific
            txtCantZar18 = oForm.Items.Item("Item_137").Specific
            txtCantZar17 = oForm.Items.Item("Item_122").Specific
            txtCantZar16 = oForm.Items.Item("Item_135").Specific
            txtCantZar15 = oForm.Items.Item("Item_134").Specific
            txtCantZar14 = oForm.Items.Item("Item_128").Specific
            txtCantZar13 = oForm.Items.Item("Item_136").Specific
            txtPorcZar18 = oForm.Items.Item("Item_145").Specific
            txtPorcZar17 = oForm.Items.Item("Item_138").Specific
            txtPorcZar16 = oForm.Items.Item("Item_143").Specific
            txtPorcZar15 = oForm.Items.Item("Item_142").Specific
            txtPorcZar14 = oForm.Items.Item("Item_139").Specific
            txtPorcZar13 = oForm.Items.Item("Item_144").Specific
            txtFondo = oForm.Items.Item("Item_132").Specific
            txtVarPrueba = oForm.Items.Item("Item_130").Specific


            txtCantNegro = oForm.Items.Item("Item_95").Specific
            txtDefcNegro = oForm.Items.Item("Item_98").Specific
            txtCantFermentado = oForm.Items.Item("Item_36").Specific
            txtDefcFermentado = oForm.Items.Item("Item_100").Specific
            txtCantCerezaSeca = oForm.Items.Item("Item_55").Specific
            txtDefcCerezaSeca = oForm.Items.Item("Item_53").Specific
            txtCantDanioHongo = oForm.Items.Item("Item_32").Specific
            txtDefcDanioHongo = oForm.Items.Item("Item_101").Specific
            txtCantMatExtrania = oForm.Items.Item("Item_92").Specific
            txtDefcMatExtrania = oForm.Items.Item("Item_87").Specific
            txtCantNegroParc = oForm.Items.Item("Item_96").Specific
            txtDefcNegroParc = oForm.Items.Item("Item_103").Specific
            txtCantParcFermn = oForm.Items.Item("Item_91").Specific
            txtDefcParcFermn = oForm.Items.Item("Item_82").Specific
            txtCantPergamino = oForm.Items.Item("Item_97").Specific
            txtDefcPergamino = oForm.Items.Item("Item_104").Specific
            txtCantFlotador = oForm.Items.Item("Item_34").Specific
            txtDefcFlotador = oForm.Items.Item("Item_60").Specific
            txtCantInmaduro = oForm.Items.Item("Item_93").Specific
            txtDefcInmaduro = oForm.Items.Item("Item_102").Specific
            txtCantAveranado = oForm.Items.Item("Item_94").Specific
            txtDefcAveranado = oForm.Items.Item("Item_88").Specific
            txtCantConcha = oForm.Items.Item("Item_49").Specific
            txtDefcConcha = oForm.Items.Item("Item_105").Specific
            txtCantMordido = oForm.Items.Item("Item_28").Specific
            txtDefcMordido = oForm.Items.Item("Item_81").Specific
            txtCantCascaraSeca = oForm.Items.Item("Item_51").Specific
            txtDefcCascaraSeca = oForm.Items.Item("Item_99").Specific
            txtCantBrocado = oForm.Items.Item("Item_30").Specific
            txtDefcBrocado = oForm.Items.Item("Item_62").Specific

            cmbFragancia = oForm.Items.Item("Item_23").Specific
            txtComFragancia = oForm.Items.Item("Item_66").Specific
            cmbAcidez = oForm.Items.Item("Item_63").Specific
            txtComAcidez = oForm.Items.Item("Item_67").Specific
            cmbCuerpo = oForm.Items.Item("Item_64").Specific
            txtComCuerpo = oForm.Items.Item("Item_68").Specific
            cmbSabor = oForm.Items.Item("Item_65").Specific
            txtComSabor = oForm.Items.Item("Item_69").Specific
            txtComentarioFinal = oForm.Items.Item("Item_71").Specific
            chkBuenoEmbarque = oForm.Items.Item("Item_76").Specific

            txtVarPrueba = oForm.Items.Item("Item_130").Specific
        End If
    End Sub

    Private Sub ActivarFormulario(Activar As Boolean)
        'txtCodFinca.Enable = Activar
        'txtTipoOT.Enable = Activar
        'txtFinca.Enable = Activar

        cmbCatado.Item.Enabled = Activar
        txtCodCafe.Item.Enabled = Activar
        txtCantidad.Item.Enabled = Activar
        txtTipo.Item.Enabled = Activar
        txtEscala.Item.Enabled = Activar
        txtEscalaRechazo.Item.Enabled = Activar
        cmbVerde.Item.Enabled = Activar
        cmbTueste.Item.Enabled = Activar
        txtComVerde.Item.Enabled = Activar
        txtComTueste.Item.Enabled = Activar

        cmbPargo.Item.Enabled = Activar
        txtComPargo.Item.Enabled = Activar
        txtHumedad.Item.Enabled = Activar
        txtRendBruto.Item.Enabled = Activar
        txtRendNeto.Item.Enabled = Activar
        txtGrsLtrs.Item.Enabled = Activar
        txtRendNetoE.Item.Enabled = Activar


        tblOrdenTrabajo.Item.Enabled = Activar
        ActivarTaza(Activar)
        ActivarComentarios(Activar)
        ActivarZarandra(Activar)
        ActivarDefectos(Activar)
        ActivarFormularioFinal(Activar)
        ActivarCalificacion(Activar)

        btnProceso.Item.Enabled = Activar
    End Sub

    Private Sub ActivarTaza(Activar As Boolean)
        chkTzSana.Item.Enabled = Activar
        chkTzFermentada.Item.Enabled = Activar
        chkTzMohosa.Item.Enabled = Activar
        chkTzFruti.Item.Enabled = Activar
        chkTzTerrosa.Item.Enabled = Activar
        chkTzFenolica.Item.Enabled = Activar
        chkTzAgria.Item.Enabled = Activar
        chkTzVinosa.Item.Enabled = Activar
        txtComTaza.Item.Enabled = Activar
        chkTzSucia.Item.Enabled = Activar
        chkTzAspera.Item.Enabled = Activar
    End Sub

    Private Sub ActivarComentarios(Activar As Boolean)
        txtObservaciones1.Item.Enabled = Activar
        txtObservaciones2.Item.Enabled = Activar
        txtObservaciones3.Item.Enabled = Activar
    End Sub

    Private Sub ActivarZarandra(Activar As Boolean)
        txtCantZar13.Item.Enabled = Activar
        txtCantZar14.Item.Enabled = Activar
        txtCantZar15.Item.Enabled = Activar
        txtCantZar16.Item.Enabled = Activar
        txtCantZar17.Item.Enabled = Activar
        txtCantZar18.Item.Enabled = Activar
        txtPorcZar13.Item.Enabled = Activar
        txtPorcZar14.Item.Enabled = Activar
        txtPorcZar15.Item.Enabled = Activar
        txtPorcZar16.Item.Enabled = Activar
        txtPorcZar17.Item.Enabled = Activar
        txtPorcZar18.Item.Enabled = Activar
        txtFondo.Item.Enabled = Activar
        txtVarPrueba.Item.Enabled = Activar
    End Sub

    Private Sub ActivarFormularioFinal(Activar As Boolean)
        cmbFragancia.Item.Enabled = Activar
        txtComFragancia.Item.Enabled = Activar
        cmbAcidez.Item.Enabled = Activar
        txtComAcidez.Item.Enabled = Activar
        cmbCuerpo.Item.Enabled = Activar
        txtComCuerpo.Item.Enabled = Activar
        cmbSabor.Item.Enabled = Activar
        txtComSabor.Item.Enabled = Activar
        txtComentarioFinal.Item.Enabled = Activar
        chkBuenoEmbarque.Item.Enabled = Activar
    End Sub

    Private Sub ActivarDefectos(Activar As Boolean)
        txtCantNegro.Item.Enabled = Activar
        txtDefcNegro.Item.Enabled = Activar
        txtCantFermentado.Item.Enabled = Activar
        txtDefcFermentado.Item.Enabled = Activar
        txtCantCerezaSeca.Item.Enabled = Activar
        txtDefcCerezaSeca.Item.Enabled = Activar
        txtCantDanioHongo.Item.Enabled = Activar
        txtDefcDanioHongo.Item.Enabled = Activar
        txtCantMatExtrania.Item.Enabled = Activar
        txtDefcMatExtrania.Item.Enabled = Activar
        txtCantNegroParc.Item.Enabled = Activar
        txtDefcNegroParc.Item.Enabled = Activar
        txtCantParcFermn.Item.Enabled = Activar
        txtDefcParcFermn.Item.Enabled = Activar
        txtCantPergamino.Item.Enabled = Activar
        txtDefcPergamino.Item.Enabled = Activar
        txtCantFlotador.Item.Enabled = Activar
        txtDefcFlotador.Item.Enabled = Activar
        txtCantInmaduro.Item.Enabled = Activar
        txtDefcInmaduro.Item.Enabled = Activar
        txtCantAveranado.Item.Enabled = Activar
        txtDefcAveranado.Item.Enabled = Activar
        txtCantConcha.Item.Enabled = Activar
        txtDefcConcha.Item.Enabled = Activar
        txtCantMordido.Item.Enabled = Activar
        txtDefcMordido.Item.Enabled = Activar
        txtCantCascaraSeca.Item.Enabled = Activar
        txtDefcCascaraSeca.Item.Enabled = Activar
        txtCantBrocado.Item.Enabled = Activar
        txtDefcBrocado.Item.Enabled = Activar
    End Sub

    Private Sub ActivarCalificacion(Activar As Boolean)
        txtRO.Item.Enabled = Activar
        txtESC.Item.Enabled = Activar
        txtBZ.Item.Enabled = Activar
        txtSZ.Item.Enabled = Activar
        txtScore.Item.Enabled = Activar
        txtEntrada.Item.Enabled = Activar
        txtTotal.Item.Enabled = True
    End Sub

    Private Function GetCalidadCafe(ByVal TipoCafe As String) As String
        Dim response As String = String.Empty
        Dim rsTipoCafe As SAPbobsCOM.Recordset = Nothing
        rsTipoCafe = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rsTipoCafe.DoQuery($"SELECT CASE T0.""U_Tipo"" WHEN 'PRIMERA' THEN 'PTPRIMERA' ELSE 'INFERIORES' END AS ""Calidad"" FROM ""@TIPOS_CAFE"" T0 WHERE T0.""Code""='{TipoCafe}'")
        If rsTipoCafe.RecordCount > 0 Then
            response = rsTipoCafe.Fields.Item("Calidad").Value.ToString()
        Else
            response = String.Empty
        End If
        rsTipoCafe = Nothing
        GC.Collect()
        Return response
    End Function

    Private Function ValidarCantidades(ByVal NumeroOT As Integer, ByVal CantidadQQOT As Double) As Boolean
        Dim response As Boolean = False
        Dim rs As SAPbobsCOM.Recordset = Nothing
        Dim query As StringBuilder = Nothing
        query = New StringBuilder()
        query.Append("SELECT").AppendLine()
        query.Append("SUM(IFNULL(T0.""U_Quintales"",0))AS ""Acumulado""").AppendLine()
        query.Append(",IFNULL(T1.""U_Cantidad_qq"",0) AS ""QuintalesOT""").AppendLine()
        query.Append("FROM ""@CATACION"" T0").AppendLine()
        query.Append("INNER JOIN ""OWOR"" T1 ON CAST(T0.""U_DocEntry"" AS NUMERIC)=T1.""DocNum""").AppendLine()
        query.Append("WHERE").AppendLine()
        query.Append($"T0.""U_DocEntry""='{NumeroOT}'")
        query.Append("GROUP BY T1.""U_Cantidad_qq""")

        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        rs.DoQuery(query.ToString())
        If rs.RecordCount > 0 Then
            If rs.Fields.Item("Respuesta").Value.ToString() = "0" Then
                response = False
            Else
                response = True
            End If
        Else
            response = False
        End If
        rs = Nothing
        GC.Collect()
        Return response
    End Function

    'Private Sub imprimeCatacion(ByVal dockey As Integer)
    '    Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
    '    Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
    '    Report1.Load(Application.StartupPath + "\Boletacatacioningreso.rpt", OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
    '    'Report1.SetDatabaseLogon("USERSAP", "Cafcom18", "10.1.1.202:30015", "cafcom")
    '    ''-----------------------------------------ENCABEZADO NO CAMBIA POR IMPRESION------------------------------------------

    '    Dim boConnectionInfo As ConnectionInfo = New ConnectionInfo

    '    boConnectionInfo.ServerName = "10.1.1.202:30015"

    '    boConnectionInfo.DatabaseName = "CAFCOM"

    '    boConnectionInfo.UserID = "USERSAP"

    '    boConnectionInfo.Password = "Cafcom18"

    '    boConnectionInfo.Type = ConnectionInfoType.Unknown

    '    For Each t As Table In Report1.Database.Tables

    '        Dim boTableLogOnInfo As TableLogOnInfo = t.LogOnInfo

    '        boTableLogOnInfo.ConnectionInfo = boConnectionInfo

    '        t.ApplyLogOnInfo(boTableLogOnInfo)

    '    Next
    '    Dim doc As Integer
    '    doc = Convert.ToInt32(dockey)
    '    Report1.SetParameterValue("DocKey@", doc)
    '    'Report1.SetParameterValue("UserCode@", 1)
    '    Report1.PrintToPrinter(1, False, 0, 0)
    '    ''Report1.SaveAs("impresion")


    '    Dim diskFileDestinationOptions As DiskFileDestinationOptions = New DiskFileDestinationOptions()
    '    diskFileDestinationOptions.DiskFileName = Application.StartupPath + "\Catacion_Orden_" & dockey & "_" & DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") & ".pdf"

    '    Dim exportOptions As ExportOptions = New ExportOptions()
    '    exportOptions.ExportDestinationType = ExportDestinationType.DiskFile
    '    exportOptions.ExportFormatOptions = Nothing
    '    exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
    '    exportOptions.ExportDestinationOptions = diskFileDestinationOptions
    '    Report1.Export(exportOptions)
    '    'SBO_Application.StatusBar.SetText("Impresion de ticket: " + doc.ToString + " Exitosa..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    '    Dim obj As New Process
    '    obj.Start(diskFileDestinationOptions.DiskFileName.ToString, AppWinStyle.MaximizedFocus)
    '    'Shell(diskFileDestinationOptions.DiskFileName.ToString)
    'End Sub

    'Private Sub imprime(ByVal dockey As Integer)
    '    Dim Report1 As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
    '    Report1.PrintOptions.PaperOrientation = PaperOrientation.Portrait
    '    Report1.Load(Application.StartupPath + "\Boleta_de_pesaje.rpt", OpenReportMethod.OpenReportByDefault.OpenReportByDefault)
    '    Report1.SetDatabaseLogon("USERSAP", "Cafcom18", "10.1.1.202:30015", "cafcom")
    '    ' Report1.SetDatabaseLogon("USERSAP", "Cafcom18")
    '    'SetTableLocation(Report1.Database.Tables)
    '    ''-----------------------------------------ENCABEZADO NO CAMBIA POR IMPRESION------------------------------------------
    '    Dim doc As Integer
    '    doc = Convert.ToInt32(dockey)
    '    Report1.SetParameterValue("DocKey@", doc)
    '    Report1.SetParameterValue("UserCode@", 1)
    '    'Report1.PrintToPrinter(1, False, 0, 0)
    '    ''Report1.SaveAs("impresion")


    '    Dim diskFileDestinationOptions As DiskFileDestinationOptions = New DiskFileDestinationOptions()
    '    diskFileDestinationOptions.DiskFileName = Application.StartupPath + "\Impresion_" & dockey & "_" & DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss") & ".pdf"

    '    Dim exportOptions As ExportOptions = New ExportOptions()
    '    exportOptions.ExportDestinationType = ExportDestinationType.DiskFile
    '    exportOptions.ExportFormatOptions = Nothing
    '    exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
    '    exportOptions.ExportDestinationOptions = diskFileDestinationOptions

    '    Report1.Export(exportOptions)
    '    SBO_Application.StatusBar.SetText("Impresion de ticket: " + doc + " Exitosa..", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    'End Sub
End Class
