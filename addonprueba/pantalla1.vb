
    Imports System.Globalization.CultureInfo
Public Class pantalla1
    Dim XmlForm As String = Replace(System.Windows.Forms.Application.StartupPath & "\pantalla1.srf", "\\", "\")

    Private WithEvents SBO_Application As SAPbouiCOM.Application
    Private oForm As SAPbouiCOM.Form
    Private oCompany As SAPbobsCOM.Company
    Private oFilters As SAPbouiCOM.EventFilters
    Private oFilter As SAPbouiCOM.EventFilter
    Dim lineinioriginal As Integer
    Dim linefinoriginal As Integer
    Dim oGrid As SAPbouiCOM.Grid

    Public Sub New()
        MyBase.New()
        Try
            Me.SBO_Application = Utilss.SBOApplication
            Me.oCompany = Utilss.Company

            If Utilss.ActivateFormIsOpen(SBO_Application, "FrmValor") = False Then
                LoadFromXML(XmlForm)
                oForm = SBO_Application.Forms.Item("FrmValor")
                oForm.Left = 400
                oForm.DataSources.DataTables.Add("MyDataTable")
                oGrid = oForm.Items.Item("grdDatos").Specific
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Else
                oForm = Me.SBO_Application.Forms.Item("FrmValor")
                oForm.Left = 400
                oForm.Visible = true
            End If

        Catch ex As Exception
            SBOApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Sub

    Private Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        oXmlDoc.Load(FileName)
        SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
    End Sub

    Private Sub LlenaGrid(valor As String)
        Try
            Dim QryStr As String


            QryStr = (String.Format("select Itemcode,(LineNum + 1) 'Linea', Dscription 'Descripcion', Quantity 'Cantidad', Price 'Precio', DiscPrcnt 'Descuento' from QUT1 where DocEntry = '{0}'", valor))
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(QryStr)
            oGrid = oForm.Items.Item("grdDatos").Specific
            oGrid.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.Columns.GetEnumerator()
            CType(oGrid.Columns.Item(0), SAPbouiCOM.EditTextColumn).LinkedObjectType = 4
            linefinoriginal = oGrid.Rows.Count.ToString()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid)
            oGrid = Nothing
            GC.Collect()
        Catch ex As Exception
            'SBO_Application.MessageBox(ex.Message)
        End Try
    End Sub


    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent

        Try
            If pVal.FormUID = "FrmValor" Then
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST Then
                    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                    oCFLEvento = pVal
                    Dim sCFL_ID As String
                    sCFL_ID = oCFLEvento.ChooseFromListUID
                    Dim oForm As SAPbouiCOM.Form
                    oForm = SBO_Application.Forms.Item(FormUID)
                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    oCFL = oForm.ChooseFromLists.Item(sCFL_ID)
                    If oCFLEvento.BeforeAction = False Then
                        Dim oDataTable As SAPbouiCOM.DataTable
                        oDataTable = oCFLEvento.SelectedObjects
                        Dim val As String


                        If (pVal.ItemUID = "Item_0") Then
                            Try
                                Dim txtFactura As SAPbouiCOM.EditText = oForm.Items.Item("Item_0").Specific
                                val = oDataTable.GetValue("DocEntry", 0)
                                LlenaGrid(val)
                                txtFactura.Value = val
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If

            End If


            If pVal.ItemUID = "cmdOk" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then

                SBO_Application.SetStatusBarMessage("Mensaje en Barra de Estatus", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                Dim resp = SBO_Application.MessageBox("Mensaje Flotante")
                BubbleEvent = False
                Return
            End If

            If pVal.ItemUID = "btnCancel" And pVal.FormUID = "FrmValor" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                Try
                    oForm.Close
                    BubbleEvent = False
                    Return
                Catch ex As Exception

                End Try

            End If
        Catch ex As Exception
            SBO_Application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            BubbleEvent = False
            Return
        End Try
    End Sub
End Class
