﻿Imports System.Windows.Forms

Friend Class CatchingEvents

    Friend WithEvents SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Friend SBOCompany As SAPbobsCOM.Company '//OBJETO COMPAÑIA
    Friend csDirectory As String '//DIRECTORIO DONDE SE ENCUENTRAN LOS .SRF
    Dim DocNum, Comments, Fecha, CardCode As String
    Dim DocEntry As Integer
    Dim DocTotal As Double

    Public Sub New()
        MyBase.New()
        SetAplication()
        SetConnectionContext() 
        ConnectSBOCompany()

        setFilters()

    End Sub

    '//----- ESTABLECE LA COMUNICACION CON SBO
    Private Sub SetAplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try 
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)
            SBOApplication = SboGuiApi.GetApplication()
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la aplicación SBO " & ex.Message)
            System.Windows.Forms.Application.Exit() 
            End
        End Try
    End Sub

    '//----- ESTABLECE EL CONTEXTO DE LA APLICACION
    Private Sub SetConnectionContext()
        Try
            SBOCompany = SBOApplication.Company.GetDICompany
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con el DI")
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
            'Finally
        End Try
    End Sub

    '//----- CONEXION CON LA BASE DE DATOS
    Private Sub ConnectSBOCompany()
        Dim loRecSet As SAPbobsCOM.Recordset
        Try
            '//ESTABLECE LA CONEXION A LA COMPAÑIA
            csDirectory = My.Application.Info.DirectoryPath
            If (csDirectory = "") Then
                System.Windows.Forms.Application.Exit()
                End
            End If
        Catch ex As Exception
            SBOApplication.MessageBox("Falló la conexión con la BD. " & ex.Message)
            SBOApplication = Nothing
            System.Windows.Forms.Application.Exit()
            End
        Finally
            loRecSet = Nothing
        End Try
    End Sub


    '//----- ESTABLECE FILTROS DE EVENTOS DE LA APLICACION
    Private Sub setFilters()
        Dim lofilter As SAPbouiCOM.EventFilter
        Dim lofilters As SAPbouiCOM.EventFilters

        Try

            lofilters = New SAPbouiCOM.EventFilters
            lofilter = lofilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            lofilter.AddEx(170) '////// FORMA PAGOS RECIBIDOS

            SBOApplication.SetFilter(lofilters)

        Catch ex As Exception
            SBOApplication.MessageBox("SetFilter: " & ex.Message)
        End Try

    End Sub


    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    ''// METODOS PARA EVENTOS DE LA APLICACION
    ''//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBOApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                System.Windows.Forms.Application.Exit()
                End
        End Select

    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// METODOS PARA MANEJO DE EVENTOS ITEM
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub SBOApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBOApplication.ItemEvent

        Try

            ''SBOApplication.MessageBox("Action: " & pVal.Before_Action & "  Type: " & pVal.FormTypeEx)
            If pVal.Before_Action = True And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case 170
                        frmPOControllerBefore(FormUID, pVal)
                End Select
            End If

            If pVal.Before_Action = False And pVal.FormTypeEx <> "" Then

                Select Case pVal.FormTypeEx
                    '////////////////FORMA PARA ACTIVAR LICENCIA
                    Case 170
                        frmPOControllerAfter(FormUID, pVal)
                End Select
            End If

        Catch ex As Exception

            SBOApplication.MessageBox("Error SBOApplication_ItemEvent: " & ex.Message)

        End Try

    End Sub


    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA PEDIDOS DE COMPRAS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmPOControllerBefore(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        'Dim oPO As PO
        'Dim otekPagos As FrmtekPagos
        Dim coForm As SAPbouiCOM.Form
        Dim stTabla As String
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim oDatatable As SAPbouiCOM.DBDataSource
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case 1

                            stTabla = "ORCT"
                            coForm = SBOApplication.Forms.Item(FormUID)

                            oDatatable = coForm.DataSources.DBDataSources.Item(stTabla)
                            DocNum = oDatatable.GetValue("DocNum", 0)
                            Comments = oDatatable.GetValue("Comments", 0)
                            Fecha = oDatatable.GetValue("DocDate", 0)
                            DocTotal = oDatatable.GetValue("DocTotal", 0)
                            CardCode = oDatatable.GetValue("CardCode", 0)

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '// CONTROLADOR DE EVENTOS FORMA PEDIDOS DE COMPRAS
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Private Sub frmPOControllerAfter(ByVal FormUID As String, ByVal pVal As SAPbouiCOM.ItemEvent)
        'Dim oPO As PO
        'Dim otekPagos As FrmtekPagos
        Dim oOBNK As OBNK
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            Select Case pVal.EventType

                                '//////Evento Presionar Item
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    Select Case pVal.ItemUID
                                    '--- Boton Movimientos del Pedido
                        Case 1

                            stQueryH = "Select Top 1 T1.""DocEntry"" from ORCT T0 inner join RCT2 T1 on T0.""DocNum""=T1.""DocNum"" and T1.""InvType""=13 where T0.""DocNum""=" & DocNum & " order by T1.""DocEntry"" asc"
                            oRecSetH.DoQuery(stQueryH)

                            If oRecSetH.RecordCount > 0 Then

                                oRecSetH.MoveFirst()
                                DocEntry = oRecSetH.Fields.Item("DocEntry").Value

                                'SBOApplication.MessageBox("Cuenta:" & Comments & " Fecha:" & Fecha & " Total:" & DocTotal & " Cliente:" & CardCode & " DocENtry:" & DocEntry)

                                oOBNK = New OBNK
                                oOBNK.UpdateOBNK(Comments, Fecha, DocTotal, CardCode, DocEntry)

                            End If

                    End Select

            End Select

        Catch ex As Exception
            SBOApplication.MessageBox("Error en el evento sobre Forma Pedido de Compras. " & ex.Message)
        Finally
            'oPO = Nothing
        End Try
    End Sub

End Class
