Public Class OBNK

    Private SBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private SBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim oInvoice As SAPbobsCOM.Documents

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        SBOApplication = oCatchingEvents.SBOApplication
        SBOCompany = oCatchingEvents.SBOCompany
    End Sub

    Public Function UpdateOBNK(ByVal Account As String, ByVal Fecha As String, ByVal DocTotal As Double, ByVal CardCode As String, ByVal DocEntry As Integer)

        Dim oCuenta As SAPbobsCOM.BankPages
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim BankAcct, Sequence As String
        oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try

            BankAcct = Account.Substring(0, 9)

            stQueryH = "Select T0.""Sequence"" from OBNK T0 where T0.""DueDate""='" & Fecha & "' and T0.""CredAmnt""=" & DocTotal & " and T0.""AcctCode""='" & BankAcct & "'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()
                Sequence = oRecSetH.Fields.Item("Sequence").Value

                oCuenta = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBankPages)

                oCuenta.GetByKey(BankAcct, Sequence)
                oCuenta.CardCode = CardCode
                oCuenta.ExternalCode = DocEntry
                oCuenta.Update()

            End If


        Catch ex As Exception

            SBOApplication.MessageBox("Error UpdateOBNK: " & ex.Message)

        End Try

    End Function

End Class
