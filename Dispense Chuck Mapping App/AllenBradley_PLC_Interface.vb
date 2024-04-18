'App Developed by Joao Pires: JUL 2019
Imports HslCommunication.Profinet.AllenBradley
Imports HslCommunication
Module AllenBradley_PLC_Interface
    Public AllenBradleyPLC As AllenBradleyNet 'Initial access to the PLC object
    Public AB_PLC_IP As String
    Public Sub ConnectTo_AB_PLC()
        'PLC connection
        Try
            PLC_Connection = False
            AllenBradleyPLC = New AllenBradleyNet(AB_PLC_IP)
            AllenBradleyPLC.ConnectTimeOut = 5000
            Dim Connect As OperateResult = AllenBradleyPLC.ConnectServer()
            If Connect.IsSuccess Then
                PLC_Connection = True
            Else
                PLC_Connection = False
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ConnectTo_AB_PLC()")
            PLC_Connection = False
        End Try
    End Sub

    Public Sub DisconnectFrom_AB_PLC()
        'Disconnect the PLC
        Try
            AllenBradleyPLC.ConnectClose()
            PLC_Connection = False

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: DisconnectFrom_AB_PLC()")
            PLC_Connection = False
        End Try
    End Sub
    Public Sub ReadFromAB_Float(Tag_Name As String)

        Try
            PLC_DataReceived = ""
            PLC_DataReceived = AllenBradleyPLC.ReadFloat(Tag_Name).Content

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ReadFromAB_Float()")
        End Try
    End Sub
    Public Sub ReadFromAB_String(Tag_Name As String)

        Try
            PLC_DataReceived = ""
            PLC_DataReceived = AllenBradleyPLC.ReadString(Tag_Name).Content

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ReadFromAB_String()")
        End Try
    End Sub
End Module
