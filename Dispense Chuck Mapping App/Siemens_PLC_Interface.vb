'App Developed by Joao Pires: JUL 2019
Imports libnodave

Module Siemens_PLC_Interface
    Public fds As libnodave.daveOSserialType
    Public di As libnodave.daveInterface
    Public dc As libnodave.daveConnection
    Public res As Integer
    Public buf(100) As Byte
    Public localMPI As Integer
    Public rack As Integer
    Public slot As Integer
    Public plcMPI As Integer
    Public Siemens_PLC_IP As String

    Public Sub ConnectTo_Siemens_PLC()

        Try
            PLC_Connection = False
            fds.rfd = libnodave.openSocket(102, Siemens_PLC_IP)
            fds.wfd = fds.rfd

            If fds.rfd > 0 Then       ' if step 1 is ok
                di = New libnodave.daveInterface(fds, "IF1",
                    0, libnodave.daveProtoISOTCP,
                    libnodave.daveSpeed187k)

                di.setTimeout(10000)
                res = di.initAdapter
                If res = 0 Then       ' init Adapter is ok
                    ' rack and slot don't matter in case of MPI
                    dc = New libnodave.daveConnection(di, 0,
                                          rack, slot)
                    res = dc.connectPLC()
                    If res = 0 Then
                        PLC_Connection = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ConnectTo_Siemens_PLC()")
            PLC_Connection = False
        End Try
    End Sub

    Public Sub DisconnectFrom_Siemens_PLC()
        Try
            dc.disconnectPLC()
            PLC_Connection = False

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: DisconnectFrom_Siemens_PLC()")
        End Try
    End Sub

    Public Sub ReadFromPLC_DBReal(DBnumber As Integer, Address As Integer)
        Try
            PLC_DataReceived = ""
            res = dc.readBytes(libnodave.daveDB, DBnumber, Address, 4, buf)
            If res = 0 Then
                PLC_DataReceived = Str(dc.getFloat)
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ReadFromPLC_DBReal()")
        End Try
    End Sub

    Public Sub ReadFromPLC_DBString(DBnumber As Integer, FirstCharAddress As Integer, length As Integer)
        Try
            PLC_DataReceived = ""
            res = dc.readBytes(libnodave.daveDB, DBnumber, FirstCharAddress, length, buf)
            If res = 0 Then
                For i = 0 To length - 1
                    If buf(i) <> 0 And buf(i) <> 10 Then PLC_DataReceived &= Chr(buf(i)) Else PLC_DataReceived = ""
                Next
                If PLC_DataReceived.Length <> length Then PLC_DataReceived = ""
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ReadFromPLC_DBString()")
        End Try
    End Sub

End Module
