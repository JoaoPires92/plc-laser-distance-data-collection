'App Developed by Joao Pires: JUL 2019

Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.IO
Module M_Main
    Public Structure PLC_Lane1Data

        Public LotCode As String

        'Siemens
        Public zPositionATCylinder As Double
        Public rHeightCalibration As Double
        Public DistanceNozzleProduct As Double
        Public DistanceNozzleDrop As Double
        Public raMeasuredHeight() As Double
        Public rActualHeight As Double
        Public raCorrectionValueCameraFocus() As Double

        'Allen Bradley
        Public NozzleTipHeight As Double
        Public LaserMeasurementToDatum As Double
        Public NozzleTipDrop As Double
        Public LaserMeasurementToLFID() As Double
        Public NozzleCalibToDatum_AxisPos As Double
        Public LaserLiveMeasurement As Double

    End Structure

    Public Structure PLC_Lane1Address

        'Siemens
        Public zPositionATCylinder As Integer
        Public rHeightCalibration As Integer
        Public DistanceNozzleProduct As Integer
        Public DistanceNozzleDrop As Integer
        Public raMeasuredHeightFirst As Integer
        Public raMeasuredHeightLast As Integer
        Public rActualHeight As Integer
        Public raCorrectionValueCameraFocusFirst As Integer
        Public raCorrectionValueCameraFocusLast As Integer
        Public LotCodeFirstChar As Integer

        'AllenBradley
        Public TagName_NozzleTipHeight As String
        Public TagName_LaserMeasurementToDatum As String
        Public TagName_NozzleTipDrop As String
        Public TagName_LaserMeasurementToLFID As String
        Public TagName_NozzleCalibToDatum_AxisPos As String
        Public TagName_LotCode As String
        Public TagName_LaserLiveMeasurement As String

    End Structure

    Public Structure PLC_Lane2Data

        Public LotCode As String

        'Siemens
        Public zPositionATCylinder As Double
        Public rHeightCalibration As Double
        Public DistanceNozzleProduct As Double
        Public DistanceNozzleDrop As Double
        Public raMeasuredHeight() As Double
        Public rActualHeight As Double
        Public raCorrectionValueCameraFocus() As Double

        'Allen Bradley
        Public NozzleTipHeight As Double
        Public LaserMeasurementToDatum As Double
        Public NozzleTipDrop As Double
        Public LaserMeasurementToLFID() As Double
        Public NozzleCalibToDatum_AxisPos As Double
        Public LaserLiveMeasurement As Double

    End Structure

    Public Structure PLC_Lane2Address

        'Siemens
        Public zPositionATCylinder As Integer
        Public rHeightCalibration As Integer
        Public DistanceNozzleProduct As Integer
        Public DistanceNozzleDrop As Integer
        Public raMeasuredHeightFirst As Integer
        Public raMeasuredHeightLast As Integer
        Public rActualHeight As Integer
        Public raCorrectionValueCameraFocusFirst As Integer
        Public raCorrectionValueCameraFocusLast As Integer
        Public LotCodeFirstChar As Integer

        'AllenBradley
        Public TagName_NozzleTipHeight As String
        Public TagName_LaserMeasurementToDatum As String
        Public TagName_NozzleTipDrop As String
        Public TagName_LaserMeasurementToLFID As String
        Public TagName_NozzleCalibToDatum_AxisPos As String
        Public TagName_LotCode As String
        Public TagName_LaserLiveMeasurement As String

    End Structure

    Public Lane1Data As PLC_Lane1Data
    Public Lane2Data As PLC_Lane2Data
    Public Lane1Address As PLC_Lane1Address
    Public Lane2Address As PLC_Lane2Address
    Public Lane1SensorNumber As Short
    Public Lane2SensorNumber As Short
    Public PLC_DataReceived As String
    Public strLane1Data As String
    Public strLane2Data As String
    Public MachineSelected As String
    Public PLC_ConfigurationLoadedOK As Boolean
    Public L1_CalibPin_HeightDifference As Double
    Public L1_ChuckRunOut As Double
    Public L2_CalibPin_HeightDifference As Double
    Public L2_ChuckRunOut As Double
    Public Lane1LDSAverage As Double
    Public Lane2LDSAverage As Double
    Public PLC_CommunicationLoadedOK As Boolean
    Public boSiemensPLC As Boolean
    Public boAllenBradleyPLC As Boolean
    Public PLC_Connection As Boolean
    Public boSimulatorMode As Boolean
    Public MachineNameLoadedOk As Boolean

    Public Sub GetLane1Data()

        Try
            If boSiemensPLC = True Then

                Lane1SensorNumber = 1

                ReadFromPLC_DBReal(1400, Lane1Address.zPositionATCylinder)
                If PLC_DataReceived <> "" Then Lane1Data.zPositionATCylinder = PLC_DataReceived Else Lane1Data.zPositionATCylinder = "#.###"

                ReadFromPLC_DBReal(1400, Lane1Address.rHeightCalibration)
                If PLC_DataReceived <> "" Then Lane1Data.rHeightCalibration = PLC_DataReceived Else Lane1Data.rHeightCalibration = "#.###"

                ReadFromPLC_DBReal(1008, Lane1Address.DistanceNozzleProduct)
                If PLC_DataReceived <> "" Then Lane1Data.DistanceNozzleProduct = PLC_DataReceived Else Lane1Data.DistanceNozzleProduct = "#.###"

                ReadFromPLC_DBReal(11, Lane1Address.DistanceNozzleDrop)
                If PLC_DataReceived <> "" Then Lane1Data.DistanceNozzleDrop = PLC_DataReceived Else Lane1Data.DistanceNozzleDrop = "#.###"

                For Address = Lane1Address.raMeasuredHeightFirst To Lane1Address.raMeasuredHeightLast Step 4
                    ReadFromPLC_DBReal(1400, Address)
                    If PLC_DataReceived <> "" Then Lane1Data.raMeasuredHeight(Lane1SensorNumber) = PLC_DataReceived Else Lane1Data.raMeasuredHeight(Lane1SensorNumber) = "#.###"
                    Lane1SensorNumber = Lane1SensorNumber + 1
                Next

                Lane1SensorNumber = 1

                For Address = Lane1Address.raCorrectionValueCameraFocusFirst To Lane1Address.raCorrectionValueCameraFocusLast Step 4
                    ReadFromPLC_DBReal(1400, Address)
                    If PLC_DataReceived <> "" Then Lane1Data.raCorrectionValueCameraFocus(Lane1SensorNumber) = PLC_DataReceived Else Lane1Data.raCorrectionValueCameraFocus(Lane1SensorNumber) = "#.###"
                    Lane1SensorNumber = Lane1SensorNumber + 1
                Next

            End If

            If boAllenBradleyPLC = True Then

                Lane1SensorNumber = 1

                ReadFromAB_Float(Lane1Address.TagName_NozzleTipHeight)
                If PLC_DataReceived <> "" Then Lane1Data.NozzleTipHeight = PLC_DataReceived Else Lane1Data.NozzleTipHeight = "#.###"

                ReadFromAB_Float(Lane1Address.TagName_LaserMeasurementToDatum)
                If PLC_DataReceived <> "" Then Lane1Data.LaserMeasurementToDatum = PLC_DataReceived Else Lane1Data.LaserMeasurementToDatum = "#.###"

                ReadFromAB_Float(Lane1Address.TagName_NozzleTipDrop)
                If PLC_DataReceived <> "" Then Lane1Data.NozzleTipDrop = PLC_DataReceived Else Lane1Data.NozzleTipDrop = "#.###"

                ReadFromAB_Float(Lane1Address.TagName_NozzleCalibToDatum_AxisPos)
                If PLC_DataReceived <> "" Then Lane1Data.NozzleCalibToDatum_AxisPos = PLC_DataReceived Else Lane1Data.NozzleCalibToDatum_AxisPos = "#.###"

                For i = 0 To 24
                    ReadFromAB_Float(Lane1Address.TagName_LaserMeasurementToLFID & "[" & i & "]")
                    If PLC_DataReceived <> "" Then Lane1Data.LaserMeasurementToLFID(Lane1SensorNumber) = PLC_DataReceived Else Lane1Data.LaserMeasurementToLFID(Lane1SensorNumber) = "#.###"
                    Lane1SensorNumber = Lane1SensorNumber + 1
                Next

            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: GetLane1Data()")
        End Try

    End Sub

    Public Sub ReadActualHeightLane1()

        Try
            If boSiemensPLC = True Then
                ReadFromPLC_DBReal(1400, Lane1Address.rActualHeight)
                If PLC_DataReceived <> "" Then Lane1Data.rActualHeight = PLC_DataReceived Else Lane1Data.rActualHeight = "#.###"
            End If

            If boAllenBradleyPLC = True Then
                ReadFromAB_Float(Lane1Address.TagName_LaserLiveMeasurement)
                If PLC_DataReceived <> "" Then Lane1Data.LaserLiveMeasurement = PLC_DataReceived Else Lane1Data.LaserLiveMeasurement = "#.###"
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ReadActualHeightLane1()")
        End Try

    End Sub

    Public Sub GetLotCodeLane1()
        Try
            If boSiemensPLC = True Then
                ReadFromPLC_DBString(11, Lane1Address.LotCodeFirstChar, 10)
                If PLC_DataReceived = "" Then
                    Lane1Data.LotCode = "No Lot Online"
                Else
                    Lane1Data.LotCode = PLC_DataReceived
                End If
            End If

            If boAllenBradleyPLC = True Then
                ReadFromAB_String(Lane1Address.TagName_LotCode)
                If PLC_DataReceived = "" Then
                    Lane1Data.LotCode = "No Lot Online"
                Else
                    Lane1Data.LotCode = PLC_DataReceived
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: GetLotCodeLane1()")
        End Try
    End Sub

    Public Sub ConvertLane1Data_ToString()

        Try
            If boSiemensPLC = True Then

                strLane1Data &= vbTab & vbTab & "zPositionATCylinder:" & vbTab & vbTab & vbTab & Format(Lane1Data.zPositionATCylinder, "0.000") & vbNewLine &
                vbTab & vbTab & "rHeightCalibration:" & vbTab & vbTab & vbTab & vbTab & Format(Lane1Data.rHeightCalibration, "0.000") & vbNewLine &
                vbTab & vbTab & "DistanceNozzleProduct:" & vbTab & vbTab & vbTab & Format(Lane1Data.DistanceNozzleProduct, "0.000") & vbNewLine &
                vbTab & vbTab & "DistanceNozzleDrop:" & vbTab & vbTab & vbTab & Format(Lane1Data.DistanceNozzleDrop, "0.000") & vbNewLine & vbNewLine

                For i = 1 To 25
                    strLane1Data &= vbTab & vbTab & "raMeasuredHeight" & "[" & i & "]" & vbTab & vbTab & vbTab & Format(Lane1Data.raMeasuredHeight(i), "0.000") & vbNewLine
                Next
                strLane1Data &= vbNewLine
                For i = 1 To 25
                    strLane1Data &= vbTab & "raCorrectionCameraFocus" & "[" & i & "]" & vbTab & vbTab & vbTab & Format(Lane1Data.raCorrectionValueCameraFocus(i), "0.000") & vbNewLine
                Next

            End If

            If boAllenBradleyPLC = True Then

                strLane1Data &= vbTab & "NozzleTipHeight:" & vbTab & vbTab & vbTab & vbTab & vbTab & Format(Lane1Data.NozzleTipHeight, "0.000") & vbNewLine &
                 vbTab & "NozzleTipToDrop:" & vbTab & vbTab & vbTab & vbTab & vbTab & Format(Lane1Data.NozzleTipDrop, "0.000") & vbNewLine &
                 vbTab & "LaserMeasurementToDatum:" & vbTab & vbTab & vbTab & Format(Lane1Data.LaserMeasurementToDatum, "0.000") & vbNewLine &
                 vbTab & "NozzleCalibToDatum_ZAxisPos:" & vbTab & vbTab & Format(Lane1Data.NozzleCalibToDatum_AxisPos, "0.000") & vbNewLine & vbNewLine

                For i = 1 To 25
                    strLane1Data &= vbTab & "LaserMeasurementToLFID" & "[" & i - 1 & "]" & vbTab & vbTab & vbTab & Format(Lane1Data.LaserMeasurementToLFID(i), "0.000") & vbNewLine
                Next

            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ConvertLane1Data_ToString()")
        End Try

    End Sub

    Public Sub GetLane2Data()

        Try
            If boSiemensPLC = True Then

                Lane2SensorNumber = 1

                ReadFromPLC_DBReal(1600, Lane2Address.zPositionATCylinder)
                If PLC_DataReceived <> "" Then Lane2Data.zPositionATCylinder = PLC_DataReceived Else Lane2Data.zPositionATCylinder = "#.###"

                ReadFromPLC_DBReal(1600, Lane2Address.rHeightCalibration)
                If PLC_DataReceived <> "" Then Lane2Data.rHeightCalibration = PLC_DataReceived Else Lane2Data.rHeightCalibration = "#.###"

                ReadFromPLC_DBReal(1008, Lane2Address.DistanceNozzleProduct)
                If PLC_DataReceived <> "" Then Lane2Data.DistanceNozzleProduct = PLC_DataReceived Else Lane2Data.DistanceNozzleProduct = "#.###"

                ReadFromPLC_DBReal(11, Lane2Address.DistanceNozzleDrop)
                If PLC_DataReceived <> "" Then Lane2Data.DistanceNozzleDrop = PLC_DataReceived Else Lane2Data.DistanceNozzleDrop = "#.###"

                For Address = Lane2Address.raMeasuredHeightFirst To Lane2Address.raMeasuredHeightLast Step 4
                    ReadFromPLC_DBReal(1600, Address)
                    If PLC_DataReceived <> "" Then Lane2Data.raMeasuredHeight(Lane2SensorNumber) = PLC_DataReceived Else Lane2Data.raMeasuredHeight(Lane2SensorNumber) = "#.###"
                    Lane2SensorNumber = Lane2SensorNumber + 1
                Next

                Lane2SensorNumber = 1

                For Address = Lane2Address.raCorrectionValueCameraFocusFirst To Lane2Address.raCorrectionValueCameraFocusLast Step 4
                    ReadFromPLC_DBReal(1600, Address)
                    If PLC_DataReceived <> "" Then Lane2Data.raCorrectionValueCameraFocus(Lane2SensorNumber) = PLC_DataReceived Else Lane2Data.raCorrectionValueCameraFocus(Lane2SensorNumber) = "#.###"
                    Lane2SensorNumber = Lane2SensorNumber + 1
                Next

            End If

            If boAllenBradleyPLC = True Then

                Lane2SensorNumber = 1

                ReadFromAB_Float(Lane2Address.TagName_NozzleTipHeight)
                If PLC_DataReceived <> "" Then Lane2Data.NozzleTipHeight = PLC_DataReceived Else Lane2Data.NozzleTipHeight = "#.###"

                ReadFromAB_Float(Lane2Address.TagName_LaserMeasurementToDatum)
                If PLC_DataReceived <> "" Then Lane2Data.LaserMeasurementToDatum = PLC_DataReceived Else Lane2Data.LaserMeasurementToDatum = "#.###"

                ReadFromAB_Float(Lane2Address.TagName_NozzleTipDrop)
                If PLC_DataReceived <> "" Then Lane2Data.NozzleTipDrop = PLC_DataReceived Else Lane2Data.NozzleTipDrop = "#.###"

                ReadFromAB_Float(Lane2Address.TagName_NozzleCalibToDatum_AxisPos)
                If PLC_DataReceived <> "" Then Lane2Data.NozzleCalibToDatum_AxisPos = PLC_DataReceived Else Lane2Data.NozzleCalibToDatum_AxisPos = "#.###"

                For i = 0 To 24
                    ReadFromAB_Float(Lane2Address.TagName_LaserMeasurementToLFID & "[" & i & "]")
                    If PLC_DataReceived <> "" Then Lane2Data.LaserMeasurementToLFID(Lane2SensorNumber) = PLC_DataReceived Else Lane2Data.LaserMeasurementToLFID(Lane2SensorNumber) = "#.###"
                    Lane2SensorNumber = Lane2SensorNumber + 1
                Next

            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: GetLane2Data()")
        End Try

    End Sub

    Public Sub ReadActualHeightLane2()

        Try
            If boSiemensPLC = True Then
                ReadFromPLC_DBReal(1600, Lane2Address.rActualHeight)
                If PLC_DataReceived <> "" Then Lane2Data.rActualHeight = PLC_DataReceived Else Lane2Data.rActualHeight = "#.###"
            End If

            If boAllenBradleyPLC = True Then
                ReadFromAB_Float(Lane2Address.TagName_LaserLiveMeasurement)
                If PLC_DataReceived <> "" Then Lane2Data.LaserLiveMeasurement = PLC_DataReceived Else Lane2Data.LaserLiveMeasurement = "#.###"
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ReadActualHeightLane2()")
        End Try

    End Sub

    Public Sub GetLotCodeLane2()
        Try
            If boSiemensPLC = True Then
                ReadFromPLC_DBString(11, Lane2Address.LotCodeFirstChar, 10)
                If PLC_DataReceived = "" Then
                    Lane2Data.LotCode = "No Lot Online"
                Else
                    Lane2Data.LotCode = PLC_DataReceived
                End If
            End If

            If boAllenBradleyPLC = True Then
                ReadFromAB_String(Lane2Address.TagName_LotCode)
                If PLC_DataReceived = "" Then
                    Lane2Data.LotCode = "No Lot Online"
                Else
                    Lane2Data.LotCode = PLC_DataReceived
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: GetLotCodeLane2()")
        End Try
    End Sub

    Public Sub ConvertLane2Data_ToString()

        Try
            If boSiemensPLC = True Then
                strLane2Data &= vbTab & vbTab & "zPositionATCylinder:" & vbTab & vbTab & vbTab & Format(Lane2Data.zPositionATCylinder, "0.000") & vbNewLine &
                vbTab & vbTab & "rHeightCalibration:" & vbTab & vbTab & vbTab & vbTab & Format(Lane2Data.rHeightCalibration, "0.000") & vbNewLine &
                vbTab & vbTab & "DistanceNozzleProduct:" & vbTab & vbTab & vbTab & Format(Lane2Data.DistanceNozzleProduct, "0.000") & vbNewLine &
                vbTab & vbTab & "DistanceNozzleDrop:" & vbTab & vbTab & vbTab & Format(Lane2Data.DistanceNozzleDrop, "0.000") & vbNewLine & vbNewLine

                For i = 1 To 25
                    strLane2Data &= vbTab & vbTab & "raMeasuredHeight" & "[" & i & "]" & vbTab & vbTab & vbTab & Format(Lane2Data.raMeasuredHeight(i), "0.000") & vbNewLine
                Next
                strLane2Data &= vbNewLine
                For i = 1 To 25
                    strLane2Data &= vbTab & "raCorrectionCameraFocus" & "[" & i & "]" & vbTab & vbTab & vbTab & Format(Lane2Data.raCorrectionValueCameraFocus(i), "0.000") & vbNewLine
                Next
            End If

            If boAllenBradleyPLC = True Then

                strLane2Data &= vbTab & "NozzleTipHeight:" & vbTab & vbTab & vbTab & vbTab & vbTab & Format(Lane2Data.NozzleTipHeight, "0.000") & vbNewLine &
                 vbTab & "NozzleTipToDrop:" & vbTab & vbTab & vbTab & vbTab & vbTab & Format(Lane2Data.NozzleTipDrop, "0.000") & vbNewLine &
                 vbTab & "LaserMeasurementToDatum:" & vbTab & vbTab & vbTab & Format(Lane2Data.LaserMeasurementToDatum, "0.000") & vbNewLine &
                 vbTab & "NozzleCalibToDatum_ZAxisPos:" & vbTab & vbTab & Format(Lane2Data.NozzleCalibToDatum_AxisPos, "0.000") & vbNewLine & vbNewLine

                For i = 1 To 25
                    strLane2Data &= vbTab & "LaserMeasurementToLFID" & "[" & i - 1 & "]" & vbTab & vbTab & vbTab & Format(Lane2Data.LaserMeasurementToLFID(i), "0.000") & vbNewLine
                Next

            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: ConvertLane2Data_ToString()")
        End Try

    End Sub

    Public Sub CalculateChuckRunOut()

        Try
            Dim Lane1_Max As Double
            Dim Lane1_Min As Double
            Dim Lane2_Max As Double
            Dim Lane2_Min As Double
            Dim A_Lane1(24) As Double
            Dim A_Lane2(24) As Double

            If boSiemensPLC = True Then
                For i = 0 To 24
                    A_Lane1(i) = Lane1Data.raMeasuredHeight(i + 1)
                    A_Lane2(i) = Lane2Data.raMeasuredHeight(i + 1)
                Next

                Lane1_Max = A_Lane1.Max()
                Lane1_Min = A_Lane1.Min()
                Lane2_Max = A_Lane2.Max()
                Lane2_Min = A_Lane2.Min()

                L1_ChuckRunOut = Lane1_Max - Lane1_Min
                L2_ChuckRunOut = Lane2_Max - Lane2_Min
            End If

            If boAllenBradleyPLC = True Then
                For i = 0 To 24
                    A_Lane1(i) = Lane1Data.LaserMeasurementToLFID(i + 1)
                    A_Lane2(i) = Lane2Data.LaserMeasurementToLFID(i + 1)
                Next

                Lane1_Max = A_Lane1.Max()
                Lane1_Min = A_Lane1.Min()
                Lane2_Max = A_Lane2.Max()
                Lane2_Min = A_Lane2.Min()

                L1_ChuckRunOut = Lane1_Max - Lane1_Min
                L2_ChuckRunOut = Lane2_Max - Lane2_Min
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: CalculateChuckRunOut()")
        End Try

    End Sub

    Public Sub CalculateCalibPinHeightDifferenceToChuck()

        Try

            If boSiemensPLC = True Then
                For i = 1 To 25
                    Lane1LDSAverage = Lane1LDSAverage + Lane1Data.raMeasuredHeight(i) / 25
                Next

                For i = 1 To 25
                    Lane2LDSAverage = Lane2LDSAverage + Lane2Data.raMeasuredHeight(i) / 25
                Next

                L1_CalibPin_HeightDifference = Lane1Data.rHeightCalibration - Lane1LDSAverage
                L2_CalibPin_HeightDifference = Lane2Data.rHeightCalibration - Lane2LDSAverage
            End If

            If boAllenBradleyPLC = True Then
                For i = 1 To 25
                    Lane1LDSAverage = Lane1LDSAverage + Lane1Data.LaserMeasurementToLFID(i) / 25
                Next

                For i = 1 To 25
                    Lane2LDSAverage = Lane2LDSAverage + Lane2Data.LaserMeasurementToLFID(i) / 25
                Next

                L1_CalibPin_HeightDifference = Lane1Data.LaserMeasurementToDatum - Lane1LDSAverage
                L2_CalibPin_HeightDifference = Lane2Data.LaserMeasurementToDatum - Lane2LDSAverage
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: CalculateCalibPinDifference()")
        End Try

    End Sub

    Public Function SaveDataIntoExcelSheet(ByVal FileName As String) As Boolean

        Dim objExcel As Excel.Application, objWorkbook As Excel.Workbook
        Dim objSheet As Excel.Worksheet
        Dim strFileName As String
        Dim Counter As Short
        Dim MyPassword As String

        Try

            MyPassword = "Admin_100" ' Excel Edit Sheet Password

            If boSiemensPLC = True Then

                objExcel = CreateObject("Excel.Application")
                objExcel.Visible = True
                strFileName = CurDir() & "\Resources\GEG_Dispense_Template.xlsx"
                objWorkbook = objExcel.Workbooks.Add(strFileName)
                objSheet = objWorkbook.Sheets("Lane 1")
                objSheet.Cells(2, 3) = L1_CalibPin_HeightDifference
                objSheet.Cells(3, 3) = L1_ChuckRunOut
                objSheet.Cells(5, 3) = Lane1Data.zPositionATCylinder
                objSheet.Cells(6, 3) = Lane1Data.rHeightCalibration
                objSheet.Cells(7, 3) = Lane1Data.DistanceNozzleProduct
                objSheet.Cells(8, 3) = Lane1Data.DistanceNozzleDrop
                objSheet.Cells(1, 4) = MachineSelected
                If Lane1Data.LotCode = "Lane 1: No Lot Online" Then objSheet.Cells(1, 5) = "No Lot Online " Else objSheet.Cells(1, 5) = "Lot: " & Lane1Data.LotCode
                objSheet.Cells(1, 7) = DateTime.Now

                Counter = 10

                For i = 1 To 25
                    objSheet.Cells(Counter, 3) = Lane1Data.raMeasuredHeight(i)
                    Counter += 1
                Next

                Counter = 36

                For i = 1 To 25
                    objSheet.Cells(Counter, 3) = Lane1Data.raCorrectionValueCameraFocus(i)
                    Counter += 1
                Next

                objSheet.Protect(MyPassword)

                objSheet = objWorkbook.Sheets("Lane 2")
                objSheet.Cells(2, 3) = L2_CalibPin_HeightDifference
                objSheet.Cells(3, 3) = L2_ChuckRunOut
                objSheet.Cells(5, 3) = Lane2Data.zPositionATCylinder
                objSheet.Cells(6, 3) = Lane2Data.rHeightCalibration
                objSheet.Cells(7, 3) = Lane2Data.DistanceNozzleProduct
                objSheet.Cells(8, 3) = Lane2Data.DistanceNozzleDrop
                objSheet.Cells(1, 4) = MachineSelected
                If Lane2Data.LotCode = "Lane 2: No Lot Online" Then objSheet.Cells(1, 5) = "No Lot Online " Else objSheet.Cells(1, 5) = "Lot: " & Lane2Data.LotCode
                objSheet.Cells(1, 7) = DateTime.Now

                Counter = 10

                For i = 1 To 25
                    objSheet.Cells(Counter, 3) = Lane2Data.raMeasuredHeight(i)
                    Counter += 1
                Next

                Counter = 36

                For i = 1 To 25
                    objSheet.Cells(Counter, 3) = Lane2Data.raCorrectionValueCameraFocus(i)
                    Counter += 1
                Next

                objSheet.Protect(MyPassword)

                objWorkbook.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookDefault)
                objWorkbook.Close(False)
                objExcel.Quit()

                Return True
            End If

            If boAllenBradleyPLC = True Then

                objExcel = CreateObject("Excel.Application")
                objExcel.Visible = True
                strFileName = CurDir() & "\Resources\ATS_Dispense_Template.xlsx"
                objWorkbook = objExcel.Workbooks.Add(strFileName)

                objSheet = objWorkbook.Sheets("Lane 1")
                objSheet.Cells(2, 3) = L1_CalibPin_HeightDifference
                objSheet.Cells(3, 3) = L1_ChuckRunOut
                objSheet.Cells(5, 3) = Lane1Data.NozzleTipHeight
                objSheet.Cells(6, 3) = Lane1Data.NozzleTipDrop
                objSheet.Cells(7, 3) = Lane1Data.LaserMeasurementToDatum
                objSheet.Cells(8, 3) = Lane1Data.NozzleCalibToDatum_AxisPos
                objSheet.Cells(1, 4) = MachineSelected
                If Lane1Data.LotCode = "Lane 1: No Lot Online" Then objSheet.Cells(1, 5) = "No Lot Online " Else objSheet.Cells(1, 5) = "Lot: " & Lane1Data.LotCode
                objSheet.Cells(1, 7) = DateTime.Now

                Counter = 10

                For i = 1 To 25
                    objSheet.Cells(Counter, 3) = Lane1Data.LaserMeasurementToLFID(i)
                    Counter += 1
                Next

                objSheet.Protect(MyPassword)

                objSheet = objWorkbook.Sheets("Lane 2")
                objSheet.Cells(2, 3) = L2_CalibPin_HeightDifference
                objSheet.Cells(3, 3) = L2_ChuckRunOut
                objSheet.Cells(5, 3) = Lane2Data.NozzleTipHeight
                objSheet.Cells(6, 3) = Lane2Data.NozzleTipDrop
                objSheet.Cells(7, 3) = Lane2Data.LaserMeasurementToDatum
                objSheet.Cells(8, 3) = Lane2Data.NozzleCalibToDatum_AxisPos
                objSheet.Cells(1, 4) = MachineSelected
                If Lane2Data.LotCode = "Lane 2: No Lot Online" Then objSheet.Cells(1, 5) = "No Lot Online " Else objSheet.Cells(1, 5) = "Lot: " & Lane2Data.LotCode
                objSheet.Cells(1, 7) = DateTime.Now

                Counter = 10

                For i = 1 To 25
                    objSheet.Cells(Counter, 3) = Lane2Data.LaserMeasurementToLFID(i)
                    Counter += 1
                Next

                objSheet.Protect(MyPassword)

                objWorkbook.SaveAs(FileName, Excel.XlFileFormat.xlWorkbookDefault)
                objWorkbook.Close(False)
                objExcel.Quit()

                Return True
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: SaveDataIntoExcelSheet() ")
            Return False
        End Try

    End Function
    Public Sub Load_Siemens_PLC_Configuration()

        Dim AuxStr As String
        Dim FilePointer As FileStream
        Dim A_AuxStr() As String

        Try

            FilePointer = New FileStream("Resources\Siemens_PLC_Configuration.Dat", FileMode.Open, FileAccess.Read)

            Dim StreamInput As New StreamReader(FilePointer)

            StreamInput.BaseStream.Seek(0, SeekOrigin.Begin)

            AuxStr = ""

            Do Until AuxStr = "End of Configuration File"

                AuxStr = StreamInput.ReadLine()

                If Len(AuxStr) > 25 Then

                    A_AuxStr = AuxStr.Split("|")

                    If A_AuxStr(0) = MachineSelected Then
                        If A_AuxStr(1) = "Lane1" Then
                            Lane1Address.zPositionATCylinder = A_AuxStr(2)
                            Lane1Address.rHeightCalibration = A_AuxStr(3)
                            Lane1Address.DistanceNozzleProduct = A_AuxStr(4)
                            Lane1Address.DistanceNozzleDrop = A_AuxStr(5)
                            Lane1Address.raMeasuredHeightFirst = A_AuxStr(6)
                            Lane1Address.raMeasuredHeightLast = A_AuxStr(7)
                            Lane1Address.rActualHeight = A_AuxStr(8)
                            Lane1Address.raCorrectionValueCameraFocusFirst = A_AuxStr(9)
                            Lane1Address.raCorrectionValueCameraFocusLast = A_AuxStr(10)
                            Lane1Address.LotCodeFirstChar = A_AuxStr(11)
                        End If
                        If A_AuxStr(12) = "Lane2" Then
                            Lane2Address.zPositionATCylinder = A_AuxStr(13)
                            Lane2Address.rHeightCalibration = A_AuxStr(14)
                            Lane2Address.DistanceNozzleProduct = A_AuxStr(15)
                            Lane2Address.DistanceNozzleDrop = A_AuxStr(16)
                            Lane2Address.raMeasuredHeightFirst = A_AuxStr(17)
                            Lane2Address.raMeasuredHeightLast = A_AuxStr(18)
                            Lane2Address.rActualHeight = A_AuxStr(19)
                            Lane2Address.raCorrectionValueCameraFocusFirst = A_AuxStr(20)
                            Lane2Address.raCorrectionValueCameraFocusLast = A_AuxStr(21)
                            Lane2Address.LotCodeFirstChar = A_AuxStr(22)
                        End If
                    End If
                Else
                    If AuxStr = "End of Configuration File" Then
                        PLC_ConfigurationLoadedOK = True
                    Else
                        MsgBox("Load PLC configuration file error.", MsgBoxStyle.Critical, "Load_PlcConfiguration()")
                        PLC_ConfigurationLoadedOK = False
                        Exit Sub
                    End If

                End If
            Loop

            StreamInput.Close()
            FilePointer.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: Load_PLCConfiguration()")
            PLC_ConfigurationLoadedOK = False
        End Try

    End Sub

    Public Sub Load_AB_PLC_Configuration()

        Dim AuxStr As String
        Dim FilePointer As FileStream
        Dim A_AuxStr() As String

        Try

            FilePointer = New FileStream("Resources\AB_PLC_Configuration.Dat", FileMode.Open, FileAccess.Read)

            Dim StreamInput As New StreamReader(FilePointer)

            StreamInput.BaseStream.Seek(0, SeekOrigin.Begin)

            AuxStr = ""

            Do Until AuxStr = "End of Configuration File"

                AuxStr = StreamInput.ReadLine()

                If Len(AuxStr) > 25 Then

                    A_AuxStr = AuxStr.Split("|")

                    If A_AuxStr(0) = MachineSelected Then
                        If A_AuxStr(1) = "Lane1" Then
                            Lane1Address.TagName_NozzleTipHeight = A_AuxStr(2)
                            Lane1Address.TagName_LaserMeasurementToDatum = A_AuxStr(3)
                            Lane1Address.TagName_NozzleTipDrop = A_AuxStr(4)
                            Lane1Address.TagName_LaserMeasurementToLFID = A_AuxStr(5)
                            Lane1Address.TagName_NozzleCalibToDatum_AxisPos = A_AuxStr(6)
                            Lane1Address.TagName_LotCode = A_AuxStr(7)
                            Lane1Address.TagName_LaserLiveMeasurement = A_AuxStr(8)
                        End If
                        If A_AuxStr(9) = "Lane2" Then
                            Lane2Address.TagName_NozzleTipHeight = A_AuxStr(10)
                            Lane2Address.TagName_LaserMeasurementToDatum = A_AuxStr(11)
                            Lane2Address.TagName_NozzleTipDrop = A_AuxStr(12)
                            Lane2Address.TagName_LaserMeasurementToLFID = A_AuxStr(13)
                            Lane2Address.TagName_NozzleCalibToDatum_AxisPos = A_AuxStr(14)
                            Lane2Address.TagName_LotCode = A_AuxStr(15)
                            Lane2Address.TagName_LaserLiveMeasurement = A_AuxStr(16)
                        End If
                    End If
                Else
                    If AuxStr = "End of Configuration File" Then
                        PLC_ConfigurationLoadedOK = True
                    Else
                        MsgBox("Load PLC configuration file error.", MsgBoxStyle.Critical, "Load_AB_PLC_Configuration()")
                        PLC_ConfigurationLoadedOK = False
                        Exit Sub
                    End If

                End If
            Loop

            StreamInput.Close()
            FilePointer.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: Load_PLCConfiguration()")
            PLC_ConfigurationLoadedOK = False
        End Try

    End Sub

    Public Sub Load_Siemens_PLC_Communication()

        Dim AuxStr As String
        Dim FilePointer As FileStream
        Dim A_AuxStr() As String

        Try
            PLC_CommunicationLoadedOK = False

            FilePointer = New FileStream("Resources\Siemens_PLC_Communication.Dat", FileMode.Open, FileAccess.Read)

            Dim StreamInput As New StreamReader(FilePointer)

            StreamInput.BaseStream.Seek(0, SeekOrigin.Begin)

            AuxStr = ""

            Do Until AuxStr = "End of Configuration File"

                AuxStr = StreamInput.ReadLine()

                If Len(AuxStr) > 50 Then
                    A_AuxStr = AuxStr.Split("|")
                    If A_AuxStr(0) = "localMPI" Then localMPI = A_AuxStr(1)
                    If A_AuxStr(2) = "rack" Then rack = A_AuxStr(3)
                    If A_AuxStr(4) = "slot" Then slot = A_AuxStr(5)
                    If A_AuxStr(6) = "plcMPI" Then plcMPI = A_AuxStr(7)
                    If A_AuxStr(8) = "plcIP" Then Siemens_PLC_IP = A_AuxStr(9)
                Else
                    If AuxStr = "End of Configuration File" Then
                        PLC_CommunicationLoadedOK = True
                    Else
                        MsgBox("Load PLC communication file error.", MsgBoxStyle.Critical, "Load_Siemens_PLC_Communication()")
                        PLC_CommunicationLoadedOK = False
                        Exit Sub
                    End If

                End If
            Loop

            StreamInput.Close()
            FilePointer.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: Load_PlcCommunication()")
            PLC_CommunicationLoadedOK = False
        End Try

    End Sub

    Public Sub Load_AB_PLC_Communication()

        Dim AuxStr As String
        Dim FilePointer As FileStream
        Dim A_AuxStr() As String

        Try
            PLC_CommunicationLoadedOK = False

            FilePointer = New FileStream("Resources\AB_PLC_Communication.Dat", FileMode.Open, FileAccess.Read)

            Dim StreamInput As New StreamReader(FilePointer)

            StreamInput.BaseStream.Seek(0, SeekOrigin.Begin)

            AuxStr = ""

            Do Until AuxStr = "End of Configuration File"

                AuxStr = StreamInput.ReadLine()

                If Len(AuxStr) = 18 Then
                    A_AuxStr = AuxStr.Split("|")
                    If A_AuxStr(0) = "PLC_IP" Then AB_PLC_IP = A_AuxStr(1)
                Else
                    If AuxStr = "End of Configuration File" Then
                        PLC_CommunicationLoadedOK = True
                    Else
                        MsgBox("Load PLC communication file error.", MsgBoxStyle.Critical, "Load_PlcCommunication()")
                        PLC_CommunicationLoadedOK = False
                        Exit Sub
                    End If

                End If
            Loop

            StreamInput.Close()
            FilePointer.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: Load_AB_PLC_Communication()")
            PLC_CommunicationLoadedOK = False
        End Try

    End Sub

    Public Sub Load_Machine_Name()

        Dim AuxStr As String
        Dim FilePointer As FileStream

        Try
            MachineNameLoadedOk = False

            FilePointer = New FileStream("Resources\Machine_Name.Dat", FileMode.Open, FileAccess.Read)

            Dim StreamInput As New StreamReader(FilePointer)

            StreamInput.BaseStream.Seek(0, SeekOrigin.Begin)

            AuxStr = ""

            Do Until AuxStr = "End of Configuration File"

                AuxStr = StreamInput.ReadLine()

                If AuxStr <> "End of Configuration File" Then
                    F_Main.ComboBox1.Items.Add(AuxStr)
                Else
                    If AuxStr = "End of Configuration File" Then
                        MachineNameLoadedOk = True
                    Else
                        MsgBox("Load Machine Name File Error", MsgBoxStyle.Critical, "Load_Machine_Name()")
                        MachineNameLoadedOk = False
                        Exit Sub
                    End If

                End If
            Loop

            StreamInput.Close()
            FilePointer.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: Load_AB_PLC_Communication()")
            PLC_CommunicationLoadedOK = False
        End Try

    End Sub

    Public Sub LoadSiemensSimulatorValues()

        Try
            'Lane 1
            Lane1Data.zPositionATCylinder = -4.541
            Lane1Data.rHeightCalibration = 1.357
            Lane1Data.DistanceNozzleProduct = 1.2
            Lane1Data.DistanceNozzleDrop = 1.144

            Lane1Data.raMeasuredHeight(1) = 1.469
            Lane1Data.raMeasuredHeight(2) = 1.483
            Lane1Data.raMeasuredHeight(3) = 1.479
            Lane1Data.raMeasuredHeight(4) = 1.477
            Lane1Data.raMeasuredHeight(5) = 1.47
            Lane1Data.raMeasuredHeight(6) = 1.488
            Lane1Data.raMeasuredHeight(7) = 1.482
            Lane1Data.raMeasuredHeight(8) = 1.483
            Lane1Data.raMeasuredHeight(9) = 1.48
            Lane1Data.raMeasuredHeight(10) = 1.475
            Lane1Data.raMeasuredHeight(11) = 1.491
            Lane1Data.raMeasuredHeight(12) = 1.486
            Lane1Data.raMeasuredHeight(13) = 1.478
            Lane1Data.raMeasuredHeight(14) = 1.479
            Lane1Data.raMeasuredHeight(15) = 1.479
            Lane1Data.raMeasuredHeight(16) = 1.471
            Lane1Data.raMeasuredHeight(17) = 1.469
            Lane1Data.raMeasuredHeight(18) = 1.474
            Lane1Data.raMeasuredHeight(19) = 1.461
            Lane1Data.raMeasuredHeight(20) = 1.472
            Lane1Data.raMeasuredHeight(21) = 1.465
            Lane1Data.raMeasuredHeight(22) = 1.461
            Lane1Data.raMeasuredHeight(23) = 1.468
            Lane1Data.raMeasuredHeight(24) = 1.449
            Lane1Data.raMeasuredHeight(25) = 1.444

            Lane1Data.raCorrectionValueCameraFocus(1) = -0.05
            Lane1Data.raCorrectionValueCameraFocus(2) = -0.05
            Lane1Data.raCorrectionValueCameraFocus(3) = -0.037
            Lane1Data.raCorrectionValueCameraFocus(4) = -0.04
            Lane1Data.raCorrectionValueCameraFocus(5) = -0.043
            Lane1Data.raCorrectionValueCameraFocus(6) = -0.049
            Lane1Data.raCorrectionValueCameraFocus(7) = -0.032
            Lane1Data.raCorrectionValueCameraFocus(8) = -0.038
            Lane1Data.raCorrectionValueCameraFocus(9) = -0.037
            Lane1Data.raCorrectionValueCameraFocus(10) = -0.04
            Lane1Data.raCorrectionValueCameraFocus(11) = -0.045
            Lane1Data.raCorrectionValueCameraFocus(12) = -0.028
            Lane1Data.raCorrectionValueCameraFocus(13) = -0.034
            Lane1Data.raCorrectionValueCameraFocus(14) = -0.042
            Lane1Data.raCorrectionValueCameraFocus(15) = -0.041
            Lane1Data.raCorrectionValueCameraFocus(16) = -0.041
            Lane1Data.raCorrectionValueCameraFocus(17) = -0.048
            Lane1Data.raCorrectionValueCameraFocus(18) = -0.051
            Lane1Data.raCorrectionValueCameraFocus(19) = -0.045
            Lane1Data.raCorrectionValueCameraFocus(20) = -0.058
            Lane1Data.raCorrectionValueCameraFocus(21) = -0.047
            Lane1Data.raCorrectionValueCameraFocus(22) = -0.054
            Lane1Data.raCorrectionValueCameraFocus(23) = -0.059
            Lane1Data.raCorrectionValueCameraFocus(24) = -0.051
            Lane1Data.raCorrectionValueCameraFocus(25) = -0.07

            'Lane 2
            Lane2Data.zPositionATCylinder = -4.566
            Lane2Data.rHeightCalibration = 1.246
            Lane2Data.DistanceNozzleProduct = 1.2
            Lane2Data.DistanceNozzleDrop = 1.154

            Lane2Data.raMeasuredHeight(1) = 1.159
            Lane2Data.raMeasuredHeight(2) = 1.162
            Lane2Data.raMeasuredHeight(3) = 1.169
            Lane2Data.raMeasuredHeight(4) = 1.17
            Lane2Data.raMeasuredHeight(5) = 1.175
            Lane2Data.raMeasuredHeight(6) = 1.168
            Lane2Data.raMeasuredHeight(7) = 1.17
            Lane2Data.raMeasuredHeight(8) = 1.182
            Lane2Data.raMeasuredHeight(9) = 1.178
            Lane2Data.raMeasuredHeight(10) = 1.182
            Lane2Data.raMeasuredHeight(11) = 1.177
            Lane2Data.raMeasuredHeight(12) = 1.172
            Lane2Data.raMeasuredHeight(13) = 1.18
            Lane2Data.raMeasuredHeight(14) = 1.172
            Lane2Data.raMeasuredHeight(15) = 1.166
            Lane2Data.raMeasuredHeight(16) = 1.177
            Lane2Data.raMeasuredHeight(17) = 1.187
            Lane2Data.raMeasuredHeight(18) = 1.2
            Lane2Data.raMeasuredHeight(19) = 1.195
            Lane2Data.raMeasuredHeight(20) = 1.195
            Lane2Data.raMeasuredHeight(21) = 1.179
            Lane2Data.raMeasuredHeight(22) = 1.2
            Lane2Data.raMeasuredHeight(23) = 1.193
            Lane2Data.raMeasuredHeight(24) = 1.172
            Lane2Data.raMeasuredHeight(25) = 1.196

            Lane2Data.raCorrectionValueCameraFocus(1) = -0.35
            Lane2Data.raCorrectionValueCameraFocus(2) = -0.35
            Lane2Data.raCorrectionValueCameraFocus(3) = -0.347
            Lane2Data.raCorrectionValueCameraFocus(4) = -0.341
            Lane2Data.raCorrectionValueCameraFocus(5) = -0.34
            Lane2Data.raCorrectionValueCameraFocus(6) = -0.334
            Lane2Data.raCorrectionValueCameraFocus(7) = -0.341
            Lane2Data.raCorrectionValueCameraFocus(8) = -0.34
            Lane2Data.raCorrectionValueCameraFocus(9) = -0.328
            Lane2Data.raCorrectionValueCameraFocus(10) = -0.331
            Lane2Data.raCorrectionValueCameraFocus(11) = -0.327
            Lane2Data.raCorrectionValueCameraFocus(12) = -0.332
            Lane2Data.raCorrectionValueCameraFocus(13) = -0.338
            Lane2Data.raCorrectionValueCameraFocus(14) = -0.329
            Lane2Data.raCorrectionValueCameraFocus(15) = -0.338
            Lane2Data.raCorrectionValueCameraFocus(16) = -0.344
            Lane2Data.raCorrectionValueCameraFocus(17) = -0.333
            Lane2Data.raCorrectionValueCameraFocus(18) = -0.322
            Lane2Data.raCorrectionValueCameraFocus(19) = -0.309
            Lane2Data.raCorrectionValueCameraFocus(20) = -0.315
            Lane2Data.raCorrectionValueCameraFocus(21) = -0.315
            Lane2Data.raCorrectionValueCameraFocus(22) = -0.33
            Lane2Data.raCorrectionValueCameraFocus(23) = -0.309
            Lane2Data.raCorrectionValueCameraFocus(24) = -0.316
            Lane2Data.raCorrectionValueCameraFocus(25) = -0.338

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: LoadSiemensSimulatorValues()")
        End Try

    End Sub

    Public Sub LoadAllenBradleySimulatorValues()

        Try
            ' Lane 1
            Lane1Data.NozzleTipHeight = 1.956
            Lane1Data.LaserMeasurementToDatum = -6.113
            Lane1Data.NozzleTipDrop = 1.143
            Lane1Data.NozzleCalibToDatum_AxisPos = 25.0

            Lane1Data.LaserMeasurementToLFID(1) = -6.105
            Lane1Data.LaserMeasurementToLFID(2) = -6.12
            Lane1Data.LaserMeasurementToLFID(3) = -6.112
            Lane1Data.LaserMeasurementToLFID(4) = -6.121
            Lane1Data.LaserMeasurementToLFID(5) = -6.113
            Lane1Data.LaserMeasurementToLFID(6) = -6.118
            Lane1Data.LaserMeasurementToLFID(7) = -6.131
            Lane1Data.LaserMeasurementToLFID(8) = -6.121
            Lane1Data.LaserMeasurementToLFID(9) = -6.128
            Lane1Data.LaserMeasurementToLFID(10) = -6.134
            Lane1Data.LaserMeasurementToLFID(11) = -6.114
            Lane1Data.LaserMeasurementToLFID(12) = -6.133
            Lane1Data.LaserMeasurementToLFID(13) = -6.131
            Lane1Data.LaserMeasurementToLFID(14) = -6.126
            Lane1Data.LaserMeasurementToLFID(15) = -6.127
            Lane1Data.LaserMeasurementToLFID(16) = -6.14
            Lane1Data.LaserMeasurementToLFID(17) = -6.142
            Lane1Data.LaserMeasurementToLFID(18) = -6.131
            Lane1Data.LaserMeasurementToLFID(19) = -6.135
            Lane1Data.LaserMeasurementToLFID(20) = -6.128
            Lane1Data.LaserMeasurementToLFID(21) = -6.127
            Lane1Data.LaserMeasurementToLFID(22) = -6.124
            Lane1Data.LaserMeasurementToLFID(23) = -6.122
            Lane1Data.LaserMeasurementToLFID(24) = -6.132
            Lane1Data.LaserMeasurementToLFID(25) = -6.126

            ' Lane 2
            Lane2Data.NozzleTipHeight = 1.938
            Lane2Data.LaserMeasurementToDatum = -6.363
            Lane2Data.NozzleTipDrop = 1.118
            Lane2Data.NozzleCalibToDatum_AxisPos = 25.2

            Lane2Data.LaserMeasurementToLFID(1) = -6.31
            Lane2Data.LaserMeasurementToLFID(2) = -6.313
            Lane2Data.LaserMeasurementToLFID(3) = -6.314
            Lane2Data.LaserMeasurementToLFID(4) = -6.316
            Lane2Data.LaserMeasurementToLFID(5) = -6.306
            Lane2Data.LaserMeasurementToLFID(6) = -6.327
            Lane2Data.LaserMeasurementToLFID(7) = -6.298
            Lane2Data.LaserMeasurementToLFID(8) = -6.308
            Lane2Data.LaserMeasurementToLFID(9) = -6.31
            Lane2Data.LaserMeasurementToLFID(10) = -6.315
            Lane2Data.LaserMeasurementToLFID(11) = -6.305
            Lane2Data.LaserMeasurementToLFID(12) = -6.318
            Lane2Data.LaserMeasurementToLFID(13) = -6.316
            Lane2Data.LaserMeasurementToLFID(14) = -6.322
            Lane2Data.LaserMeasurementToLFID(15) = -6.323
            Lane2Data.LaserMeasurementToLFID(16) = -6.325
            Lane2Data.LaserMeasurementToLFID(17) = -6.323
            Lane2Data.LaserMeasurementToLFID(18) = -6.315
            Lane2Data.LaserMeasurementToLFID(19) = -6.317
            Lane2Data.LaserMeasurementToLFID(20) = -6.301
            Lane2Data.LaserMeasurementToLFID(21) = -6.308
            Lane2Data.LaserMeasurementToLFID(22) = -6.308
            Lane2Data.LaserMeasurementToLFID(23) = -6.22
            Lane2Data.LaserMeasurementToLFID(24) = -6.299
            Lane2Data.LaserMeasurementToLFID(25) = -6.3

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Failed to Execute: LoadAllenBradleySimulatorValues()")
        End Try

    End Sub
End Module
