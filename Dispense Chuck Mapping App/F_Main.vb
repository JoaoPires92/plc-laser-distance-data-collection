'App Developed by Joao Pires: JUL 2019

Public Class F_Main
    Dim SimulatorCounter As Short
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Load F_Main first Time

        ReDim Lane1Data.raMeasuredHeight(25) ' Define Array Size for LDS Height Measured for 25 Sensors
        ReDim Lane2Data.raMeasuredHeight(25)
        ReDim Lane1Data.raCorrectionValueCameraFocus(25) ' Define Array Size for Camera Focus Correction offsets for 25 Sensors
        ReDim Lane2Data.raCorrectionValueCameraFocus(25)
        ReDim Lane1Data.LaserMeasurementToLFID(25) ' Define Array Size for LDS Height Measured for 25 Sensors
        ReDim Lane2Data.LaserMeasurementToLFID(25)

        Label9.Visible = False
        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = False
        TextBox1.Enabled = False
        TextBox4.Enabled = False
        Timer1.Enabled = False
        GroupBox4.Visible = False
        RichTextBox1.Enabled = False
        TextBox3.Enabled = False
        RichTextBox2.Enabled = False
        RichTextBox3.Enabled = False
        PLC_Connection = False
        PLC_ConfigurationLoadedOK = False
        PLC_CommunicationLoadedOK = False
        RichTextBox4.Enabled = False
        RichTextBox5.Enabled = False
        RichTextBox6.Enabled = False
        RichTextBox7.Enabled = False
        GroupBox2.Text = "Dispenser 1 - Lane 1"
        GroupBox3.Text = "Dispenser 2 - Lane 2"
        boSiemensPLC = False
        boAllenBradleyPLC = False
        Timer2.Enabled = True ' Start 1s timer to display time and date
        Button12.Enabled = False

        Load_Machine_Name() ' Load File with Machine Names

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Connect To PLC Button

        boSimulatorMode = True ' Enable / Disable Simulator Mode to work Offline

        If (MachineSelected <> "") Then

            If boSiemensPLC = True Then
                Load_Siemens_PLC_Communication() ' Call Function to Load Siemens PLC communication settings (Loads PLC IP Address, Rack, and Slot number of CPU)
                If PLC_CommunicationLoadedOK And boSimulatorMode = False Then
                    ConnectTo_Siemens_PLC() ' Call function to connect to Siemens PLC (Return PLC_Connection)
                ElseIf PLC_CommunicationLoadedOK And boSimulatorMode = True Then
                    PLC_Connection = True
                End If

            ElseIf boAllenBradleyPLC = True Then
                Load_AB_PLC_Communication() ' Call Function to Load AB PLC communication settings (Loads PLC IP Address, Rack, and Slot number of CPU)
                If PLC_CommunicationLoadedOK And boSimulatorMode = False Then
                    ConnectTo_AB_PLC() ' Call function to connect to AB PLC (Return PLC_Connection)
                ElseIf PLC_CommunicationLoadedOK And boSimulatorMode = True Then
                    PLC_Connection = True
                End If

            End If

            If PLC_Connection = True Then
                Label2.ForeColor = Color.Green
                If boSiemensPLC = True Then Label2.Text = "Connected to: " & Siemens_PLC_IP & " - SIEMENS S7-319F 3PN/DP"
                If boAllenBradleyPLC = True Then Label2.Text = "Connected to: " & AB_PLC_IP & " - Allen-Bradley-Logix5582E"
                Button3.BackColor = Color.LightGreen
                Button5.BackColor = Color.LightGreen
                Button3.Enabled = True
                Button2.Enabled = True
                Button1.Enabled = False
                ComboBox1.Enabled = False
                Button5.Enabled = True
            Else
                Label2.ForeColor = Color.IndianRed
                Label2.Text = "Unable to Connect to PLC"
                MsgBox("Unable to Connect to PLC", MsgBoxStyle.Critical, "PLC Connection Timeout")
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Disconnect from PLC Button
        If boSimulatorMode = False Then
            If boSiemensPLC = True Then DisconnectFrom_Siemens_PLC() ' Call Function Disconnecto from PLC
            If boAllenBradleyPLC = True Then DisconnectFrom_AB_PLC()
        ElseIf boSimulatorMode = True Then
            PLC_Connection = False
        End If
        Label2.ForeColor = Color.IndianRed
        Label2.Text = "Disconnected"
        Button3.BackColor = Color.LightGray
        Button4.BackColor = Color.LightGray
        Button5.BackColor = Color.LightGray
        Button12.BackColor = Color.LightGray
        Button3.Enabled = False
        Button4.Enabled = False
        Button1.Enabled = True
        Button2.Enabled = False
        Button5.Enabled = False
        Button12.Enabled = False
        ComboBox1.Enabled = True
        TextBox1.Text = ""
        TextBox4.Text = ""
        strLane1Data = ""
        strLane2Data = ""
        RichTextBox4.Text = ""
        RichTextBox5.Text = ""
        RichTextBox6.Text = ""
        RichTextBox7.Text = ""
        TextBox1.Enabled = False
        TextBox4.Enabled = False
        GroupBox4.Visible = False
        Timer1.Enabled = False
        PLC_ConfigurationLoadedOK = False
        RichTextBox4.BackColor = Color.White
        RichTextBox5.BackColor = Color.White
        RichTextBox6.BackColor = Color.White
        RichTextBox7.BackColor = Color.White
        RichTextBox4.Enabled = False
        RichTextBox5.Enabled = False
        RichTextBox6.Enabled = False
        RichTextBox7.Enabled = False
        GroupBox2.Text = "Dispenser 1 - Lane 1"
        GroupBox3.Text = "Dispenser 2 - Lane 2"
        Lane1LDSAverage = 0
        Lane2LDSAverage = 0
        L1_CalibPin_HeightDifference = 0
        L2_CalibPin_HeightDifference = 0
        L1_ChuckRunOut = 0
        L2_ChuckRunOut = 0


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Get Data From PLC Button
        If (MachineSelected <> "" And PLC_Connection) Then

            TextBox1.Text = ""
            TextBox4.Text = ""
            RichTextBox4.Text = ""
            RichTextBox5.Text = ""
            RichTextBox6.Text = ""
            RichTextBox7.Text = ""
            strLane1Data = ""
            strLane2Data = ""
            Lane1LDSAverage = 0
            Lane2LDSAverage = 0
            L1_CalibPin_HeightDifference = 0
            L2_CalibPin_HeightDifference = 0
            L1_ChuckRunOut = 0
            L2_ChuckRunOut = 0

            If boSimulatorMode = False Then
                GetLotCodeLane1() 'Call Function to Get Lane 1 Lot Code Information
                GetLotCodeLane2() 'Call Function to Get Lane 2 Lot Code Information

                GetLane1Data() ' Call Function to Get Lane 1 Data from PLC
                GetLane2Data() ' Call Function to Get Lane 2 Data from PLC
            ElseIf boSimulatorMode = True Then
                Lane1Data.LotCode = "SIMULATOR L1"
                Lane2Data.LotCode = "SIMULATOR L2"
                If boSiemensPLC Then LoadSiemensSimulatorValues()
                If boAllenBradleyPLC Then LoadAllenBradleySimulatorValues()
            End If

            ConvertLane1Data_ToString() ' Function to Convert Received data to String
            ConvertLane2Data_ToString()

            CalculateCalibPinHeightDifferenceToChuck() ' Call Function to Calculate Difference Between Pin and Chuck
            CalculateChuckRunOut() ' Call Function to Calculate Chuck Runout


            ' Display PLC Data Results 
            If Lane1Data.LotCode = "No Lot Online" Then GroupBox2.Text = "Dispenser 1 - Lane 1" Else GroupBox2.Text = "Dispenser 1 - Lane 1 Lot: " & Lane1Data.LotCode
                If Lane2Data.LotCode = "No Lot Online" Then GroupBox3.Text = "Dispenser 2 - Lane 2" Else GroupBox3.Text = "Dispenser 2 - Lane 2 Lot: " & Lane2Data.LotCode

                TextBox1.Text = strLane1Data
                TextBox4.Text = strLane2Data

                If L1_CalibPin_HeightDifference >= 0.1 Or L1_CalibPin_HeightDifference <= -0.1 Then RichTextBox4.BackColor = Color.Red Else RichTextBox4.BackColor = Color.Green
                If L2_CalibPin_HeightDifference >= 0.1 Or L2_CalibPin_HeightDifference <= -0.1 Then RichTextBox7.BackColor = Color.Red Else RichTextBox7.BackColor = Color.Green
                RichTextBox4.Text = vbTab & Format(L1_CalibPin_HeightDifference, "0.000") & " mm"
                RichTextBox7.Text = vbTab & Format(L2_CalibPin_HeightDifference, "0.000") & " mm"

                If L1_ChuckRunOut >= 0.1 Or L1_ChuckRunOut <= -0.1 Then RichTextBox5.BackColor = Color.Red Else RichTextBox5.BackColor = Color.Green
                If L2_ChuckRunOut >= 0.1 Or L2_ChuckRunOut <= -0.1 Then RichTextBox6.BackColor = Color.Red Else RichTextBox6.BackColor = Color.Green
                RichTextBox5.Text = vbTab & Format(L1_ChuckRunOut, "0.000") & " mm"
                RichTextBox6.Text = vbTab & Format(L2_ChuckRunOut, "0.000") & " mm"

                Button4.BackColor = Color.LightGreen
                Button12.BackColor = Color.LightGreen
                Button4.Enabled = True
                Button12.Enabled = True
                TextBox1.Enabled = True
                TextBox4.Enabled = True
                RichTextBox4.Enabled = True
                RichTextBox5.Enabled = True
                RichTextBox6.Enabled = True
                RichTextBox7.Enabled = True

            End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Export to Excel File Button
        Dim AuxStr As String

        SaveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

        DialogResult = SaveFileDialog1.ShowDialog()

        AuxStr = SaveFileDialog1.FileName

        If AuxStr <> Nothing And DialogResult <> Windows.Forms.DialogResult.Cancel Then
            If SaveDataIntoExcelSheet(AuxStr) Then
                MsgBox("File has been Saved.", MsgBoxStyle.Information)
            Else
                MsgBox("Failed to Save File.", MsgBoxStyle.Critical)
            End If
        Else
            MsgBox("File has not been saved.", MsgBoxStyle.Information)

        End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ' Drop Down Menu to Select Target Machine

        boSiemensPLC = False
        boAllenBradleyPLC = False
        Button1.Enabled = False

        If (ComboBox1.SelectedItem <> "") Then

            MachineSelected = ComboBox1.SelectedItem

            If MachineSelected.StartsWith("H") Then
                boSiemensPLC = True
                Load_Siemens_PLC_Configuration() ' Load correct DB adress values from Machine Selected
            ElseIf MachineSelected.StartsWith("W") Then
                boAllenBradleyPLC = True
                Load_AB_PLC_Configuration() ' Load correct Tag name from Machine Selected
            Else
                Select Case MsgBox("Is the selected machine being controlled by an Allen Bradley PLC?", vbYesNo Or vbQuestion Or vbDefaultButton1)
                    Case vbYes
                        boAllenBradleyPLC = True
                        Load_AB_PLC_Configuration()
                    Case vbNo
                        Select Case MsgBox("Is the selected machine being controlled by a Siemens PLC?", vbYesNo Or vbQuestion Or vbDefaultButton1)
                            Case vbYes
                                boSiemensPLC = True
                                Load_Siemens_PLC_Configuration() ' Load correct DB adress values from Machine Selected
                            Case vbNo
                                MsgBox("Unable to load configuration for the selected machine. No PLC type selected.", MsgBoxStyle.Critical)
                                Exit Sub
                        End Select
                End Select

            End If

            If PLC_ConfigurationLoadedOK = True Then
                    Button1.Enabled = True 'Enable Connecto to PLC Button
                Else
                    Button1.Enabled = False
                    MsgBox("Failed to Load PLC Configuration for Target Machine Selected", MsgBoxStyle.Critical)
                End If
                PLC_ConfigurationLoadedOK = False
            End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'LDS Actual Height Reading Button
        If (MachineSelected <> "" And PLC_Connection) Then

            PLC_DataReceived = ""
            GroupBox4.Visible = True
            Button3.Enabled = False

            If boSimulatorMode = True Then
                If boSiemensPLC Then LoadSiemensSimulatorValues()
                If boAllenBradleyPLC Then LoadAllenBradleySimulatorValues()
                SimulatorCounter = 1
            End If

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' Timmer runs every 250ms
        If (MachineSelected <> "" And PLC_Connection) Then

            If boSimulatorMode = False Then
                ReadActualHeightLane1() ' Read PLC Live LDS Measurements from Lane 1
                ReadActualHeightLane2() ' Read PLC Live LDS Measurements from Lane 2
            ElseIf boSimulatorMode = True Then
                SimulatorCounter = SimulatorCounter + 1
                If SimulatorCounter > 25 Then
                    SimulatorCounter = 1
                End If
                If boSiemensPLC Then
                    Lane1Data.rActualHeight = Lane1Data.raMeasuredHeight(SimulatorCounter)
                    Lane2Data.rActualHeight = Lane2Data.raMeasuredHeight(SimulatorCounter)
                ElseIf boAllenBradleyPLC Then
                    Lane1Data.LaserLiveMeasurement = Lane1Data.LaserMeasurementToLFID(SimulatorCounter)
                    Lane2Data.LaserLiveMeasurement = Lane2Data.LaserMeasurementToLFID(SimulatorCounter)
                End If
            End If
            If boSiemensPLC = True Then
                Label3.Text = Format(Lane1Data.rActualHeight, "0.000")
                Label4.Text = Format(Lane2Data.rActualHeight, "0.000")
            End If

            If boAllenBradleyPLC = True Then
                Label3.Text = Format(Lane1Data.LaserLiveMeasurement, "0.000")
                Label4.Text = Format(Lane2Data.LaserLiveMeasurement, "0.000")
            End If

        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Close LDS Actual Reading Screen
        Timer1.Enabled = False
        GroupBox4.Visible = False
        Button3.Enabled = True
        Label3.Text = "0.000"
        Label4.Text = "0.000"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' LDS Start Trigger Button
        Timer1.Enabled = True ' Enable Timmer to Get LDS Online Readings
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' LDS Stop Trigger Button
        Timer1.Enabled = False ' Disable Timmer to Stop getting LDS Online Readings
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If Button9.Text = "Hide LDS Reading" Then
            Label3.Visible = False
            Button9.Text = "Show LDS Reading"
        Else
            Label3.Visible = True
            Button9.Text = "Hide LDS Reading"
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        If Button10.Text = "Hide LDS Reading" Then
            Label4.Visible = False
            Button10.Text = "Show LDS Reading"
        Else
            Label4.Visible = True
            Button10.Text = "Hide LDS Reading"
        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        '1s Timer to display date and time 
        Label9.Text = Date.Now
        Label9.Visible = True
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'Export Data to PDF File Button

        'PrintPreviewDialog1.WindowState = 2
        'PrintPreviewDialog1.Document = PrintDocumentToPDF()
        'PrintPreviewDialog1.ShowDialog()

        PrintDialog1.Document = PrintDocumentToPDF()
        PrintDialog1.Document.Print()

    End Sub

End Class
