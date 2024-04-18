Imports System.Drawing.Printing
Module Printing

    Dim __CurrentPage As Short
    Public Function PrintDocumentToPDF() As PrintDocument

        Dim Print_document As New PrintDocument

        Try
            __CurrentPage = 1

            Print_document.PrinterSettings.PrinterName = "Microsoft Print To PDF"

            AddHandler Print_document.PrintPage, AddressOf Print_PDF_Page

            Return Print_document

        Catch ex As Exception

            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "PrepareToPrintDocument()")
            Exit Function

        End Try

    End Function

    Private Sub Print_PDF_Page(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)

        Try

            Dim __drawFont As New Font("Arial", 10, FontStyle.Bold)
            Dim __drawFont1 As New Font("Arial", 10)
            Dim __drawFont2 As New Font("Arial", 14, FontStyle.Bold)
            Dim __drawFont3 As New Font("Arial", 10)
            Dim __drawFont4 As New Font("Arial", 28, FontStyle.Bold)
            Dim __drawFont5 As New Font("Arial", 10, FontStyle.Bold)
            Dim __drawFont6 As New Font("Arial", 12, FontStyle.Bold)
            Dim __drawFont7 As New Font("Arial", 10, FontStyle.Bold)
            Dim __drawBrush As New SolidBrush(Color.Black)
            Dim __StringFormat1 As New StringFormat(StringFormatFlags.NoClip)
            Dim __BlackPen As New Pen(Color.Black)


            Dim __CurrMargin_X As Short = 100
            Dim __CurrMargin_Y As Short = 100

            Dim __Aux3 As Short = 0

            __StringFormat1.Alignment = StringAlignment.Near

            If __CurrentPage = 1 Then

                Dim __Image As Image = New Bitmap(650, 900)
                e.Graphics.DrawImage(__Image, __CurrMargin_X, __CurrMargin_Y)

                ' Draw Company Logo
                Dim __Img_CompanyLogo As Image = Image.FromFile("Resources\Abbott.png")
                Dim __Img2 As Image = New Bitmap(200, 110)
                Dim __g As Graphics = Graphics.FromImage(__Img2)
                __g.DrawImage(__Img_CompanyLogo, 0, 0, 200, 110)
                e.Graphics.DrawImage(__Img2, 45, 5)

                __g = Nothing
                __Img2 = Nothing
                __Img_CompanyLogo = Nothing

                e.Graphics.DrawString("Abbott Diabetes Care", __drawFont4, __drawBrush, 275, 35, __StringFormat1)
                e.Graphics.DrawString("Engineering Controlled Document", __drawFont5, __drawBrush, 350, 78, __StringFormat1)


                e.Graphics.DrawString("Dispense Chuck Mapping Report", __drawFont2, __drawBrush, 250, 141, __StringFormat1)
                e.Graphics.DrawRectangle(__BlackPen, 25, 135, 780, 35)



                e.Graphics.DrawString("MACHINE", __drawFont, __drawBrush, 25, 195, __StringFormat1)
                e.Graphics.DrawRectangle(__BlackPen, 25, 215, 350, 30)
                e.Graphics.DrawString(MachineSelected, __drawFont3, __drawBrush, 35, 222, __StringFormat1)


                e.Graphics.DrawString("DATE & TIME", __drawFont, __drawBrush, 580, 195, __StringFormat1)
                e.Graphics.DrawRectangle(__BlackPen, 580, 215, 225, 30)
                e.Graphics.DrawString(Date.Now, __drawFont3, __drawBrush, 590, 222, __StringFormat1)


                e.Graphics.DrawString("RESULTS", __drawFont6, __drawBrush, 377, 285, __StringFormat1)
                e.Graphics.DrawRectangle(__BlackPen, 25, 275, 780, 35)

                e.Graphics.DrawRectangle(__BlackPen, 25, 310, 780, 825)
                e.Graphics.DrawLine(__BlackPen, 422, 310, 422, 1135)

                e.Graphics.DrawString("Dispenser 1 - Lane 1", __drawFont6, __drawBrush, 125, 320, __StringFormat1)
                e.Graphics.DrawString("Dispenser 2 - Lane 2", __drawFont6, __drawBrush, 525, 320, __StringFormat1)
                e.Graphics.DrawLine(__BlackPen, 25, 350, 805, 350)

                'Lane1 Data
                e.Graphics.DrawString("LDS Measurement to Datum: ", __drawFont1, __drawBrush, 35, 365, __StringFormat1)
                If boAllenBradleyPLC Then e.Graphics.DrawString(Format(Lane1Data.LaserMeasurementToDatum, "0.000"), __drawFont1, __drawBrush, 230, 365, __StringFormat1)
                If boSiemensPLC Then e.Graphics.DrawString(Format(Lane1Data.rHeightCalibration, "0.000"), __drawFont1, __drawBrush, 230, 365, __StringFormat1)

                Dim YposLane1 = 390
                For i = 1 To 25
                    e.Graphics.DrawString("LDS Measurement to [Sensor " & i & "]: ", __drawFont1, __drawBrush, 35, YposLane1, __StringFormat1)
                    If boAllenBradleyPLC Then e.Graphics.DrawString(Format(Lane1Data.LaserMeasurementToLFID(i), "0.000"), __drawFont1, __drawBrush, 255, YposLane1, __StringFormat1)
                    If boSiemensPLC Then e.Graphics.DrawString(Format(Lane1Data.raMeasuredHeight(i), "0.000"), __drawFont1, __drawBrush, 255, YposLane1, __StringFormat1)
                    YposLane1 += 25
                Next

                e.Graphics.DrawString("LDS Measurement [1 to 25] Average: ", __drawFont7, __drawBrush, 35, 1015, __StringFormat1)
                e.Graphics.DrawString(Format(Lane1LDSAverage, "0.000"), __drawFont7, __drawBrush, 292, 1015, __StringFormat1)

                e.Graphics.DrawString("Chuck To Datum Height Difference: ", __drawFont7, __drawBrush, 35, 1070, __StringFormat1)
                e.Graphics.DrawString(Format(L1_CalibPin_HeightDifference, "0.000") & " mm", __drawFont7, __drawBrush, 282, 1070, __StringFormat1)

                e.Graphics.DrawString("Chuck Run Out: ", __drawFont7, __drawBrush, 35, 1095, __StringFormat1)
                e.Graphics.DrawString(Format(L1_ChuckRunOut, "0.000") & " mm", __drawFont7, __drawBrush, 148, 1095, __StringFormat1)

                'Lane2 Data
                e.Graphics.DrawString("LDS Measurement to Datum: ", __drawFont1, __drawBrush, 432, 365, __StringFormat1)
                If boAllenBradleyPLC Then e.Graphics.DrawString(Format(Lane2Data.LaserMeasurementToDatum, "0.000"), __drawFont1, __drawBrush, 627, 365, __StringFormat1)
                If boSiemensPLC Then e.Graphics.DrawString(Format(Lane2Data.rHeightCalibration, "0.000"), __drawFont1, __drawBrush, 627, 365, __StringFormat1)

                Dim YposLane2 = 390
                For i = 1 To 25
                    e.Graphics.DrawString("LDS Measurement to [Sensor " & i & "]: ", __drawFont1, __drawBrush, 432, YposLane2, __StringFormat1)
                    If boAllenBradleyPLC Then e.Graphics.DrawString(Format(Lane2Data.LaserMeasurementToLFID(i), "0.000"), __drawFont1, __drawBrush, 652, YposLane2, __StringFormat1)
                    If boSiemensPLC Then e.Graphics.DrawString(Format(Lane2Data.raMeasuredHeight(i), "0.000"), __drawFont1, __drawBrush, 652, YposLane2, __StringFormat1)
                    YposLane2 += 25
                Next

                e.Graphics.DrawString("LDS Measurement [1 to 25] Average: ", __drawFont7, __drawBrush, 432, 1015, __StringFormat1)
                e.Graphics.DrawString(Format(Lane2LDSAverage, "0.000"), __drawFont7, __drawBrush, 688, 1015, __StringFormat1)

                e.Graphics.DrawString("Chuck To Datum Height Difference: ", __drawFont7, __drawBrush, 432, 1070, __StringFormat1)
                e.Graphics.DrawString(Format(L2_CalibPin_HeightDifference, "0.000") & " mm", __drawFont7, __drawBrush, 678, 1070, __StringFormat1)

                e.Graphics.DrawString("Chuck Run Out: ", __drawFont7, __drawBrush, 432, 1095, __StringFormat1)
                e.Graphics.DrawString(Format(L2_ChuckRunOut, "0.000") & " mm", __drawFont7, __drawBrush, 545, 1095, __StringFormat1)


                e.Graphics.DrawString("Page 1 of 1", __drawFont1, __drawBrush, 729, 1145, __StringFormat1)

            ElseIf __CurrentPage = 2 Then

            End If

            __CurrentPage += 1

            If __CurrentPage >= 1 Then
                e.HasMorePages = False
                __CurrentPage = 1
            Else
                e.HasMorePages = True
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Exclamation, "Print_PDF_Page()")
            Exit Sub
        End Try

    End Sub
End Module
