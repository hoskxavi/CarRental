'Xavier Hoskins
'RCET 0265
'Fall 2022
'Car Rental
'https://www.github.com/hoskxavi/CarRental.git

Option Explicit On
Option Strict On
Option Compare Binary
Public Class RentalForm
    Dim totalMilesSummary As Integer
    Dim totalCustomersSummary As Integer
    Dim totalChargesSummary As Double

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim startMiles As Integer
        Dim endMiles As Integer
        Dim numberOfDays As Integer
        Dim totalMiles As Integer
        Dim milesCharge As Double
        Dim dayCharge As Double
        Dim discount As Double
        Dim milesConvert As Double
        Dim totalCharge As Double
        Static count As Integer


        SummaryButton.Enabled = True

        If AAAcheckbox.Checked And Seniorcheckbox.Checked Then
            discount = 0.92
        ElseIf AAAcheckbox.Checked Then
            discount = 0.95
        ElseIf Seniorcheckbox.Checked Then
            discount = 0.97
        Else
            discount = 1
        End If

        If MilesradioButton.Checked Then
            milesConvert = 1
        ElseIf KilometersradioButton.Checked Then
            milesConvert = 0.62
        End If

        startMiles = ConvertandValidate(BeginOdometerTextBox.Text, 1)
        endMiles = ConvertandValidate(EndOdometerTextBox.Text, 2)
        numberOfDays = ConvertandValidate(DaysTextBox.Text, 3)

        If startMiles > endMiles Then
            MsgBox("Error! Start Miles cannot be More than End Miles!")
            Exit Sub
        Else
            totalMiles = endMiles - startMiles
            totalMilesSummary += totalMiles
            TotalMilesTextBox.Text = CStr(totalMiles * milesConvert) & "mi"
            dayCharge = numberOfDays * 15
            DayChargeTextBox.Text = "$" & CStr(dayCharge)

            If totalMiles < 201 Then
                milesCharge = totalMiles * milesConvert * 0
                MileageChargeTextBox.Text = "$" & CStr(milesCharge)
            ElseIf totalMiles > 200 Then
                milesCharge = totalMiles * milesConvert * 0.12
                MileageChargeTextBox.Text = "$" & CStr(milesCharge)
            ElseIf totalMiles > 500 Then
                milesCharge = totalMiles * milesConvert * 0.1
                MileageChargeTextBox.Text = "$" & CStr(milesCharge)
            End If

            count += 1
            totalCustomersSummary = count
            totalCharge = (milesCharge + dayCharge) * discount
            totalChargesSummary += totalCharge
            TotalChargeTextBox.Text = "$" & CStr(totalCharge)
            TotalDiscountTextBox.Text = "$" & CStr((milesCharge + dayCharge) - totalCharge)
        End If

    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        Dim msg As String = "Are you sure you wish to quit?"
        Dim msgStyle = MsgBoxStyle.YesNo
        Dim response = MsgBox(msg, msgStyle)

        If response = MsgBoxResult.Yes Then
            Me.Close()
        End If

    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        AddressTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        CityTextBox.Text = ""
        DayChargeTextBox.Text = ""
        DaysTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        NameTextBox.Text = ""
        StateTextBox.Text = ""
        TotalChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        ZipCodeTextBox.Text = ""

        MilesradioButton.Checked = True
        AAAcheckbox.Checked = False
        Seniorcheckbox.Checked = False
    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SummaryButton.Enabled = False
    End Sub

    Function ConvertandValidate(input As String, location As Integer) As Integer
        Dim convertedNumber As Integer
        Static count As Integer

        Try
            convertedNumber = CInt(input)
            count = 0
            Return convertedNumber
        Catch ex As Exception
            count += 1
            Select Case location
                Case 1
                    BeginOdometerTextBox.BackColor = Color.Red
                Case 2
                    EndOdometerTextBox.BackColor = Color.Red
                Case 3
                    DaysTextBox.BackColor = Color.Red
            End Select
            If count > 1 Then
                Exit Function
            Else
                MsgBox("Error! All fields must be filled in to continue.")
            End If
            Return 0
        End Try

    End Function

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        MsgBox("Total Customers: " & CStr(totalCustomersSummary) & vbCrLf &
               "Total Miles: " & CStr(totalMilesSummary) & "mi" & vbCrLf &
               "Total Charges: " & "$" & CStr(totalChargesSummary))
    End Sub

End Class
