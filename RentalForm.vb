Option Explicit On
Option Strict On
Option Compare Binary
'Jacob Horsley
'RCET 0265
'Spring 2025
'URL: 



'[ ] Remove invalid data from textbox 
'[ ] Set tab order
'[ ] Validate odometer readings
'[ ] Validate number of days
'[ ] single message box to display an improper input
'[ ] do special mile math for each range
'[ ] add radio button functions discounts
'[ ] 

Imports System.Net.Security

Public Class RentalForm

    Private Sub SetDefaults()
        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""


    End Sub

    Private Sub TotalMath()

    End Sub

    Private Function dailyDollars() As Integer
        Dim dailyCharge As Integer
        Dim days As Integer

        If Integer.TryParse(DaysTextBox.Text, days) Then
            dailyCharge = days * 15
            Return dailyCharge
        Else
            Return 0
        End If
    End Function
    Private Function test() As Integer

        Dim beginOdometer As Integer
        Dim endOdometer As Integer
        Dim totalMiles As Integer
        If Integer.TryParse(BeginOdometerTextBox.Text, beginOdometer) AndAlso
           Integer.TryParse(EndOdometerTextBox.Text, endOdometer) Then
            totalMiles = endOdometer - beginOdometer
            Return totalMiles
        Else
            Return 0
        End If

    End Function
    Private Function calculateMiles() As Double
        Dim BeginOdometer As Integer
        Dim EndOdometer As Integer
        Dim miles As Double
        Dim kilometers As Double
        Dim twelveCentMiles As Double
        Dim tenCentMiles As Double
        Dim _mileageCharge As Double


        If Integer.TryParse(BeginOdometerTextBox.Text, BeginOdometer) AndAlso
            Integer.TryParse(EndOdometerTextBox.Text, EndOdometer) Then
            miles = EndOdometer - BeginOdometer
            'Return miles
        Else
            Return 0
        End If

        If MilesradioButton.Checked = True Then

            Select Case miles
                Case 0 To 200
                    _mileageCharge = 0
                Case 201 To 499
                    _mileageCharge += (miles - 200) * 0.12
                Case Else
                    tenCentMiles = miles - 499
                    twelveCentMiles += miles - tenCentMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                    _mileageCharge += (tenCentMiles * 0.1)

            End Select
        End If

        If KilometersradioButton.Checked = True Then

            Select Case kilometers
                Case 0 To 200
                    _mileageCharge = 0
                Case 201 To 499
                    _mileageCharge = ((miles / 0.62) - 200) * 0.12
                Case Else
                    tenCentMiles = (miles / 0.62) - 499
                    twelveCentMiles += (miles / 0.62) - tenCentMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                    _mileageCharge += (tenCentMiles * 0.1)
                    _mileageCharge = ((miles / 0.62) - 200) * 0.1
            End Select
        End If

        MileageChargeTextBox.Text = _mileageCharge.ToString("C2")

        Return _mileageCharge

    End Function

    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        Dim miles As Double = calculateMiles()
        Dim discount As Double = 0
        Dim totalCharge As Double
        Dim discountAmount As Double

        totalCharge = dailyDollars() + calculateMiles()
        If AAAcheckbox.Checked = True Then
            discount += 0.05
            discountAmount = totalCharge * discount
        End If
        If Seniorcheckbox.Checked = True Then
            discount += 0.03
            discountAmount = totalCharge * discount
        End If
        If discount > 0 Then
            discountAmount = totalCharge * discount
            totalCharge -= totalCharge * discount
        End If

        TotalDiscountTextBox.Text = discountAmount.ToString("C2")
        TotalChargeTextBox.Text = "$" & totalCharge.ToString("C2")
        DayChargeTextBox.Text = dailyDollars.ToString("C2")
        TotalMilesTextBox.Text = CStr(test() & " mi.")
    End Sub


    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        If ValidInputs() Then
            Dim summary As String =
            "Name: " & NameTextBox.Text & vbCrLf &
            "Address: " & AddressTextBox.Text & vbCrLf &
            "City: " & CityTextBox.Text & vbCrLf &
            "State: " & StateTextBox.Text & vbCrLf &
            "Zip Code: " & ZipCodeTextBox.Text & vbCrLf &
            "Beginning Odometer: " & BeginOdometerTextBox.Text & vbCrLf &
            "Ending Odometer: " & EndOdometerTextBox.Text & vbCrLf &
            "Days: " & DaysTextBox.Text & vbCrLf &
            "Total Miles: " & TotalMilesTextBox.Text & vbCrLf &
            "Mileage Charge: " & MileageChargeTextBox.Text & vbCrLf &
            "Total Discount: " & TotalDiscountTextBox.Text & vbCrLf &
            "Total Charge: " & TotalChargeTextBox.Text

            'Test box to see if stuff shoes up
            MsgBox(summary, MsgBoxStyle.Information, "rental summary")
        End If
    End Sub

    Private Function ValidInputs() As Boolean
        Dim valid As Boolean = True
        Dim message As String

        If NameTextBox.Text = "" Then 'OrElse not lettersOnly(NametextBox.text) then 'Optional letter only thing
            valid = False
            NameTextBox.Clear()
            NameTextBox.Focus()
            message &= "Name must be added."
        End If
        If AddressTextBox.Text = "" Then
            valid = False
            AddressTextBox.Clear()
            AddressTextBox.Focus()
            message &= "Address must be added."
        End If
        If CityTextBox.Text = "" Then
            valid = False
            CityTextBox.Clear()
            CityTextBox.Focus()
            message &= "City must be added."
        End If
        If StateTextBox.Text = "" Then
            valid = False
            StateTextBox.Clear()
            StateTextBox.Focus()
            message &= "State must be added."
        End If
        If ZipCodeTextBox.Text = "" Then
            valid = False
            ZipCodeTextBox.Clear()
            ZipCodeTextBox.Focus()
            message &= "Zip must be added."
        End If
        If BeginOdometerTextBox.Text = "" Then
            valid = False
            BeginOdometerTextBox.Clear()
            BeginOdometerTextBox.Focus()
            message &= "Beginning Odometer reading must be added."
        End If
        If EndOdometerTextBox.Text = "" Then
            valid = False
            EndOdometerTextBox.Clear()
            EndOdometerTextBox.Focus()
            message &= "Final Odometer Reading must be added."
        End If
        If DaysTextBox.Text = "" Then
            valid = False
            DaysTextBox.Clear()
            DaysTextBox.Focus()
            message &= "Days occupied must be added."
        End If
        If CInt(BeginOdometerTextBox.Text) >= CInt(EndOdometerTextBox.Text) Then
            valid = False
            BeginOdometerTextBox.Clear()
            EndOdometerTextBox.Clear()
            message &= "Beginning odometer reading must be less than the odometer reading at return..."
        End If
        'If TotalMilesTextBox.Text = "" Then
        '    valid = False
        '    TotalMilesTextBox.Clear()
        '    TotalMilesTextBox.Focus()
        '    message &= "Total Miles must be added."
        'End If
        'If MileageChargeTextBox.Text = "" Then
        '    valid = False
        '    MileageChargeTextBox.Clear()
        '    MileageChargeTextBox.Focus()
        '    message &= "Mileage Charge must be added."
        'End If
        'If TotalDiscountTextBox.Text = "" Then
        '    valid = False
        '    TotalDiscountTextBox.Clear()
        '    TotalDiscountTextBox.Focus()
        '    message &= "Total Discount must be added."
        'End If
        'If TotalChargeTextBox.Text = "" Then
        '    valid = False
        '    TotalChargeTextBox.Clear()
        '    TotalChargeTextBox.Focus()
        '    message &= "Total Charge must be added."
        'End If
        If Not valid Then
            MsgBox(message, MsgBoxStyle.Exclamation, "User input fail")
        End If
        Return valid
    End Function

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click

        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Confirm Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Application.Exit()
        End If

    End Sub


    Private Sub OnlyLetters(sender As Object, e As KeyPressEventArgs) Handles NameTextBox.KeyPress, CityTextBox.KeyPress, StateTextBox.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub OnlyNumbers(sender As Object, e As KeyPressEventArgs) Handles ZipCodeTextBox.KeyPress, BeginOdometerTextBox.KeyPress, EndOdometerTextBox.KeyPress, DaysTextBox.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        NameTextBox.Clear()
        AddressTextBox.Clear()
        CityTextBox.Clear()
        StateTextBox.Clear()
        ZipCodeTextBox.Clear()
        BeginOdometerTextBox.Clear()
        EndOdometerTextBox.Clear()
        DaysTextBox.Clear()
        TotalMilesTextBox.Clear()
        MileageChargeTextBox.Clear()
        DayChargeTextBox.Clear()
        TotalDiscountTextBox.Clear()
        TotalChargeTextBox.Clear()
    End Sub
End Class
