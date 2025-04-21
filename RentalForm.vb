Option Explicit On
Option Strict On
Option Compare Binary
'Jacob Horsley
'RCET 0265
'Spring 2025
'URL: https://github.com/horsjaco117/CarRental



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

            Select Case miles
                Case 0 To 200
                    _mileageCharge = 0
                Case 201 To 499
                    _mileageCharge = (miles - 200) * 0.12
                Case Else
                    tenCentMiles = (miles) - 499
                    twelveCentMiles += (miles) - tenCentMiles - 200
                    _mileageCharge += (twelveCentMiles * 0.12)
                    _mileageCharge += (tenCentMiles * 0.1)
                    ' _mileageCharge = ((miles) - 200) * 0.1'
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

        ValidInputs()

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
        TotalMilesTextBox.Text = MeasureofDistance()

        SummaryButton.Enabled = True

        UserSummary(True)
        MileSummary(True, CInt(miles))
        MoneySummary(True, CInt(totalCharge))
    End Sub


    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click
        If ValidInputs() Then
            Dim summary As String = ""

            summary &= "Number of Customers: " & CStr(UserSummary(False)) & vbNewLine
            summary &= "Total miles Accumulated: " & CStr(MileSummary(False, 0)) & " mi." & vbCrLf
            summary &= "Total fees: " & CStr(MoneySummary(False, 0)) & vbCrLf
            MsgBox(summary)
            SetDefaults()

        End If
    End Sub

    Private Function ValidInputs() As Boolean
        Dim valid As Boolean = True
        Dim message As String
        Dim dayLimit As Integer
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
        'If CInt(BeginOdometerTextBox.Text) >= CInt(EndOdometerTextBox.Text) Then
        '    valid = False
        '    BeginOdometerTextBox.Clear()
        '    EndOdometerTextBox.Clear()
        '    message &= "Beginning odometer reading must be less than the odometer reading at return..."
        'End If
        If Not Integer.TryParse(DaysTextBox.Text, dayLimit) OrElse dayLimit < 1 OrElse dayLimit > 45 Then
            valid = False
            DaysTextBox.Clear()
            DaysTextBox.Focus()
            message &= "The allowable range of days for rental is 1 day minimum to 45 days maximum per rental"
        End If

        Dim beginningOdometer As Integer
        Dim endOdometer As Integer
        If Integer.TryParse(BeginOdometerTextBox.Text, beginningOdometer) AndAlso
                Integer.TryParse(EndOdometerTextBox.Text, endOdometer) Then
            If beginningOdometer >= endOdometer Then
                valid = False
                BeginOdometerTextBox.Clear()
                EndOdometerTextBox.Clear()
                message &= "Beggining odometer reading must be less than the odometer reading upon return."
            End If
        End If
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

    Private Function MeasureofDistance() As String

        Dim _measureOfDistance As String = " mi."
        If MilesradioButton.Checked = True Then
            _measureOfDistance = CStr(test()) & " mi."
            Return _measureOfDistance
        Else
            _measureOfDistance = CStr(test()) & " km."
            Return _measureOfDistance
        End If

    End Function

    Private Sub DaysTextBox_TextChanged(sender As Object, e As EventArgs) Handles DaysTextBox.TextChanged

    End Sub

    Private Sub RentalForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        SummaryButton.Enabled = False
    End Sub

    Function UserSummary(counting As Boolean) As Integer
        Static _UserSummary As Integer

        If counting = True Then
            _UserSummary += 1
        End If

        Return _UserSummary
    End Function

    Function MileSummary(counting As Boolean, miles As Integer) As Double
        Static _mileSummary As Double
        If counting = True Then
            _mileSummary += miles
        End If
        Return _mileSummary
    End Function
    Function MoneySummary(counting As Boolean, charge As Integer) As Double
        Static _moneySummary As Double
        If counting = True Then
            _moneySummary += charge
        End If
        Return _moneySummary
    End Function


End Class

