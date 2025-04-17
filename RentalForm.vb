'Jacob Horsley
'RCET 0265
'Spring 2025
'URL: 



'[ ] Remove invalid data from textbox 
'[ ] Set tab order
'[ ] Validate odometer readings
'[ ] Validate number of days
'[ ] single message box to display an improper input

Option Explicit On
Option Strict On
Option Compare Binary
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

    'Private Sub ValidateInput()
    '    Dim nameValid As Boolean = Not String.IsNullOrWhiteSpace(NameTextBox.Text)
    '    Dim addressValid As Boolean = Not String.IsNullOrWhiteSpace(AddressTextBox.Text)
    '    Dim cityValid As Boolean = Not String.IsNullOrWhiteSpace(CityTextBox.Text)
    '    Dim stateValid As Boolean = Not String.IsNullOrWhiteSpace(StateTextBox.Text)
    '    Dim zipValid As Boolean = Not String.IsNullOrWhiteSpace(ZipCodeTextBox.Text)
    '    Dim beginOdom As Boolean = Not String.IsNullOrWhiteSpace(BeginOdometerTextBox.Text)
    '    Dim endOdom As Boolean = Not String.IsNullOrWhiteSpace(EndOdometerTextBox.Text)
    '    Dim days As Boolean = Not String.IsNullOrWhiteSpace(DaysTextBox.Text)
    '    Dim totalMi As Boolean = Not String.IsNullOrWhiteSpace(TotalMilesTextBox.Text)
    '    Dim mileageCharge As Boolean = Not String.IsNullOrWhiteSpace(TotalMilesTextBox.Text)
    '    Dim discount As Boolean = Not String.IsNullOrWhiteSpace(TotalDiscountTextBox.Text)

    '    Dim allValid As Boolean = nameValid And addressValid And cityValid And stateValid And zipValid And
    '    beginOdom And endOdom And days And totalMi And mileageCharge And discount
    'End Sub
    'Private Sub NameTextBox_TextChanged(sender As Object, e As EventArgs) Handles NameTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub AddressTextBox_TextChanged(sender As Object, e As EventArgs) Handles AddressTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub CityTextBox_TextChanged(sender As Object, e As EventArgs) Handles CityTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub StateTextBox_TextChanged(sender As Object, e As EventArgs) Handles StateTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub ZipCodeTextBox_TextChanged(sender As Object, e As EventArgs) Handles ZipCodeTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub BeginOdometerTextBox_TextChanged(sender As Object, e As EventArgs) Handles BeginOdometerTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub EndOdometerTextBox_TextChanged(sender As Object, e As EventArgs) Handles EndOdometerTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub DaysTextBox_TextChanged(sender As Object, e As EventArgs) Handles DaysTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub TotalMilesTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalMilesTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub MileageChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles MileageChargeTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub DayChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles DayChargeTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub TotalDiscountTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalDiscountTextBox.TextChanged
    '    ValidateInput()
    'End Sub

    'Private Sub TotalChargeTextBox_TextChanged(sender As Object, e As EventArgs) Handles TotalChargeTextBox.TextChanged
    '    ValidateInput()
    'End Sub



    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click

        DayChargeTextBox.Text = dailyDollars.ToString
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
        Me.Close()
    End Sub

    Private Sub NameTextBox_TextChanged(sender As Object, e As EventArgs) Handles NameTextBox.TextChanged

    End Sub

    Private Sub NameTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles NameTextBox.KeyPress, CityTextBox.KeyPress, StateTextBox.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsWhiteSpace(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub ZipCodeTextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ZipCodeTextBox.KeyPress, BeginOdometerTextBox.KeyPress, EndOdometerTextBox.KeyPress, DaysTextBox.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub






    'Function LettersOnly(input As String) As Boolean
    '    For Each c As Char In input
    '        If Not Char.IsLetter(c) AndAlso Not Char.IsWhiteSpace(c) Then
    '            Return False
    '        End If
    '    Next
    '    Return True
    'End Function

    'Function NumbersOnly(input As String) As Boolean ' More optional functions
    '    Return NumbersOnly(input)
    'End Function

End Class
