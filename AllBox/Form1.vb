Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices


Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllBox7.BindingTo(Me.TextBox1, "Text")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox(AllBox7.ValueDate)
    End Sub
End Class


Class AllBox
    Inherits TextBox

    Property Type As AllBoxTypes

    Sub New()
        Dim check As New CheckBox
        check.Top = 0
        check.Left = 0

        check.Text = "teste"

        Controls.Add(check)

        Dim f As New Form

        Cursor = Cursors.Default
        BorderStyle = Windows.Forms.BorderStyle.None
        BackColor = f.BackColor
        Multiline = True
        Height = 40
    End Sub
#Region "Value"
    ReadOnly Property [ValueString] As String
        Get
            If Type = AllBoxTypes.String Then
                Return Text
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueDate] As Date
        Get
            If Type = AllBoxTypes.Date And Not String.IsNullOrEmpty(Text) Then
                Return Date.Parse(Text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueDatetime] As DateTime
        Get
            If Type = AllBoxTypes.Datetime And Not String.IsNullOrEmpty(Text) Then
                Return Date.Parse(Text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueTimeSpan] As TimeSpan
        Get
            If Type = AllBoxTypes.TimeSpan And Not String.IsNullOrEmpty(Text) Then
                Return TimeSpan.Parse(Text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueSByte] As SByte
        Get
            If Type = AllBoxTypes.SByte And Not String.IsNullOrEmpty(Text) Then
                Return SByte.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueByte] As Byte
        Get
            If Type = AllBoxTypes.Byte And Not String.IsNullOrEmpty(Text) Then
                Return Byte.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueInt16] As Int16
        Get
            If Type = AllBoxTypes.Int16 And Not String.IsNullOrEmpty(Text) Then
                Return Int16.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueInt32] As Int32
        Get
            If Type = AllBoxTypes.Int32 And Not String.IsNullOrEmpty(Text) Then
                Return Int32.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueInt64] As Int64
        Get
            If Type = AllBoxTypes.Int64 And Not String.IsNullOrEmpty(Text) Then
                Return Int64.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueInteger] As Integer
        Get
            If Type = AllBoxTypes.Integer And Not String.IsNullOrEmpty(Text) Then
                Return Integer.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueDecimal] As Decimal
        Get
            If Type = AllBoxTypes.Decimal And Not String.IsNullOrEmpty(Text) Then
                Return Decimal.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueSingle] As Single
        Get
            If Type = AllBoxTypes.Single And Not String.IsNullOrEmpty(Text) Then
                Return Single.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueDouble] As Double
        Get
            If Type = AllBoxTypes.Double And Not String.IsNullOrEmpty(Text) Then
                Return Double.Parse(OnlyNumbersAndComma(Text))
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueBoolean] As Boolean
        Get
            If Type = AllBoxTypes.Boolean And Not String.IsNullOrEmpty(Text) Then
                Return Boolean.Parse(Text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueGuid] As Guid
        Get
            If Type = AllBoxTypes.Guid And Not String.IsNullOrEmpty(Text) Then
                Return Guid.Parse(Text)
            Else
                Return Nothing
            End If
        End Get
    End Property

    ReadOnly Property [ValueBinary] As Byte()
        Get
            If Type = AllBoxTypes.Binary And Not String.IsNullOrEmpty(Text) Then
                Using inputStream As New MemoryStream(Text)
                    Using outputStream As New MemoryStream
                        Using SR As New StreamReader(inputStream)
                            Do While SR.Peek > 0

                                Dim Line As String = SR.ReadLine()

                                For i As Integer = 0 To Line.Length - 1 Step 2

                                    Dim x2str As String = Line.Substring(i, 2)
                                    Dim b As Byte = Convert.ToByte(x2str, 16)

                                    outputStream.WriteByte(b)

                                Next

                            Loop
                        End Using

                        Return outputStream.ToArray
                    End Using
                End Using
            Else
                Return Nothing
            End If
        End Get
    End Property
#End Region

    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)

        Select Case Type
            Case AllBoxTypes.Currency, AllBoxTypes.Decimal, AllBoxTypes.Single, AllBoxTypes.Double, _
                 AllBoxTypes.Int16, AllBoxTypes.Int32, AllBoxTypes.Int64, AllBoxTypes.Integer, AllBoxTypes.SByte, AllBoxTypes.Byte

                If SelectionStart = Me.Text.Length Then
                    Dim SValue As String = OnlyNumbersAndComma(Text)

                    If Char.IsDigit(e.KeyChar) Then
                        SValue = SValue & e.KeyChar
                        e.Handled = True
                    ElseIf Asc(e.KeyChar) <> 8 Then
                        e.Handled = True
                    End If


                    If Char.IsDigit(e.KeyChar) Then
                        Select Case Type
                            Case AllBoxTypes.Currency
                                If Not Decimal.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Decimal
                                If Not Decimal.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Single
                                If Not Single.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Double
                                If Not Double.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Int16
                                If Not Int16.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Int32
                                If Not Int32.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Int64
                                If Not Int64.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Integer
                                If Not Integer.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.SByte
                                If Not SByte.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                            Case AllBoxTypes.Byte
                                If Not Byte.TryParse(SValue, Nothing) Then
                                    Exit Sub
                                End If
                        End Select
                    End If



                    If Not String.IsNullOrEmpty(SValue) Then
                        Dim DigitsCount As Integer = 0
                        Dim HasComma As Boolean = False
                        Dim p As Integer = SValue.IndexOf(",")

                        If p > 0 Then
                            DigitsCount = SValue.Length - p - 1
                            HasComma = True
                        End If



                        If Type = AllBoxTypes.Currency Then
                            Text = FormatCurrency(SValue, DigitsCount, TriState.UseDefault, TriState.UseDefault, TriState.True)
                        Else
                            Text = FormatNumber(SValue, DigitsCount, TriState.UseDefault, TriState.UseDefault, TriState.True)
                        End If

                        If Not HasComma And e.KeyChar = "," Then

                            Select Case Type
                                Case AllBoxTypes.Currency, AllBoxTypes.Decimal, AllBoxTypes.Single, AllBoxTypes.Double
                                    Text = Text & ","
                            End Select

                        End If

                        SelectionStart = Text.Length
                    End If
                End If
            Case AllBoxTypes.Date, AllBoxTypes.[TimeSpan], AllBoxTypes.Datetime
                If ((Not Char.IsNumber(e.KeyChar)) And (Not Char.IsControl(e.KeyChar) And (Not e.KeyChar = ":") And (Not e.KeyChar = "/") And (Not e.KeyChar = " "))) Then
                    e.Handled = True
                End If

                If Text.Length >= 19 Then
                    e.Handled = True
                End If
            Case AllBoxTypes.Guid
                Dim AllowChars As Char() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "A", "b", "B", "c", "C", "d", "D", "e", "E", "f", "F", "-"}

                If ((Not Char.IsControl(e.KeyChar)) And (Not AllowChars.Contains(e.KeyChar))) Then
                    e.Handled = True
                End If

                If Text.Length >= 36 Then
                    e.Handled = True
                End If

            Case AllBoxTypes.Binary
                Dim AllowChars As Char() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "A", "b", "B", "c", "C", "d", "D", "e", "E", "f", "F"}

                If ((Not Char.IsControl(e.KeyChar)) And (Not AllowChars.Contains(e.KeyChar))) Then
                    e.Handled = True
                End If
        End Select

    End Sub

    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)

        Select Case Type
            Case AllBoxTypes.Currency
                If String.IsNullOrEmpty(Text) Then
                    Text = RegionInfo.CurrentRegion.CurrencySymbol & " "
                End If
            Case AllBoxTypes.Decimal, AllBoxTypes.Single, AllBoxTypes.Double
            Case AllBoxTypes.Int16, AllBoxTypes.Int32, AllBoxTypes.Int64, AllBoxTypes.Integer, AllBoxTypes.SByte, AllBoxTypes.Byte
            Case AllBoxTypes.Date
            Case AllBoxTypes.Datetime
            Case AllBoxTypes.[TimeSpan]
            Case AllBoxTypes.Guid
            Case AllBoxTypes.Binary
        End Select
    End Sub

    Protected Overrides Sub OnValidating(e As System.ComponentModel.CancelEventArgs)
        MyBase.OnLeave(e)

        Select Case Type
            Case AllBoxTypes.Currency
                If Not String.IsNullOrEmpty(Me.Text) Then
                    Dim SValue As String = OnlyNumbersAndComma(Text)
                    Dim DigitCount As Integer = 2

                    Dim p As Integer = Text.IndexOf(",")

                    If p > 0 Then
                        DigitCount = Text.Length - p - 1
                    End If

                    If Not String.IsNullOrEmpty(SValue) Then
                        Dim S As String = FormatCurrency(SValue, DigitCount)
                        Me.Text = S
                    Else
                        Me.Text = String.Empty
                    End If
                End If
            Case AllBoxTypes.Decimal, AllBoxTypes.Single, AllBoxTypes.Double
                If Not String.IsNullOrEmpty(Me.Text) Then

                    Dim DigitsCount As Integer = 2
                    Dim p As Integer = Me.Text.IndexOf(",")
                    If p > 0 Then
                        DigitsCount = Me.Text.Length - p - 1
                    End If

                    Dim S As String = FormatNumber(Me.Text, DigitsCount, TriState.UseDefault, TriState.UseDefault, TriState.True)
                    Me.Text = S
                End If
            Case AllBoxTypes.Int16, AllBoxTypes.Int32, AllBoxTypes.Int64, AllBoxTypes.Integer, AllBoxTypes.SByte, AllBoxTypes.Byte
                If Not String.IsNullOrEmpty(Me.Text) Then
                    Dim S As String = FormatNumber(Me.Text, 0, TriState.UseDefault, TriState.UseDefault, TriState.True)
                    Me.Text = S
                End If
            Case AllBoxTypes.Date
                If Not String.IsNullOrEmpty(Me.Text) Then
                    Try
                        Dim S As String = FormatDateTime(Me.Text, DateFormat.ShortDate)
                        Me.Text = S
                    Catch ex As Exception
                        e.Cancel = True
                    End Try
                End If
            Case AllBoxTypes.Datetime
                Try
                    If Not String.IsNullOrEmpty(Me.Text) Then
                        Dim S As String = FormatDateTime(Me.Text, DateFormat.GeneralDate)
                        Me.Text = S
                    End If
                Catch ex As Exception
                    e.Cancel = True
                End Try
            Case AllBoxTypes.TimeSpan
                Try
                    If Not String.IsNullOrEmpty(Me.Text) Then
                        Dim S As String = FormatDateTime(Me.Text, DateFormat.LongTime)
                        Me.Text = S
                    End If
                Catch ex As Exception
                    e.Cancel = True
                End Try
            Case AllBoxTypes.Guid
            Case AllBoxTypes.Binary
        End Select

        TransferValueToBoundObject()
    End Sub

    Private _DataSource As Object
    Private _DataMember As String

    Friend Sub BindingTo(dataSource As Object, dataMember As String)
        _DataSource = dataSource
        _DataMember = dataMember
    End Sub

    Private Sub TransferValueToBoundObject()
        If _DataSource Is Nothing Or _DataMember Is Nothing Then Exit Sub

        Dim pi = _DataSource.GetType.GetProperties.Single(Function(x) x.Name = _DataMember)

        If pi.PropertyType = GetType(String) Then
            pi.SetValue(_DataSource, Text, Nothing)
        Else
            Select Case Type
                Case AllBoxTypes.String
                    pi.SetValue(_DataSource, ValueString, Nothing)
                Case AllBoxTypes.Currency
                    pi.SetValue(_DataSource, ValueDecimal, Nothing)
                Case AllBoxTypes.Decimal
                    pi.SetValue(_DataSource, ValueDecimal, Nothing)
                Case AllBoxTypes.Single
                    pi.SetValue(_DataSource, ValueSingle, Nothing)
                Case AllBoxTypes.Double
                    pi.SetValue(_DataSource, ValueDouble, Nothing)
                Case AllBoxTypes.Int16
                    pi.SetValue(_DataSource, ValueInt16, Nothing)
                Case AllBoxTypes.Int32
                    pi.SetValue(_DataSource, ValueInt32, Nothing)
                Case AllBoxTypes.Int64
                    pi.SetValue(_DataSource, ValueInt64, Nothing)
                Case AllBoxTypes.Integer
                    pi.SetValue(_DataSource, ValueInteger, Nothing)
                Case AllBoxTypes.SByte
                    pi.SetValue(_DataSource, ValueSByte, Nothing)
                Case AllBoxTypes.Byte
                    pi.SetValue(_DataSource, ValueByte, Nothing)
                Case AllBoxTypes.Date
                    pi.SetValue(_DataSource, ValueDate, Nothing)
                Case AllBoxTypes.Datetime
                    pi.SetValue(_DataSource, ValueDatetime, Nothing)
                Case AllBoxTypes.TimeSpan
                    pi.SetValue(_DataSource, ValueTimeSpan, Nothing)
                Case AllBoxTypes.Guid
                    pi.SetValue(_DataSource, ValueGuid, Nothing)
                Case AllBoxTypes.Binary
                    pi.SetValue(_DataSource, ValueBinary, Nothing)
            End Select
        End If

    End Sub

    Private Function OnlyNumbersAndComma(S As String) As String
        Dim Ret As String = String.Empty
        For Each c In S
            If Char.IsDigit(c) Or c = "," Then
                Ret += c
            End If
        Next
        Return Ret
    End Function

    Private Function OnlyNumbers(S As String) As String
        Dim Ret As String = String.Empty
        For Each c In S
            If Char.IsDigit(c) Then
                Ret += c
            End If
        Next
        Return Ret
    End Function

End Class

Enum AllBoxTypes
    [String]
    [Date]
    [Datetime]
    [TimeSpan]
    [SByte]
    [Byte]
    [Int16]
    [Int32]
    [Int64]
    [Integer]
    [Decimal]
    [Single]
    [Double]
    [Boolean]
    [Currency]
    [Guid]
    [Binary]
End Enum



Public Class KeyCodeToAscii
    <DllImport("User32.dll")> _
    Private Shared Function ToAscii(ByVal uVirtKey As Integer, ByVal uScanCode As Integer, ByVal lpbKeyState As Byte(), ByVal lpChar As Byte(), ByVal uFlags As Integer) As Integer
    End Function

    <DllImport("User32.dll")> _
    Private Shared Function GetKeyboardState(ByVal pbKeyState As Byte()) As Integer
    End Function

    Public Shared Function GetAsciiCharacter(ByVal uVirtKey As Integer) As Char
        Dim lpKeyState As Byte() = New Byte(255) {}
        GetKeyboardState(lpKeyState)
        Dim lpChar As Byte() = New Byte(1) {}
        If ToAscii(uVirtKey, 0, lpKeyState, lpChar, 0) = 1 Then
            Return Convert.ToChar((lpChar(0)))
        Else
            Return New Char()
        End If
    End Function
End Class