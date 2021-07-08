Imports System.Threading
Imports System.IO.Ports
Imports System.IO
Imports System.Text


Public Class winserial

    Public Event OnPortBeforeConnected(ss As SerialPort)
    Public Event OnConnected(sender As Object)
    Public Event OnDisconnect(sender As Object)
    Public Event OnErrorFound(sender As Object, errorcode As Integer, reason As String)
    Public Event OnByteRead(sender As Object, bytes As Byte())
    Public Event OnStringRead(sender As Object, str As String)
    Public Event OnStartReading(sender As Object)

    Private _readmanual As Boolean = False
    Public Property ReadManually() As Boolean
        Get
            Return _readmanual
        End Get
        Set(ByVal value As Boolean)
            _readmanual = value
        End Set
    End Property

    Private _isconnected As Boolean = False
    Public ReadOnly Property IsConnected() As Boolean
        Get
            Return _isconnected
        End Get
    End Property

    Private _lasterror As String
    Public ReadOnly Property LastError() As String
        Get
            Return _lasterror
        End Get
    End Property

    Private _connectedCom As String
    Public Property ComSerial() As String
        Get
            Return _connectedCom
        End Get
        Set(ByVal value As String)
            _connectedCom = value
        End Set
    End Property

    Private _baudrate As Integer
    Public Property BaudRate() As Integer
        Get
            Return _baudrate
        End Get
        Set(ByVal value As Integer)
            _baudrate = value
        End Set
    End Property

    Private _stopbits As StopBits
    Public Property StopBit() As StopBits
        Get
            Return _stopbits
        End Get
        Set(ByVal value As StopBits)
            _stopbits = value
        End Set
    End Property

    Private _parity As Parity
    Public Property Paritys() As Parity
        Get
            Return _parity
        End Get
        Set(ByVal value As Parity)
            _parity = value
        End Set
    End Property

    Private _bytesized As Integer
    Public Property Bytesized() As Integer
        Get
            Return _bytesized
        End Get
        Set(ByVal value As Integer)
            _bytesized = value
        End Set
    End Property

    Private _writetimeout As Integer
    Public Property WriteTimeout() As Integer
        Get
            Return _writetimeout
        End Get
        Set(ByVal value As Integer)
            _writetimeout = value
        End Set
    End Property

    Private _handshake As Handshake
    Public Property Handshakes() As Handshake
        Get
            Return _handshake
        End Get
        Set(ByVal value As Handshake)
            _handshake = value
        End Set
    End Property

    Private _readbuffersize As Integer
    Public Property NewProperty() As Integer
        Get
            Return _readbuffersize
        End Get
        Set(ByVal value As Integer)
            _readbuffersize = value
        End Set
    End Property


    Private _serialports As SerialPort
    Private _cancelationtokensource As CancellationTokenSource
    Private _startreading As Boolean

    Sub New()
        MyBase.New
        _isconnected = False
        _lasterror = ""
        _connectedCom = "COM1"
        _baudrate = 9600
        _stopbits = StopBits.One
        _parity = Parity.None
        _bytesized = 8
        _handshake = Handshake.RequestToSend
        _writetimeout = 1000
    End Sub

    Public Function Connect() As Boolean
        _lasterror = ""
        If _isconnected Then Return True
        If _serialports Is Nothing Then _serialports = New SerialPort With {
               .BaudRate = _baudrate,
               .DataBits = _bytesized,
               .Parity = _parity,
               .StopBits = _stopbits,
               .PortName = _connectedCom,
               .WriteTimeout = _writetimeout,
                .Handshake = _handshake
            }

        If _readbuffersize = 0 Then _readbuffersize = 4096

        RaiseEvent OnPortBeforeConnected(_serialports)

        _lasterror = ""
        Try
            _serialports.Open()
        Catch ex As Exception
            _lasterror = ex.Message
        End Try
        _isconnected = _serialports.IsOpen
        If _lasterror.Length > 0 Then
            RaiseEvent OnErrorFound(Me, 1, _lasterror)
        Else
            RaiseEvent OnConnected(Me)
        End If

        Return _isconnected

    End Function

    Public Sub StartRead()
        If _startreading Then Return
        Dim iscontinue As Boolean = _serialports IsNot Nothing
        If iscontinue Then iscontinue = _serialports.IsOpen

        If Not iscontinue Then
            _lasterror = "Serialport not open"
            RaiseEvent OnErrorFound(Me, 2, _lasterror)
            Return
        End If

        If _cancelationtokensource Is Nothing Then
            _cancelationtokensource = New CancellationTokenSource()
        ElseIf _cancelationtokensource.IsCancellationRequested Then
            _cancelationtokensource = New CancellationTokenSource()
        End If

        _startreading = True
        RaiseEvent OnStartReading(Me)

        Task.Run(Sub()
                     Dim sdatabyte(_readbuffersize) As Byte
                     While True
                         If _cancelationtokensource.IsCancellationRequested Then Exit While
                         If _readmanual Then
                             Array.Clear(sdatabyte, 0, sdatabyte.Length)
                             Dim i = _serialports.Read(sdatabyte, 0, sdatabyte.Length)
                             If i > 0 Then
                                 RaiseEvent OnByteRead(Me, sdatabyte)
                             End If
                         Else
                             Dim s As String
                             s = _serialports.ReadLine
                             RaiseEvent OnStringRead(Me, s)
                         End If
                     End While
                     _startreading = False
                 End Sub)
    End Sub

    Public Sub StopReading()
        If _serialports.IsOpen Then
            If _startreading Then
                _cancelationtokensource.Cancel()
                Thread.Sleep(TimeSpan.FromMilliseconds(50))
            End If
        End If
    End Sub

    Public Sub WriteLine(s As String)
        If _isconnected = False Then Return
        If _startreading Then Return
        Task.Run(Sub()
                     _serialports.WriteLine(s)
                     Dim ss As String = _serialports.ReadLine()
                     If ss.Length > 0 Then
                         RaiseEvent OnStringRead(Me, ss)
                     End If
                 End Sub)
    End Sub

    Public Sub CloseConnection()
        If _serialports Is Nothing Then Return
        StopReading()
        _serialports.Close()
        _isconnected = False
        RaiseEvent OnDisconnect(Me)
    End Sub

End Class
