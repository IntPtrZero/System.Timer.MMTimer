Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.ComponentModel.Design
Imports System.Drawing

<DefaultProperty("Interval"), DefaultEvent("Tick"), Description("高精度的多媒体时钟组件"), ToolboxBitmap(GetType(System.Windows.Forms.Timer))> _
Public Class MMTimer : Inherits Component
    Implements ISupportInitialize

#Region "pInvoke"
    <DllImport("winmm.dll")>
    Private Shared Function timeGetDevCaps(ByRef LPTIMECAPS As TimeCaps, ByVal cbSize As Integer) As Integer
    End Function

    <DllImport("winmm.dll")>
    Private Shared Function timeBeginPeriod(ByVal uPeriod As UInteger) As Integer
    End Function

    <DllImport("winmm.dll")>
    Private Shared Function timeSetEvent(ByVal uDelay As UInteger, ByVal uResolution As UInteger, ByVal lpTimeProc As TimerProc, ByRef lpdwuser As UInteger, ByVal fuEvebt As UInteger) As UInteger
    End Function

    <DllImport("winmm.dll")>
    Private Shared Function timeKillEvent(ByVal uTimerID As UInteger) As Integer
    End Function

    <DllImport("winmm.dll")>
    Private Shared Function timeEndPeriod(ByVal uPeriod As UInteger) As Integer
    End Function

    Private Delegate Sub TimerProc(ByVal uId As UInteger, ByVal uMsg As UInteger, ByVal dwUser As UInteger, ByVal dw1 As UInteger, ByVal dw2 As UInteger)

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure TimeCaps
        Public PeriodMin As UInteger
        Public PeriodMax As UInteger
    End Structure
#End Region

    Private Shared mCaps As TimeCaps
    Private mResolution As Integer
    Private mInterval As Integer
    Private mEnabled As Boolean
    Private mSynchronizingObject As ISynchronizeInvoke
    Private mTimerProc As TimerProc
    Private mTimerID As UInteger
    Private mInitialzing As Boolean

    <EditorBrowsable(EditorBrowsableState.Never)> _
    Public Delegate Sub TickEventHandler(sender As Object, e As EventArgs)
    ''' <summary>
    ''' 计时器到期时引发事件
    ''' </summary>
    ''' <remarks></remarks>
    Public Event Tick As TickEventHandler
    Private mOnTickHandler As TickEventHandler

    Shared Sub New()
        timeGetDevCaps(mCaps, Marshal.SizeOf(mCaps))
    End Sub

    Public Sub New()
        MyBase.New()
        mTimerID = 0
        mInterval = 100
        mEnabled = False
        mInitialzing = False
        mResolution = mCaps.PeriodMin
        mTimerProc = New TimerProc(AddressOf TimerCallBack)
        mOnTickHandler = New TickEventHandler(AddressOf OnTick)
    End Sub

    Public Sub New(ByVal itv As Integer)
        Me.New()
        If itv >= mCaps.PeriodMin And itv <= mCaps.PeriodMax Then
            mInterval = itv
        End If
    End Sub

    ''' <summary>
    ''' 获取或设置计时器是否正在运行
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DefaultValue(False), Browsable(True), Category("Behavior"), Description("启用Tick事件生成")> _
    Public Property Enabled As Boolean
        Get
            Return mEnabled
        End Get
        Set(value As Boolean)
            mEnabled = value
            If Not mInitialzing Then
                ResetTimer()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Tick事件的频率(以毫秒为单位)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DefaultValue(100), Browsable(True), Category("Behavior"), Description("Tick事件的频率(以毫秒为单位)")> _
    Public Property Interval As Integer
        Get
            Return mInterval
        End Get
        Set(value As Integer)
            If value >= mCaps.PeriodMin And value <= mCaps.PeriodMax Then
                mInterval = value
                If Not mInitialzing Then
                    ResetTimer()
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' 计时器的最小精度(以毫秒为单位)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(False)> _
    Public Shared ReadOnly Property PeriodMin As UInteger
        Get
            Return mCaps.PeriodMin
        End Get
    End Property

    ''' <summary>
    ''' 计时器的最大精度(以毫秒为单位)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(False)> _
    Public Shared ReadOnly Property PeriodMax As UInteger
        Get
            Return mCaps.PeriodMax
        End Get
    End Property

    ''' <summary>
    ''' 时钟的精度(以毫秒为单位)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Browsable(True), Description("时钟的精度(以毫秒为单位)")> _
    Public Property Resolution As UInteger
        Get
            Return mResolution
        End Get
        Set(value As UInteger)
            If value >= mCaps.PeriodMin And value <= mCaps.PeriodMax Then
                mResolution = value
                If Not mInitialzing Then
                    ResetTimer()
                End If
            End If
        End Set
    End Property

    <DefaultValue(vbNull), Browsable(False), EditorBrowsable(EditorBrowsableState.Advanced)> _
    Public Property SynchronizingObject As ISynchronizeInvoke
        Get
            If (mSynchronizingObject Is Nothing) And DesignMode Then
                Dim host As IDesignerHost = GetService(GetType(IDesignerHost))
                If host IsNot Nothing Then
                    Dim baseComponent = host.RootComponent
                    If baseComponent IsNot Nothing AndAlso TypeOf (baseComponent) Is ISynchronizeInvoke Then
                        mSynchronizingObject = baseComponent
                    End If
                End If
            End If
            Return mSynchronizingObject
        End Get
        Set(value As ISynchronizeInvoke)
            mSynchronizingObject = value
        End Set
    End Property

    ''' <summary>
    ''' 启动计时器
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Start()
        Enabled = True
    End Sub

    ''' <summary>
    ''' 停止计时器
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub [Stop]()
        Enabled = False
    End Sub

    Private Sub StartTimer()
        mTimerID = timeSetEvent(mInterval, mResolution, mTimerProc, 0, 1)
    End Sub

    Private Sub StopTimer()
        If mTimerID Then
            timeKillEvent(mTimerID)
            mTimerID = 0
        End If
    End Sub

    Private Sub ResetTimer()
        StopTimer()
        If mEnabled Then
            StartTimer()
        End If
    End Sub

    ''' <summary>
    ''' 释放由 System.ComponentModel.Component 使用的所有资源
    ''' </summary>
    ''' <remarks></remarks>
    Public Overloads Sub Dispose()
        StopTimer()
        MyBase.Dispose()
    End Sub

    Private Sub TimerCallBack(ByVal uId As UInteger, ByVal uMsg As UInteger, ByVal dwUser As UInteger, ByVal dw1 As UInteger, ByVal dw2 As UInteger)
        If (mSynchronizingObject IsNot Nothing) AndAlso mSynchronizingObject.InvokeRequired Then
            mSynchronizingObject.Invoke(mOnTickHandler, New Object() {Me, Nothing})
        Else
            OnTick(Me, Nothing)
        End If
    End Sub

    Private Sub OnTick(sender As Object, e As EventArgs)
        RaiseEvent Tick(sender, e)
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never)> _
    Public Sub BeginInit() Implements ISupportInitialize.BeginInit
        mInitialzing = True
    End Sub

    <EditorBrowsable(EditorBrowsableState.Never)> _
    Public Sub EndInit() Implements ISupportInitialize.EndInit
        mInitialzing = False
        If mEnabled Then
            StartTimer()
        End If
    End Sub

End Class
