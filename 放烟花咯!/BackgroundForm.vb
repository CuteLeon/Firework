Public Class BackgroundForm
    Private Declare Function ReleaseCapture Lib "user32" () As Integer
    Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As VariantType) As Integer


#Region "模拟物理世界"
    Private Structure Firework
        Dim 种类 As Byte
        Dim 横坐标 As Integer
        Dim 竖坐标 As Integer
        Dim 宽度 As Integer
        Dim 垂直速度 As Integer
        Dim 帧标识 As Integer
        Dim 爆炸 As Boolean
    End Structure

    Private Const DefaultIntervalTime As Integer = 20     '默认更新时间间隔(单位：毫秒)
    Private Const Gravity As Single = 1.0F                        '默认重力加速度
    Dim VelocityRange() As Integer = {900, 930} '速度取值范围
    Dim Random As New Random                        '随机数生成器
    Dim FireworksWidth() As Int16 = {219, 200, 225}
    Dim FramesCount() As Int16 = {80, 30, 68}
    Dim FireworkBitmap As Bitmap
    Dim FireworkGraphics As Graphics
    Dim Fireworks As New ArrayList
#End Region

#Region "窗体和控件"

    Private Sub BackgroundForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Debug.Print("############################")
        Debug.Print("烟花 程序启动：" & My.Computer.Clock.LocalTime)
        Debug.Print("画面更新时间间隔： " & DefaultIntervalTime & " 毫秒")
        Debug.Print("模拟重力加速度G： " & Gravity.ToString("0.00 M/s^2"))
        Debug.Print("############################")

        UpdateFireworkTimer.Start()
    End Sub

    '值守引擎
    Private Sub UpdateFireworkTimer_Tick(sender As Object, e As EventArgs) Handles UpdateFireworkTimer.Tick
        If Random.NextDouble < 0.1 Then
            EmitSingleFirework()
        Else
            If Fireworks.Count = 0 Then Exit Sub
        End If

        FireworkBitmap = Nothing
        FireworkBitmap = My.Resources.FireworksResource.夜景_B
        FireworkGraphics = Graphics.FromImage(FireworkBitmap)

        Dim Index As Integer = 0, FireworksCount As Integer = Fireworks.Count
        Dim FireworkInst As Firework
        Do Until Index = FireworksCount
            UpdateFirework(Fireworks(Index), DefaultIntervalTime)
            FireworkInst = CType(Fireworks(Index), Firework)
            If FireworkInst.帧标识 = FramesCount(FireworkInst.种类) Then
                Fireworks(Index) = Nothing
                Fireworks.RemoveAt(Index)
                FireworksCount -= 1
                Continue Do
            End If

            FireworkGraphics.DrawImage(My.Resources.FireworksResource.ResourceManager.GetObject("烟花_" & FireworkInst.种类 & "_" & FireworkInst.帧标识), FireworkInst.横坐标, FireworkInst.竖坐标)

            Index += 1
        Loop

        FireworkGraphics.DrawImageUnscaled(My.Resources.FireworksResource.夜景_F, 0, 0)
        FireworkGraphics.DrawString(FireworksCount, Me.Font, Brushes.Red, 10, 10)
        FireworkGraphics.Dispose()
        Me.BackgroundImage = FireworkBitmap

        '强制回收内存
        GC.Collect()
    End Sub

    '鼠标拖动
    Private Sub BackgroundForm_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        ReleaseCapture()
        SendMessageA(Me.Handle, &HA1, 2, 0&)
    End Sub

#End Region

#Region "烟花相关函数"

    '发射烟花
    Private Sub EmitSingleFirework()
        Dim FireworkDemo As Firework
        With FireworkDemo
            .种类 = Random.Next(3)
            .横坐标 = Random.Next(Me.Width - 120 - FireworksWidth(.种类)) + 60
            .竖坐标 = Me.Height
            .帧标识 = 0
            .宽度 = FireworksWidth(.种类)
            .垂直速度 = Random.Next(VelocityRange.First, VelocityRange.Last)
        End With
        Fireworks.Add(FireworkDemo)

        Debug.Print(My.Computer.Clock.LocalTime & " !发射新的烟花： (烟花种类标识：" & FireworkDemo.种类 & "; 烟花坐标：(" & FireworkDemo.横坐标 & "," & FireworkDemo.竖坐标 & "); 垂直速度：" & FireworkDemo.垂直速度 & "; 烟花位图宽度：" & FireworkDemo.宽度 & ")")
    End Sub

    '更新烟花状态
    Private Sub UpdateFirework(ByRef Firework As Firework, ByVal IntervalTime As Integer)
        With Firework
            .竖坐标 -= Firework.垂直速度 * IntervalTime / 1000.0F
            .垂直速度 -= Gravity * IntervalTime * IIf(.爆炸, 0.05, 1)

            If .垂直速度 < 0 Then .爆炸 = True
            If .爆炸 Then
                .帧标识 += 1
            End If
        End With

        If Firework.帧标识 = 1 Then Debug.Print(My.Computer.Clock.LocalTime & " @@@ 烟花开始爆炸： (烟花种类标识：" & Firework.种类 & "; 烟花坐标：(" & Firework.横坐标 & "," & Firework.竖坐标 & "); 垂直速度：" & Firework.垂直速度 & ")")
    End Sub
#End Region

End Class