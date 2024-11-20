Imports Ivi.Visa.Interop
Imports Microsoft.Office.Interop

Public Class Form2
    Dim i As Integer 'loop 変数add

    Dim myPos As Integer '4278ACdとDのデータ値区切り","位置
    Dim measdata As String 'GPIB測定データ
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '測定ボタン
        '測定ボタン
        MsgBox("■書き込みセルをクリックしましたか？" & vbCrLf &
               "■測定端子にチップコンデンサをセットしましたか？" & vbCrLf &
               "OKで測定します。")

        richi = xlsApplication.ActiveCell.Column
        cichi = xlsApplication.ActiveCell.Row
        TextBox4.Text = richi.ToString
        TextBox5.Text = cichi.ToString

        Dim ioMgr As Ivi.Visa.Interop.ResourceManager
        Dim instrument As Ivi.Visa.Interop.FormattedIO488
        Dim idn As String

        '***************************************************************
        'High Accuracy Mode Auto Setting routine 
        'ADD 2022NOV17
        Dim HI_CRead As String
        Dim HI_CSet As String
        HI_CRead = 0
        HI_CSet = 0

        HI_CRead = xlsApplication.Cells(8, 24).Text '"150pF"

        Dim pfsakujo As Integer
        pfsakujo = Len(HI_CRead)
        Select Case pfsakujo
            Case 5
                HI_CSet = Trim(Mid(HI_CRead, 1, 3)) '"100-300pF:ex)300pF" -> "300"
            Case 6
                HI_CSet = Trim(Mid(HI_CRead, 1, 4)) '"1000pF" -> "1000"
            Case Else
                MsgBox("EXCEL CELL READING ERROR" & vbCrLf &
                       "Check Excel CAPA Setting CELL Value" & vbCrLf &
                       "Cells(8,24)")
        End Select
        '***************************************************************
        ' MsgBox(HI_CSet)


        'Dim GPIBDAT As String
        'Dim GPIBAD As String

        'GPIBDAT = "17"
        'GPIBDAT = Trim(TextBox1.Text)

        '****************************************
        'GPIBAD = "GPIB0::" & GPIBDAT & "::INSTR"  '現行SUB-STD用4278A
        'GPIBAD = "GPIB1::" & GPIBDAT & "::INSTR" 'for TEST
        'GPIBAD = "GPIB2::" & GPIBDAT & "::INSTR" 'テスト用4278A
        '****************************************

        'ioMgr = New Ivi.Visa.Interop.ResourceManager
        ioMgr = New ResourceManager
        'instrument = New Ivi.Visa.Interop.FormattedIO488
        instrument = New FormattedIO488

        instrument.IO = ioMgr.Open(GPIBAD)

        '4278A設定 *******************************************************************************
        '
        'Debug.WriteLine(CAPA)

        instrument.WriteString("MPAR1")     '測定ﾊﾟﾗﾒｰﾀ(Cp-D)
        instrument.WriteString("FREQ2")     'FREQUENCY (1MHz)
        instrument.WriteString("OSC=1.0")   'OSC LEVEL(1.0V)

        instrument.WriteString("HIAC1") 'SM-11S96 Line 4278A also  setting ok!

        '***************************************************
        instrument.WriteString("RC=" & HI_CSet & "E-12")     'set range for HI_ACCURACY
        '***************************************************

        ' instrument.WriteString("RB0")       '測定ﾚﾝｼﾞ（RB:AUTO range at 1MHz)
        instrument.WriteString("ITIM3")     '積分時間（LONG)
        instrument.WriteString("DTIM=0")    'delaytime 0ms

        instrument.WriteString("AVE=32")    'avaraging 32

        instrument.WriteString("TRIG1")     'internal trigger mode
        instrument.WriteString("CABL0")     'cable length 0m

        '*****************************************************************************************

        instrument.WriteString("DATA?")
        idn = instrument.ReadString()

        'MsgBox(idn)  '
        measdata = Trim(idn) '    
        '                                123456789012345678    9
        'for check                      ":DATA +15.0690E+03" & vbLf
        Dim lmeasdata As Integer
        lmeasdata = Len(measdata)

        '*****************************************************
        'デバッグモード設定用　
        superslim = 0      ' Set 1:7555MultiMeter, Set 0:4278A
        '*****************************************************
        '                                123456789012345678    9
        'for check                      ":DATA +15.0690E+03" & vbLf

        Select Case superslim
            Case 1
                sngC = Mid(measdata, 8, 18) '"15.0690E+03"
                TextBox2.Text = sngC
                sngD = 0.001
                TextBox3.Text = sngD
            Case 0
                myPos = InStr(1, measdata, ",", vbTextCompare)　'データ区切り位置
                sngC = Mid(measdata, 1, myPos - 1) 'Cd値抜き取り
                sngC = sngC * 1000000000000.0# 'pF単位に変換処理
                TextBox2.Text = sngC
                sngD = Mid(measdata, myPos + 1) 'D値抜き取り
                sngD = sngD * 100.0#
                TextBox3.Text = sngD
            Case Else
        End Select

        xlsRange = xlsWorkSheet.Cells(cichi, richi)
        xlsRange.Value = sngC
        xlsRange = xlsWorkSheet.Cells(cichi + 1, richi)
        xlsRange.Value = sngD / 100 'for Excel cell setting is % so that need real D value


        'tooltipにて説明に変更
        'MsgBox("測定継続→次のセルクリック→コンデンサ入替→測定” & vbCrLf & vbCrLf & "測定終了→「名前を付けて保存」")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '名前を付けて保存処理ボタン
        Me.Hide()
        Form1.Show()

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        ToolTip1.ShowAlways = True

        ToolTip1.SetToolTip(Button1, "測定終了時は" & vbCrLf & "  [名前を付けて保存処理]ボタン  " & vbCrLf &
                                     "を押してください。")
        ToolTip1.SetToolTip(Button2, "データを入れるセルをクリックしてから " & vbCrLf &
                                     "測定ボタンを押して下さい。")

        '****************************************************************************
        'Automatically Get Visa Address & Visa Alias if any
        Dim VisaCount As Integer
        VisaCount = 0

        Dim RM = New Ivi.Visa.Interop.ResourceManager
        VisaAdds = RM.FindRsrc("GPIB?*INSTR")
        GPIBAD = ""

        For i = 0 To UBound(VisaAdds)
            RM.ParseRsrcEx(VisaAdds(i), plnterfaceType, plnterfaceNumber,
                           pSessionType, pUnaliasedExpandedResourceName, pAliaslfExists)
        Next
        Me.TextBox1.Text = VisaAdds(0)
        'Me.TextBox6.Text = pAliaslfExists
        Me.TextBox6.Text = VisaCount.tostring


        'GPIBのVISAアドレスをグローバル変数GPIBADに設定
        'EX)設定VISAアドレス： "GPIB2::17::INSTR"
        GPIBAD = VisaAdds(0) '最終設定Visa Address
        'GPIBAD = VisaAdds(1) '前の設定Visa Address

        RM = Nothing

    End Sub
End Class



