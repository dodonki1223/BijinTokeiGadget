Option Explicit On

Imports BijinTokeiGadget.BijinTokeiDefinition
Imports System.IO
Imports System.Windows.Shell
Imports System.ComponentModel
Imports System.Windows.Interop
Imports System.Runtime.InteropServices
Imports System.Collections.Generic

''' <summary>美人時計画像を表示するウィンドウの機能を提供します</summary>
''' <remarks>
'''  'ToDo:ダウンロード処理に失敗した時、トレイアイコンのマウスオーバー時のテキストがダウンロード中のものから
'''        変更されない不具合あり。修正方法がわからないのでとりあえず放置する……。
''' </remarks>
Class MainWindow
    Inherits System.Windows.Window

#Region "変数"

    ''' <summary>画像切り替え用タイマー</summary>
    ''' <remarks></remarks>
    Private _ChangeImageTimer As System.Windows.Threading.DispatcherTimer

    ''' <summary>トレイアイコンコントロール</summary>
    ''' <remarks></remarks>
    Private _NotifyIcon As System.Windows.Forms.NotifyIcon

    ''' <summary>ファイルダウンロード用クラス</summary>
    ''' <remarks></remarks>
    Private _DownloadForNet As DownloadForNet

    ''' <summary>美人時計画像がダウンロード中か</summary>
    ''' <remarks></remarks>
    Private _IsDownloadingForBijinTokeiImages As Boolean = False

    ''' <summary>画像ダウンロード数</summary>
    ''' <remarks></remarks>
    Private _ImageDownloadCount As ImageDownloadCount

    ''' <summary>表示時計ID</summary>
    ''' <remarks></remarks>
    Private _TargetTokeiId As Integer

#End Region

#Region "プロパティ"

    ''' <summary>美人時計画像サイズ</summary>
    ''' <returns></returns>
    Public ReadOnly Property BijinTokeiImageSize As BijinTokeiImageInfo

        Get

            Dim ImageSize As New BijinTokeiImageInfo

            ImageSize.Width = Me.BijinTokei.TargetBijinTokei.ImageWidth
            ImageSize.Height = Me.BijinTokei.TargetBijinTokei.ImageHeight
            ImageSize.MinimumSizeRate = Me.BijinTokei.TargetBijinTokei.MinimumSizeRate

            Return ImageSize

        End Get

    End Property

    ''' <summary>美人時計ガジェットのウィンドウの幅・高さの修正率</summary>
    ''' <returns></returns>
    Public ReadOnly Property WindowFixRate As Double

        Get

            Return Me.BijinTokei.TargetBijinTokei.ImageWidth / Me.BijinTokei.TargetBijinTokei.ImageHeight

        End Get

    End Property

    ''' <summary>美人時計画像ダウンロード数</summary>
    ''' <remarks>１回のダウンロード時にダウンロードする画像数</remarks>
    Public ReadOnly Property ImageDownloadCount As ImageDownloadCount

        Get

            Return _ImageDownloadCount

        End Get

    End Property

    ''' <summary>表示時計情報</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property BijinTokei As BijinTokeiInfo

        Get

            Return New BijinTokeiInfo(_TargetTokeiId)

        End Get

    End Property

    ''' <summary>美人時計画像がダウンロード可能か</summary>
    ''' <returns>
    '''   ネットワークが使用可能で無い時、美人時計画像ダウンロード中の時
    ''' </returns>
    Public ReadOnly Property IsAvailableDownloadBijinTokei As Boolean

        Get

            'ネットワークが使用可能な時
            If System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable() Then

                '美人時計画像がダウンロード中の時
                If Me.IsDownloadingForBijinTokeiImages Then

                    Return False

                Else

                    Return True

                End If

            Else

                Return False

            End If

        End Get

    End Property

    ''' <summary>美人時計画像のダウンロード中か</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property IsDownloadingForBijinTokeiImages As Boolean

        Get

            Return _IsDownloadingForBijinTokeiImages

        End Get

    End Property

#End Region

#Region "DLL関数"

    ''' <summary>指定されたウィンドウの表示状態、および通常表示のとき、最小化されたとき、最大化されたときの位置を返します。</summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="lpwndpl">位置データ</param>
    ''' <returns>
    '''   関数が成功すると、0 以外の値が返ります。
    '''   関数が失敗すると、0 が返ります。拡張エラー情報を取得するには、 関数を使います。
    ''' </returns>
    ''' <remarks>
    '''   この関数が取得する WINDOWPLACEMENT 構造体の flags メンバは、常に 0 です。
    '''   hWnd パラメータで指定したウィンドウが最大化されている場合、
    '''   showCmd メンバが SW_SHOWMAXIMIZED に設定されます。
    '''   ウィンドウが最小化されている場合は、showCmd メンバが SW_SHOWMINIMIZED に設定されます。
    '''   それ以外の場合は、SW_SHOWNORMAL に設定されます。
    '''   WINDOWPLACEMENT 構造体の length メンバは、sizeof(WINDOWPLACEMENT) に設定されていなければなりません。
    '''   このメンバが正しく設定されていないと、0（FALSE）が返ります。
    '''   ウィンドウの位置座標の正しい扱い方の詳細については、 構造体の説明を参照してください。
    ''' </remarks>
    <DllImport("user32.dll")> _
    Private Shared Function GetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

    ''' <summary>指定されたウィンドウの表示状態を設定し、そのウィンドウの通常表示のとき、最小化されたとき、および最大化されたときの位置を設定します</summary>
    ''' <param name="hWnd">ウィンドウのハンドル</param>
    ''' <param name="lpwndpl">位置データ</param>
    ''' <returns>
    '''   関数が成功すると、0 以外の値が返ります。
    '''   関数が失敗すると、0 が返ります。拡張エラー情報を取得するには、 関数を使います。
    ''' </returns>
    ''' <remarks>
    '''   WINDOWPLACEMENT 構造体で指定された情報を適用するとウィンドウが完全に画面の外に出てしまう場合は、
    '''   ウィンドウが画面に現れるように座標が自動調整されます。
    '''   この調整では、画面の解像度の変更や複数モニタの構成も考慮されます。
    '''   WINDOWPLACEMENT 構造体の length メンバは、sizeof(WINDOWPLACEMENT) に設定されていなければなりません。
    '''   このメンバが正しく設定されていないと、0（FALSE）が返ります。
    '''   ウィンドウの位置座標の正しい扱い方の詳細については、 構造体の説明を参照してください。
    ''' </remarks>
    <DllImport("user32.dll")> _
    Private Shared Function SetWindowPlacement(ByVal hWnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Boolean
    End Function

#End Region

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks></remarks>
    Public Sub New()

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。

        'トレイアイコン設定
        Call _ToSetupTrayIcon()

    End Sub

#End Region

#Region "イベント"

#Region "メインウィンドウ（MainWindow）"

    ''' <summary>メインウィンドウのLoadイベント</summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">Loadedイベント</param>
    ''' <remarks></remarks>
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        '美人時計画像がダウンロードが可能でない時	、バルーンウィンドウにネットワーク使用不可メッセージを表示
        If Not Me.IsAvailableDownloadBijinTokei Then Call _SetBalloonWindowText(cMessage.NetworkNotAvailable, BalloonWindowDisplayTime.NetworkNotAvailable)

        'コントロール設定
        Call _ToSetupControls()

        'トレイアイコンのテキスト設定
        Call _SetTrayIconText()

    End Sub

    ''' <summary>メインウィンドウのStateChangedイベント</summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">StateChangedイベント</param>
    ''' <remarks>最小化・最大化を無効化する（一瞬最大化・最小化処理がされてしまう……）</remarks>
    Private Sub MainWindow_StateChanged(sender As Object, e As EventArgs) Handles Me.StateChanged

        'ウィンドウの状態が「最小化」 または ウィンドウの状態が「最大化」の時
        If Me.WindowState = WindowState.Minimized OrElse Me.WindowState = Windows.WindowState.Maximized Then

            'ウィンドウの状態を「通常にする」
            Me.WindowState = Windows.WindowState.Normal

        End If

    End Sub

    ''' <summary>メインウィンドウのMouseEnterイベント</summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">MouseEnterイベント</param>
    ''' <remarks>ウィンドウにマウスオーバー時</remarks>
    Private Sub MainWindow_MouseEnter(sender As Object, e As MouseEventArgs) Handles Me.MouseEnter

        '対象の時計表示文言に現在の対象時計をセット
        txtbTargetTokei.Text = String.Format(cMessage.DisplayTokei, Me.BijinTokei.TargetBijinTokei.TargetTokei)

    End Sub

    ''' <summary>メインウィンドウのMouseLeaveイベント</summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">MouseLeaveイベント</param>
    ''' <remarks>ウィンドウからマウスが離れた時</remarks>
    Private Sub MainWindow_MouseLeave(sender As Object, e As MouseEventArgs) Handles Me.MouseLeave

        '対象の時計表示文言に空文字をセット
        txtbTargetTokei.Text = String.Empty

    End Sub

    ''' <summary>MainWindowのUnloadedイベント</summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">Unloadedイベント</param>
    ''' <remarks></remarks>
    Private Sub MainWindow_Unloaded(sender As Object, e As RoutedEventArgs) Handles Me.Unloaded

        'オーナーウィンドウを閉じる
        '※Me.Close()だけだとプログラムが終了しないのでここでオーナーウィンドウを閉じるようにする
        Me.Owner.Close()

    End Sub

    ''' <summary>MainWindowのMouseLeftButtonDownイベント</summary>
    ''' <param name="sender">MainWindowオブジェクト</param>
    ''' <param name="e">MouseButtonEventArgs</param>
    ''' <remarks>MainWindow上にあるとき（またはマウスがキャプチャされたとき）にマウスの左ボタンが押されると発生する</remarks>
    Private Sub MainWindow_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles Me.MouseLeftButtonDown

        '左クリック押下でドラッグしウインドウを移動出来るようにする
        Me.DragMove()

    End Sub

#End Region

#Region "ドックパネル（dpBijinTokei）"

    ''' <summary>dpBijinTokeiのコンテキストメニューを表示前イベント</summary>
    ''' <param name="sender">dpBijinTokeiオブジェクト</param>
    ''' <param name="e">コンテキストメニューイベントデータ</param>
    ''' <remarks>
    '''   dpBijinTokeiのコンテキスト メニューを開いたときに発生します
    ''' </remarks>
    Private Sub dpBijinTokei_ContextMenuOpening(sender As Object, e As ContextMenuEventArgs) Handles dpBijinTokei.ContextMenuOpening

        'DockPanelにコンテキストメニューを設定
        Me.dpBijinTokei.ContextMenu = _CreateContextMenu(ContextMenueType.DockPanel)

    End Sub

#End Region

#Region "コンテキストメニュー"

    ''' <summary>コンテキストメニューClickイベント</summary>
    ''' <param name="sender">MenuItem</param>
    ''' <param name="e">コンテキストメニューClickイベント</param>
    ''' <remarks></remarks>
    Private Sub ContextMenuItem_Click(sender As Object, e As System.EventArgs)

        'クリックされたアイテム名を取得する
        Dim mClickItemName As String = _GetClickContextMenuItemName(sender)

        'クリックされたメニューアイテムごと処理を分岐
        Select Case True

            Case mClickItemName = ContextMenuItem.最前面に保持.ToString

                '現在の「最上位フォームとして表示するかどうかのプロパティ」の値の逆を設定する
                Me.Topmost = Not (Me.Topmost)

            Case ImageDownloadCount.IsDefined(GetType(ImageDownloadCount), mClickItemName) 'クリックしたメニュー名が美人時計画像ダウンロード数と一致した時

                '選択した美人時計ダウンロード数を取得
                Dim mSelectDownloadCount As ImageDownloadCount = [Enum].Parse(GetType(ImageDownloadCount), mClickItemName)

                '美人時計ダウンロード数変数に選択した美人時計ダウンロード数をセット
                _ImageDownloadCount = mSelectDownloadCount

            Case Me.BijinTokei.TokeiList.Contains(mClickItemName) 'クリックしたメニュー名が美人時計リスト内と一致した時

                '切り替え前の表示時計を取得
                Dim mBeforeSwitchingTargetTokei As Integer = Me.BijinTokei.TargetBijinTokei.Id

                '選択した時計を取得し、表示時計変数に選択した時計をセット
                Dim mSelectTargetTokei As Integer = Me.BijinTokei.TokeiList.IndexOf(mClickItemName)
                _TargetTokeiId = mSelectTargetTokei

                'コントロールを再設定する
                _ToSetupControls(True)

                'トレイアイコンのテキスト設定
                Call _SetTrayIconText()

                '       切り替え前の表示時計が「公式」かつ 選択した時計が「地方」の時
                'または 切り替え前の表示時計が「地方」かつ 選択した時計が「公式」の時
                If Me.BijinTokei.IsOfficial(mBeforeSwitchingTargetTokei) AndAlso Me.BijinTokei.IsLocal(mSelectTargetTokei) _
                OrElse Me.BijinTokei.IsLocal(mBeforeSwitchingTargetTokei) AndAlso Me.BijinTokei.IsOfficial(mSelectTargetTokei) Then

                    Me.Width = Me.BijinTokei.TargetBijinTokei.ImageWidth / Me.BijinTokei.TargetBijinTokei.MinimumSizeRate
                    Me.Height = Me.BijinTokei.TargetBijinTokei.ImageHeight / Me.BijinTokei.TargetBijinTokei.MinimumSizeRate

                End If

            Case mClickItemName = ContextMenuItem.プロフィールを見に行く.ToString

                '現在時刻に表示されている人のプロフィールへ遷移（ブラウザで開く）
                Call _RunFile(Me.BijinTokei.TargetBijinTokei.Url)

            Case mClickItemName = ContextMenuItem.画像のダウンロード.ToString

                '美人時計画像をダウンロード
                Call _DownloadBijinTokeiImage()

            Case mClickItemName = ContextMenuItem.現在の画像を表示する.ToString

                '現在の時刻の画像パスを取得
                Dim mDisplayImagePath As String = _GetImagePath(ImagePathKbn.SavePath, _GetNowTime(TimeFormat.HourAndMinute))

                '現在の時刻の画像が存在しなかった時
                If Not File.Exists(mDisplayImagePath) Then

                    'ファイルが存在しませんメッセージを表示し処理を終了
                    MessageBox.Show(cMessage.NotExistsFile, cMessage.MsgBoxTitle, MessageBoxButton.OK, MessageBoxImage.Information)
                    Exit Sub

                End If

                '現在の時刻の画像を実行する（規定のソフトで開く）
                Call _RunFile(mDisplayImagePath)

            Case mClickItemName = ContextMenuItem.設定ファイル保存フォルダを開く.ToString

                '設定ファイル(user.config)を取得
                Dim mConfig As System.Configuration.Configuration = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal)

                '設定ファイルの親ディレクトリ情報を取得
                Dim mDi As DirectoryInfo = New DirectoryInfo(mConfig.FilePath)
                Dim mParentFolder As DirectoryInfo = mDi.Parent

                '設定ファイルの親ディレクトリが存在しなかったらフォルダを作成する
                If Not mParentFolder.Exists Then mParentFolder.Create()

                '設定ファイルの格納されいてるフォルダを開く
                _RunFile(mParentFolder.FullName)

            Case mClickItemName = ContextMenuItem.閉じる.ToString

                'フォームを閉じる
                Me.Close()

        End Select

    End Sub

#End Region

#Region "トレイアイコン"

    ''' <summary>トレイアイコンのマウスクリックイベント</summary>
    ''' <param name="sender">トレイアイコンオブジェクト</param>
    ''' <param name="e">MouseDownイベント</param>
    ''' <remarks>MouseDownイベント（マウスがクリックされた時）</remarks>
    Private Sub _NotifyIcon_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        'クリックされたボタンごと処理を分岐
        Select Case e.Button

            Case Forms.MouseButtons.Left

                'ウィンドウの最前面状態を取得する
                Dim mTopMostStatus As Boolean = Me.Topmost

                'ウィンドウを最前面にする
                Me.Topmost = True

                'ウィンドウを最前面にする前の最前面状態をセットする
                Me.Topmost = mTopMostStatus

            Case Forms.MouseButtons.Right

                'トレイアイコン用のコンテキストメニューをセット
                _NotifyIcon.ContextMenuStrip = _CreateContextMenu(ContextMenueType.TrayIcon)

        End Select

    End Sub

#End Region

#Region "その他"

    ''' <summary>SourceInitializedイベントを発生</summary>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>ウインドウの初期化中に呼び出されます</remarks>
    Protected Overrides Sub OnSourceInitialized(ByVal e As EventArgs)

        '基底クラスのSourceInitializedイベントを発生させる
        MyBase.OnSourceInitialized(e)

        '----------------------------------------
        ' ウィンドウプロシージャをフックする設定
        '----------------------------------------
        'WPFコンテンツを格納するWin32のウィンドウを取得する
        Dim mHwndSource As HwndSource = CType(HwndSource.FromVisual(Me), HwndSource)

        'ウィンドウメッセージを受信するイベントハンドラーを追加
        mHwndSource.AddHook(AddressOf WndHookProc)

        '----------------------------------------
        ' アプリケーション設定から値を取得
        '----------------------------------------
        'ウィンドウの位置／サイズ／状態をWindowPlacementから取得
        Dim mWindowPlacement As BijinTokeiDefinition.WINDOWPLACEMENT = My.Settings.WindowPlacement

        '最前面状態をIsTopMostから取得
        Dim mIsTopMost As Boolean = My.Settings.IsTopMost

        '美人時計画像ダウンロード数をImageDownloadCountから取得
        _ImageDownloadCount = My.Settings.ImageDownloadCount

        '対象時計をTargetTokeiから取得
        _TargetTokeiId = My.Settings.TargetTokei

        '----------------------------------------
        ' 前回起動時ウィンドウの状態を復元
        '----------------------------------------
        'ウィンドウの位置／サイズ／状態を設定
        Dim mWp As WINDOWPLACEMENT = mWindowPlacement
        mWp.length = Marshal.SizeOf(GetType(WINDOWPLACEMENT))
        mWp.flags = 0

        'ウィンドウの状態が最小化の場合、位置とサイズを元に戻す
        mWp.showCmd = IIf((mWp.showCmd = SW_SHOWMINIMIZED), SW_SHOWNORMAL, mWp.showCmd)

        'ウィンドウの位置／サイズ／状態を設定
        MainWindow.SetWindowPlacement(New WindowInteropHelper(Me).Handle, mWp)

        'ウィンドウの最前面状態を設定
        Me.Topmost = mIsTopMost

    End Sub

    ''' <summary>MainWindowのClosingイベント</summary>
    ''' <param name="e">キャンセルできるイベントのデータ</param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)

        '----------------------------------------------
        ' ウィンドウプロシージャをフックする設定を解除
        '----------------------------------------------
        'WPFコンテンツを格納するWin32のウィンドウを取得する
        Dim mHwndSource As HwndSource = DirectCast(PresentationSource.FromVisual(Me), HwndSource)

        'ウィンドウメッセージを受信するイベントハンドラーを削除
        mHwndSource.RemoveHook(AddressOf WndHookProc)

        '----------------------------------------------
        ' Closedイベントを発生
        '----------------------------------------------
        '基底クラスのClosedイベントを発生させる
        MyBase.OnClosing(e)

        '----------------------------------------------
        ' アプリケーション設定に現在の状態を設定
        '----------------------------------------------
        'ウィンドウの位置／サイズ／状態を取得
        Dim mWp As WINDOWPLACEMENT = New WINDOWPLACEMENT
        MainWindow.GetWindowPlacement(New WindowInteropHelper(Me).Handle, mWp)

        'ウィンドウの位置／サイズ／状態をアプリケーション設定のWindowPlacementへセット
        My.Settings.WindowPlacement = mWp

        'ウィンドウの最前面状態をアプリケーション設定のIsTopMostへセット
        My.Settings.IsTopMost = Me.Topmost

        '美人時計画像ダウンロード数をアプリケーション設定のImageDownloadCountへセット
        My.Settings.ImageDownloadCount = _ImageDownloadCount

        '対象時計をアプリケーション設定のTargetTokeiへセット
        My.Settings.TargetTokei = _TargetTokeiId

        'アプリケーション設定を保存
        My.Settings.Save()

    End Sub

    ''' <summary>ウィンドウプロシージャをフック</summary>
    ''' <param name="hwnd">ウィンドウのハンドル</param>
    ''' <param name="msg">メッセージの識別子</param>
    ''' <param name="wParam">メッセージの最初のパラメータ</param>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <param name="handled">ハンドルフラグ</param>
    ''' <returns>0 に初期化されたポインターまたはハンドル</returns>
    ''' <remarks>
    '''   ウインドウプロシージャ
    '''     メッセージを処理する専用のルーチン
    '''   Hook（フック）
    '''     独自の処理を割り込ませるための仕組み
    '''      注意：デバッグ時はこのメソッドの処理で止まりません。処理を確認した時は「System.Diagnostics.DebuggerStepThrough()」行を削除して下さい
    ''' </remarks>
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Function WndHookProc(ByVal hwnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr

        If msg = cWM_SIZING.Message Then

            'アスペクト比を保ったままウィンドウサイズを変更
            Call _ResizeWindowKeepingAspectRatio(wParam, lParam)

        ElseIf msg = cWM_NCHITTEST.Message Then

            '現在のマウス位置を返す
            Return _GetMousePotisionInTheForm(lParam, handled)

        End If

        Return IntPtr.Zero

    End Function

#End Region

#End Region

#Region "メソッド"

#Region "ウィンドウプロシージャをフック関連"

    ''' <summary>ウィンドウサイズをアスペクト比率を保って変更</summary>
    ''' <param name="wParam">メッセージの最初のパラメータ</param>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <remarks>
    '''   「ウィンドウプロシージャをフック」するイベントから呼ばれます
    '''    ※ウィンドウのサイズを変更した時、アスペクト比を保ったままサイズが変更されるようにするためのメソッド
    ''' </remarks>
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub _ResizeWindowKeepingAspectRatio(ByVal wParam As IntPtr, ByVal lParam As IntPtr)

        'アンマネージメモリのRECT構造体をマネージオブジェクト（RECT構造体）にデータをマーシャリングする
        '※ウィンドウプロシージャに渡ってきた「lParam」を.NET側で使えるようにデータを変換する。Marshalingは「整列」という意味の英単語
        Dim mRect As RECT = Marshal.PtrToStructure(lParam, GetType(RECT))

        'ウィンドウの幅と高さを求める
        Dim mWindowWidth As Integer = mRect.Right - mRect.Left
        Dim mWindowHeight As Integer = mRect.Bottom - mRect.Top

        'ウィンドウの幅と高さの増減値を取得  
        ' ウィンドウ幅  の増減値：「(ウィンドウ高さ * 修正率) - ウィンドウ幅  」
        ' ウィンドウ高さの増減値：「(ウィンドウ幅   * 修正率) - ウィンドウ高さ」
        Dim mChangeWidth As Integer = Math.Round((mWindowHeight * Me.WindowFixRate)) - mWindowWidth
        Dim mChangeHeight As Integer = Math.Round((mWindowWidth / Me.WindowFixRate)) - mWindowHeight

        Select Case wParam.ToInt32()

            Case cWM_SIZING.wParam.WMSZ_LEFT, cWM_SIZING.wParam.WMSZ_RIGHT

                '「左端」と「右端」の時は、ウインドウ幅の増減値を右下隅のＹ座標に設定
                mRect.Bottom = mRect.Bottom + mChangeHeight

            Case cWM_SIZING.wParam.WMSZ_TOP, cWM_SIZING.wParam.WMSZ_BOTTOM

                '「上端」と「下端」の時は、ウインドウ高さの増減値を右下隅のＸ座標に設定
                mRect.Right = mRect.Right + mChangeWidth

            Case cWM_SIZING.wParam.WMSZ_TOPLEFT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの左位置を再設定「ウィンドウの左位置 - ウィンドウ幅の増減値」
                    mRect.Left = mRect.Left - mChangeWidth

                Else

                    'ウィンドウの上位置を再設定「ウィンドウの上位置 - ウィンドウ高さの増減値」
                    mRect.Top = mRect.Top - mChangeHeight

                End If

            Case cWM_SIZING.wParam.WMSZ_TOPRIGHT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの右位置を再設定「ウィンドウの右位置 + ウィンドウ幅の増減値」
                    mRect.Right = mRect.Right + mChangeWidth

                Else

                    'ウィンドウの上位置を再設定「ウィンドウの上位置 - ウィンドウ高さの増減値」
                    mRect.Top = mRect.Top - mChangeHeight

                End If

            Case cWM_SIZING.wParam.WMSZ_BOTTOMLEFT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの左位置を再設定「ウィンドウの左位置 - ウィンドウ幅の増減値」
                    mRect.Left = mRect.Left - mChangeWidth

                Else

                    'ウィンドウの下位置を再設定「ウィンドウの下位置 + ウィンドウ高さの増減値」
                    mRect.Bottom = mRect.Bottom + mChangeHeight

                End If

            Case cWM_SIZING.wParam.WMSZ_BOTTOMRIGHT

                'ウィンドウ幅の増減値が０より大きい時
                If (mChangeWidth > 0) Then

                    'ウィンドウの右位置を再設定「ウィンドウの右位置 + ウィンドウ幅の増減値」
                    mRect.Right = mRect.Right + mChangeWidth

                Else

                    'ウィンドウの下位置を再設定「ウィンドウの下位置 + ウィンドウ高さの増減値」
                    mRect.Bottom = mRect.Bottom + mChangeHeight

                End If

        End Select

        'マネージオブジェクト（RECT構造体）をアンマネージメモリブロックにデータをマーシャリングする
        '※この処理で変更したRECT構造体の値を
        Marshal.StructureToPtr(mRect, lParam, False)

    End Sub

    ''' <summary>マウス位置を取得</summary>
    ''' <param name="lParam">メッセージの２番目のパラメータ</param>
    ''' <param name="handled">ハンドルフラグ</param>
    ''' <returns>現在のマウス位置を返す</returns>
    ''' <remarks>
    '''  ・「ウィンドウプロシージャをフック」するイベントから呼ばれます
    '''  ・現在のウィンドウのスタイルだと右下の部分でしかウィンドウサイズの変更が不可能である
    '''      スタイルの設定：WindowStyle="None" AllowsTransparency="True" ResizeMode="CanResizeWithGrip"
    '''    そのためウィンドウの端に来た時、リサイズ可能領域内（マウス位置がキャプションバー内）であることを知らせる
    '''  ・スクリーン座標とクライアント座標について
    '''      スクリーン座標            ：画面の左上隅の点を原点とした座標
    '''      フォームのクライアント座標：フォームの描画可能なクライアント領域の左上隅の点を原点とした座標
    '''      ※参考URL：http://www.atmarkit.co.jp/fdotnet/dotnettips/377screentoclient/screentoclient.html
    ''' </remarks>
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Function _GetMousePotisionInTheForm(ByVal lParam As IntPtr, ByRef handled As Boolean) As IntPtr

        'これ以上処理させない（完全に処理を横取りする）
        handled = True

        '------------------------------------
        ' クライアント座標のマウス位置を取得
        '------------------------------------
        'スクリーン座標のマウス位置を取得
        Dim mMousePositionOnScreen As New System.Windows.Point(CInt(lParam) And &HFFFF, (CInt(lParam) >> 16) And &HFFFF)

        'スクリーン座標のマウス位置をクライアント座標のマウス位置に変換
        Dim mMousePositionOnClient As System.Windows.Point = PointFromScreen(mMousePositionOnScreen)

        '------------------------------------
        ' リサイズ可能とするサイズを取得
        '------------------------------------
        'ウィンドウの周囲にある水平サイズ変更境界の高さサイズを取得
        Dim ResizableHorizontal As Double = SystemParameters.ResizeFrameHorizontalBorderHeight

        'ウィンドウの周囲にある垂直サイズ変更境界の幅サイズを取得
        Dim ResizableVertical As Double = SystemParameters.ResizeFrameVerticalBorderWidth

        'タイトルバーの高さを取得
        Dim ResizableCaptionHeader As Double = SystemParameters.CaptionHeight

        '------------------------------------
        ' 四隅の斜め方向にリサイズ可能
        '------------------------------------
        '左上の斜め方向にリサイズ可能
        If New System.Windows.Rect(0, 0, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTTOPLEFT)

        '右上の斜め方向にリサイズ可能
        If New System.Windows.Rect(Me.Width - ResizableVertical, 0, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTTOPRIGHT)

        '左下の斜め方向にリサイズ可能
        If New System.Windows.Rect(0, Height - ResizableHorizontal, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTBOTTOMLEFT)

        '右下の斜め方向にリサイズ可能
        If New System.Windows.Rect(Me.Width - ResizableVertical, Me.Height - ResizableHorizontal, ResizableVertical, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTBOTTOMRIGHT)

        '------------------------------------
        ' 四辺の直交方向にリサイズ可能
        '------------------------------------
        '上に直交方向にリサイズ可能
        If New System.Windows.Rect(0, 0, Me.Width, ResizableVertical).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTTOP)

        '左に直交方向にリサイズ可能
        If New System.Windows.Rect(0, 0, ResizableVertical, Me.Height).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTLEFT)

        '右に直交方向にリサイズ可能
        If New System.Windows.Rect(Me.Width - ResizableVertical, 0, ResizableVertical, Me.Height).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTRIGHT)

        '下に直交方向にリサイズ可能
        If New System.Windows.Rect(0, Me.Height - ResizableHorizontal, Me.Width, ResizableHorizontal).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTBOTTOM)

        '------------------------------------
        ' タイトルバーにマウスがあるか判断
        '------------------------------------
        'マウスがタイトルバーにある
        If New System.Windows.Rect(0, 0, Me.Width, ResizableCaptionHeader).Contains(mMousePositionOnClient) Then Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTCAPTION)

        '上記以外はクライアント領域とする
        Return New IntPtr(cWM_NCHITTEST.CursorHotSpot.HTCLIENT)

    End Function

#End Region

#Region "コントロール設定（全般）"

    ''' <summary>コントロール設定</summary>
    ''' <param name="pIsReRun">再実行か（デフォルトはFalse：再実行でない）</param>
    ''' <remarks>
    '''  トレイアイコン以外のコントロールの設定を行う
    '''  ※トレイアイコンはWPFの機能では無いのでこことは別で設定を行う
    ''' </remarks>
    Private Sub _ToSetupControls(Optional ByVal pIsReRun As Boolean = False)

        'メインウィンドウ設定
        Call _ToSetupMainWindow(pIsReRun)

        'DockPanel設定
        Call _ToSetupDockPanel()

        'コントロール設定が再実行でなければ、画像切り替え用タイマー設定を行う
        If Not pIsReRun Then Call _ToSetupChangeImageTimer()

        '美人時計画像をセット
        Call _SetImageBijinTokei()

    End Sub

#End Region

#Region "メインウィンドウ（MainWindow）設定関連"

    ''' <summary>MainWindow設定</summary>
    ''' <param name="pIsReRun">再実行か（デフォルトはFalse：再実行でない）</param>
    ''' <remarks></remarks>
    Private Sub _ToSetupMainWindow(Optional ByVal pIsReRun As Boolean = False)

        'ウィンドウの最大サイズ設定（美人時計画像サイズ）
        Me.MaxWidth = Me.BijinTokeiImageSize.Width
        Me.MaxHeight = Me.BijinTokeiImageSize.Height

        'ウィンドウの最小サイズ設定
        Me.MinWidth = Me.BijinTokeiImageSize.Width / Me.BijinTokeiImageSize.MinimumSizeRate
        Me.MinHeight = Me.BijinTokeiImageSize.Height / Me.BijinTokeiImageSize.MinimumSizeRate

        'MainWindow設定が再実行でなければ、「Alt+Tab」ウィンドウから非表示処理を実行
        If Not pIsReRun Then Call _HideAltTabWindow()

    End Sub

    ''' <summary>「Alt+Tab」ウインドウに表示させない</summary>
    ''' <remarks>
    '''   Alt+Tab ダイアログに表示されない Window をオーナーに設定し
    '''   ShowInTaskbar プロパティに False を設定することでAlt+Tabに表
    '''   示されなくなる。オーナーウィンドウを非表示にしておけば、
    '''   SingleBorderWindow や ThreeDBorderWindow の Window が 
    '''   Alt+Tab ダイアログに表示されないようにすることができます。
    ''' </remarks>
    Private Sub _HideAltTabWindow()

        'オーナー用ウィンドウを作成
        Dim mOwnerWindow As New Window

        'Alt＋Tabに表示されない設定にする
        mOwnerWindow.WindowStyle = Windows.WindowStyle.ToolWindow
        mOwnerWindow.ShowInTaskbar = False

        '表示領域外にオーナー用ウィンドウを表示させる（一旦表示させる必要があるため）
        mOwnerWindow.Left = -100
        mOwnerWindow.Height = 0
        mOwnerWindow.Width = 0
        mOwnerWindow.Show()

        'オーナー用ウィンドウを非表示
        mOwnerWindow.Hide()

        'オーナーウィンドウにオーナー用ウィンドウを設定する
        Me.Owner = mOwnerWindow

    End Sub

#End Region

#Region "ドックパネル（dpBijinTokei）設定関連"

    ''' <summary>DockPanel設定</summary>
    ''' <remarks></remarks>
    Private Sub _ToSetupDockPanel()

        'DockPanelの最大サイズ設定（美人時計画像サイズ）
        Me.dpBijinTokei.MaxWidth = Me.BijinTokeiImageSize.Width
        Me.dpBijinTokei.MaxHeight = Me.BijinTokeiImageSize.Height

        'DockPanelの最小サイズ設定
        Me.dpBijinTokei.MinWidth = Me.BijinTokeiImageSize.Width / Me.BijinTokeiImageSize.MinimumSizeRate
        Me.dpBijinTokei.MinHeight = Me.BijinTokeiImageSize.Height / Me.BijinTokeiImageSize.MinimumSizeRate

        'コンテキストメニューを設定
        Me.dpBijinTokei.ContextMenu = _CreateContextMenu(ContextMenueType.DockPanel)

    End Sub

#End Region

#Region "トレイアイコン設定関連"

    ''' <summary>トレイアイコン設定</summary>
    ''' <remarks></remarks>
    Private Sub _ToSetupTrayIcon()

        _NotifyIcon = New System.Windows.Forms.NotifyIcon

        'アイコンを設定(リソースからセット)
        _NotifyIcon.Icon = My.Resources.App

        'トレイアイコンを表示させる
        _NotifyIcon.Visible = True

        'マウスのクリックイベントを設定
        AddHandler _NotifyIcon.MouseDown, New System.Windows.Forms.MouseEventHandler(AddressOf _NotifyIcon_MouseDown)

    End Sub

    ''' <summary>トレイアイコンのテキスト設定</summary>
    ''' <remarks>
    '''   マウスオーバー時のテキストを設定します
    ''' </remarks>
    Private Sub _SetTrayIconText()

        _NotifyIcon.Text = String.Format(cMessage.TrayIconText, Me.BijinTokei.TargetBijinTokei.TargetTokei)

    End Sub

    ''' <summary>トレイアイコンのテキスト設定</summary>
    ''' <param name="pProgress">ファイルダウンロードの進捗率情報</param>
    ''' <remarks>
    '''   マウスオーバー時のテキストとダウンロードの完了時バルーンウィンドウの設定を行う
    ''' </remarks>
    Private Sub _SetTrayIconText(ByVal pProgress As DownloadForNet.FileDownloadProgress)

        'ダウンロード処理が完了の時
        If pProgress.IsComplete Then

            '表示メッセージを作成
            Dim BalloonWindowText As String = String.Format(cMessage.DownloadComplete, Me.BijinTokei.TargetBijinTokei.TargetTokei, Me.ImageDownloadCount.ToString)

            'バルーンウィンドウにダウンロード完了メッセージを設定
            Call _SetBalloonWindowText(BalloonWindowText, BalloonWindowDisplayTime.DownloadComplete)

            'トレイアイコンのテキスト設定
            Call _SetTrayIconText()

        Else

            '表示メッセージを作成
            Dim mMouseOverText As String = String.Format(cMessage.DownloadStatus, Me.BijinTokei.TargetBijinTokei.TargetTokei, pProgress.DownloadedFileCount, pProgress.DownloadFileCount)

            'マウスオーバー時のテキストを設定
            '※文言例 ○○版　ダウンロード状況：120/1440
            _NotifyIcon.Text = mMouseOverText

        End If

    End Sub

    ''' <summary>バルーンウィンドウテキスト設定</summary>
    ''' <param name="pMessage">メッセージ</param>
    ''' <param name="pDisplayTime">バルーンウィンドウ表示時間</param>
    ''' <remarks></remarks>
    Private Sub _SetBalloonWindowText(ByVal pMessage As String, ByVal pDisplayTime As BalloonWindowDisplayTime)

        'バルーンウィンドウに'ローカルの画像使用メッセージを表示させる
        _NotifyIcon.BalloonTipText = pMessage
        _NotifyIcon.ShowBalloonTip(pDisplayTime)

    End Sub

#End Region

#Region "画像切り替え用タイマー設定関連"

    ''' <summary>画像切り替え用タイマー設定</summary>
    ''' <remarks></remarks>
    Private Sub _ToSetupChangeImageTimer()

        'インスタンスを作成
        _ChangeImageTimer = New System.Windows.Threading.DispatcherTimer

        'フォームポップアップ用タイマーにTickイベントを設定
        '※イベントの中の処理が１行で済む場合はラムダ式で記述したほうがスッキリする
        AddHandler _ChangeImageTimer.Tick, Sub(sender As Object, e As EventArgs) _SetImageBijinTokei()

        'Tickイベントが発生する間隔を設定
        _ChangeImageTimer.Interval = cImageChangeTime

        'フォームポップアップ用タイマーを起動
        _ChangeImageTimer.Start()

    End Sub

#End Region

#Region "時間関連"

    ''' <summary>時間のリストを作成</summary>
    ''' <returns>時間のリスト</returns>
    ''' <remarks>順番は現在時刻～MAX時間、MIN時間～現在時刻という順番になる
    '''          例：現在時刻が「15:40」の時
    '''              1540,1541,1542……2358,2359,0000,0001,……1537,1538,1539</remarks>
    Private Function _CreateTimeList() As ArrayList

        Dim mTimeList As New ArrayList

        '---------------------------------------
        '現在時刻の「時間」と「分」を取得
        '---------------------------------------
        '現在時刻の時間と分を文字列で取得する
        Dim mNowHourString As String = String.Format("{0:D1}", _GetNowTime(TimeFormat.Hour))
        Dim mNowMinuteString As String = String.Format("{0:D1}", _GetNowTime(TimeFormat.Minute))

        '現在時刻の時間と分を数値に変換
        Dim mNowHour As Integer = Integer.Parse(mNowHourString)
        Dim mNowMinute As Integer = Integer.Parse(mNowMinuteString)

        '---------------------------------------
        '現在時刻～MAX時間までの時間リストを作成
        '---------------------------------------
        '「mNowHour ～ MAX時間 - 1」まで繰り返す
        For i As Integer = mNowHour To cTimeDef.Hour - 1

            '分の繰り返しスタート数値（i と 現在時刻の時間部分 が一致しなかった時は「0」をセットする）
            '※現在時刻 ～ 23:59 まで作成する用の処理
            Dim mRoopStartMinute As Integer = mNowMinute
            If i <> mNowHour Then mRoopStartMinute = 0

            '「mRoopStartMinute  ～ MAX分 - 1」まで繰り返す
            For j As Integer = mRoopStartMinute To cTimeDef.Minute - 1

                '時間と分の文字列を「99」の形式で作成
                '※00,01,02,……,57,58,59
                Dim mHour As String = String.Format("{0:D2}", i)
                Dim mMinute As String = String.Format("{0:D2}", j)

                mTimeList.Add(mHour & mMinute)

            Next

        Next

        '---------------------------------------
        'MIN時間～現在時刻までの時間リストを作成
        '---------------------------------------
        '「0 ～ 現在時刻」まで繰り返す
        For i As Integer = 0 To mNowHour

            '分の繰り返しスタート数値（i と 現在時刻の時間部分 が一致しなかった時は「59」をセットする）
            '※00：00 ～ 現在時刻 まで作成する用の処理
            Dim mRoopEndMinute As Integer = mNowMinute - 1
            If i <> mNowHour Then mRoopEndMinute = cTimeDef.Minute - 1

            '「mRoopStartMinute  ～ MAX分 - 1」まで繰り返す
            For j As Integer = 0 To mRoopEndMinute

                '時間と分の文字列を「99」の形式で作成
                '※00,01,02,……,57,58,59
                Dim mHour As String = String.Format("{0:D2}", i)
                Dim mMinute As String = String.Format("{0:D2}", j)

                mTimeList.Add(mHour & mMinute)

            Next

        Next

        Return mTimeList

    End Function

    ''' <summary>現在時刻の文字列の取得</summary>
    ''' <param name="pTimeFormat">時間形式</param>
    ''' <returns>時間形式にあった現在時刻の文字列</returns>
    ''' <remarks>取得結果例
    '''            時間    ：99（00～24）
    '''            分      ：99（00～59）
    '''            時間＆分：9999（0000～2359）</remarks>
    Private Function _GetNowTime(Optional ByVal pTimeFormat As TimeFormat = TimeFormat.HourAndMinute) As String

        Dim mTime As String = String.Empty

        Select Case pTimeFormat

            Case TimeFormat.Hour

                '現在時刻の「時間」の部分をセット
                mTime = DateTime.Now.ToString("HH")

            Case TimeFormat.Minute

                '現在時刻の「分」の部分をセット
                mTime = DateTime.Now.ToString("mm")

            Case TimeFormat.HourAndMinute

                '現在時刻の「時間＆分」の部分をセット
                mTime = DateTime.Now.ToString("HHmm")

        End Select

        Return mTime

    End Function

#End Region

#Region "コンテキストメニュー関連"

#Region "コンテキストメニュー全般"

    ''' <summary>コンテキストメニューを作成</summary>
    ''' <param name="pMenuType">コンテキストメニュータイプ</param>
    ''' <returns>コンテキストメニュー</returns>
    ''' <remarks>
    '''   コンテキストメニュータイプごとそれに対応したコンテキストメニューを返す
    ''' </remarks>
    Private Function _CreateContextMenu(ByVal pMenuType As ContextMenueType) As Object

        'コンテキストメニューのインスタンスを作成
        Dim mContextMenu As Object = _CreateContextMenuInstance(pMenuType)

        'コンテキストメニューのアイテムリストを作成
        mContextMenu = _GetContextMenu(Of ContextMenuItem)(mContextMenu, pMenuType)

        '美人時計画像のダウンロード数の設定により対象の時間にチェックマークを表示させる
        Call _SetCheckMarkForContextMenu(mContextMenu.Items(ContextMenuItem.画像のダウンロード数) _
                                       , _GetHierarchyMenuItemType(mContextMenu.Items(ContextMenuItem.画像のダウンロード数)) _
                                       , cImageDownloadCountId(Me.ImageDownloadCount) _
                                       , True)

        '「対象時計」になっているアイテムの隣にチェックマークを表示させます
        Call _SetCheckMarkForContextMenu(mContextMenu.Items(ContextMenuItem.対象時計) _
                                       , _GetHierarchyMenuItemType(mContextMenu.Items(ContextMenuItem.対象時計)) _
                                       , Me.BijinTokei.TargetBijinTokei.Id _
                                       , True)

        '「最前面に保持」アイテムの横にウィンドウの状態によりチェックマークを表示させる
        '※最前面：チェックマークあり、最前面でない：チェックマークなし
        Call _SetCheckMarkForContextMenu(mContextMenu, pMenuType, ContextMenuItem.最前面に保持, Me.Topmost)

        '美人時計画像がダウンロード中の時
        If Me.IsDownloadingForBijinTokeiImages Then

            '「画像のダウンロード数」アイテムのEnableを不可に設定
            Call _SetEnabledForContextMenu(mContextMenu, pMenuType, ContextMenuItem.画像のダウンロード数, False)

            '「対象時計」アイテムのEnableを不可に設定
            Call _SetEnabledForContextMenu(mContextMenu, pMenuType, ContextMenuItem.対象時計, False)

        End If

        '美人時計画像がダウンロード可能でない時、「画像のダウンロード」アイテムのEnableを不可に設定
        If Not Me.IsAvailableDownloadBijinTokei Then Call _SetEnabledForContextMenu(mContextMenu, pMenuType, ContextMenuItem.画像のダウンロード, False)

        Return mContextMenu

    End Function

    ''' <summary>コンテキストメニューを取得する</summary>
    ''' <typeparam name="MenuList">コンテキストメニューを作成する列挙体（※必ず列挙体を指定すること）</typeparam>
    ''' <param name="pMenuItem">コンテキストメニュー</param>
    ''' <param name="pMenuType">コンテキストメニュータイプ</param>
    ''' <returns>作成したコンテキストメニュー</returns>
    ''' <remarks>
    '''   コンテキストメニュー作成用の列挙体の項目数分メニューを作成する
    '''   またメニュー名に応じて区切り線や階層メニューを作成する
    ''' </remarks>
    Private Function _GetContextMenu(Of MenuList)(ByVal pMenuItem As Object, ByVal pMenuType As ContextMenueType) As Object

        '列挙体数分繰り返す（コンテキストメニューを作成していく）
        For Each mItemName As String In System.Enum.GetNames(GetType(MenuList))

            'メニュー名から作成するコンテキストメニューアイテムタイプを取得
            Dim mContextMenuItemType As ContextMenuItemType = GetCreateContextMenuItemType(mItemName)

            'コンテキストメニューアイテムを作成する
            Call _CreateContextMenuItem(pMenuItem, mItemName, pMenuType, mContextMenuItemType)

        Next

        Return pMenuItem

    End Function

#End Region

#Region "コンテキストメニューアイテム関連"

    ''' <summary>コンテキストメニューアイテムを作成</summary>
    ''' <param name="pMenu"></param>
    ''' <param name="pMenuName">メニュー名</param>
    ''' <param name="pMenuType">コンテキストメニュータイプ</param>
    ''' <param name="pMenuItemType">コンテキストメニューアイテムタイプ</param>
    ''' <remarks></remarks>
    Private Sub _CreateContextMenuItem(ByVal pMenu As Object, ByVal pMenuName As String, ByVal pMenuType As ContextMenueType, ByVal pMenuItemType As ContextMenuItemType)

        Select Case pMenuItemType

            Case ContextMenuItemType.Normal

                'コンテキストメニューアイテム（通常）を作成する
                pMenu.Items.Add(_CreateContextMenuNormalItem(pMenuName, pMenuType))

            Case ContextMenuItemType.HierarchyMenu

                If pMenuName = ContextMenuItem.画像のダウンロード数.ToString Then

                    'コンテキストメニューアイテム（画像のダウンロード数階層メニュー）を作成する
                    pMenu.Items.Add(_CreateContextMenuHierarchyItem(Of ImageDownloadCount)(pMenuName, pMenuType))

                ElseIf pMenuName = ContextMenuItem.対象時計.ToString Then

                    'コンテキストメニューアイテム（対象時計階層メニュー）を作成する
                    pMenu.Items.Add(_CreateContextMenuHierarchyItem(pMenuName, pMenuType, Me.BijinTokei.TokeiList))

                End If

            Case ContextMenuItemType.Separator

                'コンテキストメニューアイテム（区切り線）を作成する
                pMenu.Items.Add(_CreateContextMenuSeparatorItem(pMenuType))

        End Select

    End Sub

    ''' <summary>コンテキストメニューアイテムタイプを取得</summary>
    ''' <param name="pMenuName">メニュー名</param>
    ''' <returns>コンテキストメニューアイテムタイプ</returns>
    ''' <remarks>
    '''   メニュー名に応じて作成するコンテキストメニューアイテムタイプを判断し返す
    ''' </remarks>
    Private Function GetCreateContextMenuItemType(ByVal pMenuName As String) As ContextMenuItemType

        If cContextMenuSeparatorItems.Contains(pMenuName) Then

            'アイテム名に「区切り線」という文字列が含まれる時は「区切り線」を返す
            Return ContextMenuItemType.Separator

        ElseIf cContextMenuHierarchyMenuItems.Contains(pMenuName) Then

            'アイテム名に「対象時計」、「画像のダウンロード数」という文字列が含まれる時は「階層メニュー」を返す
            Return ContextMenuItemType.HierarchyMenu

        Else

            '上記以外は「通常アイテム」を返す
            Return ContextMenuItemType.Normal

        End If

    End Function

    ''' <summary>コンテキストメニューアイテム（通常）を作成する</summary>
    ''' <param name="pMenuName">メニューアイテム名</param>
    ''' <param name="pContextMenuType">作成するメニューのタイプ（WPF用、トレイアイコン用）</param>
    ''' <returns>メニュータイプに対応してコンテキストメニューアイテム（通常）</returns>
    ''' <remarks></remarks>
    Private Function _CreateContextMenuNormalItem(ByVal pMenuName As String, ByVal pContextMenuType As ContextMenueType) As Object

        Dim mAddMenuItem As New Object

        'メニュータイプによって作成するアイテムの型を変更する
        Select Case pContextMenuType

            Case ContextMenueType.DockPanel

                'WPF用のコンテキストメニュー作成時
                mAddMenuItem = New System.Windows.Controls.MenuItem

                'メニューの名前を設定
                mAddMenuItem.Header = pMenuName

                'メニューのクリック時のイベントを設定
                AddHandler DirectCast(mAddMenuItem, System.Windows.Controls.MenuItem).Click, New System.Windows.RoutedEventHandler(AddressOf ContextMenuItem_Click)

            Case ContextMenueType.TrayIcon

                'トレイアイコン用のコンテキストメニュー作成時
                mAddMenuItem = New System.Windows.Forms.ToolStripMenuItem

                'メニューの名前を設定
                mAddMenuItem.Text = pMenuName

                'メニューのクリック時のイベントを設定
                AddHandler DirectCast(mAddMenuItem, System.Windows.Forms.ToolStripMenuItem).Click, New System.EventHandler(AddressOf ContextMenuItem_Click)

        End Select

        Return mAddMenuItem

    End Function

    ''' <summary>コンテキストメニューアイテム（区切り線）を作成する</summary>
    ''' <param name="pContextMenuType">作成するメニューのタイプ（WPF用、トレイアイコン用）</param>
    ''' <returns>メニュータイプに対応したコンテキストメニューアイテム（区切り線）</returns>
    ''' <remarks></remarks>
    Private Function _CreateContextMenuSeparatorItem(ByVal pContextMenuType As ContextMenueType) As Object

        Dim mAddMenuItem As New Object

        'メニュータイプによって作成するアイテムの型を変更する
        Select Case pContextMenuType

            Case ContextMenueType.DockPanel

                'WPF用のコンテキストメニューの区切り線を作成
                mAddMenuItem = New System.Windows.Controls.Separator

            Case ContextMenueType.TrayIcon

                'トレイアイコン用のコンテキストメニューの区切り線を作成
                mAddMenuItem = New System.Windows.Forms.ToolStripSeparator

        End Select

        Return mAddMenuItem

    End Function

    ''' <summary>コンテキストメニューアイテム（階層メニュー）を作成する</summary>
    ''' <typeparam name="MenuList">階層メニューを作成する列挙体（※必ず列挙体を指定すること）</typeparam>
    ''' <param name="pMenuName">メニュー名</param>
    ''' <param name="pMenuType">コンテキストメニュータイプ（WPF用、トレイアイコン用）</param>
    ''' <returns>階層メニューアイテム</returns>
    ''' <remarks>
    '''   「MenuList」列挙体から階層メニューを作成する
    ''' </remarks>
    Private Function _CreateContextMenuHierarchyItem(Of MenuList)(ByVal pMenuName As String, ByVal pMenuType As ContextMenueType) As Object

        'コンテキストメニューのインスタンスを作成
        Dim mContextMenu As Object

        Select Case pMenuType

            Case ContextMenueType.DockPanel

                'WPF用のコンテキストメニューアイテムのインスタンスを作成
                mContextMenu = _CreateContextMenuInstance(ContextMenueType.DockPanelItem)

                'WPF用のコンテキストメニュー名をセット
                mContextMenu.Header = pMenuName

                '列挙体数分繰り返す（階層メニューの項目を追加していく）
                For Each mItemName As String In System.Enum.GetNames(GetType(MenuList))

                    mContextMenu.Items.Add(_CreateContextMenuNormalItem(mItemName, ContextMenueType.DockPanel))

                Next

            Case ContextMenueType.TrayIcon

                'トレイアイコン用のコンテキストメニューアイテムのインスタンスを作成
                mContextMenu = _CreateContextMenuInstance(ContextMenueType.TrayIconItem)

                'トレイアイコン用のコンテキストメニュー名をセット
                mContextMenu.Text = pMenuName

                '列挙体数分繰り返す（階層メニューの項目を追加していく）
                For Each mItemName As String In System.Enum.GetNames(GetType(MenuList))

                    mContextMenu.DropDownItems.Add(_CreateContextMenuNormalItem(mItemName, ContextMenueType.TrayIcon))

                Next

            Case Else

                mContextMenu = Nothing

        End Select

        Return mContextMenu

    End Function

    ''' <summary>コンテキストメニューアイテム（階層メニュー）を作成する</summary>
    ''' <param name="pMenuName">メニュー名</param>
    ''' <param name="pMenuType">コンテキストメニュータイプ（WPF用、トレイアイコン用）</param>
    ''' <param name="pMenuList">コンテキストメニューに作成するリスト</param>
    ''' <returns>階層メニューアイテム</returns>
    ''' <remarks>
    '''   「pMenuList」リストから階層メニューを作成する
    ''' </remarks>
    Private Function _CreateContextMenuHierarchyItem(ByVal pMenuName As String, ByVal pMenuType As ContextMenueType, ByVal pMenuList As List(Of String)) As Object

        'コンテキストメニューのインスタンスを作成
        Dim mContextMenu As Object

        Select Case pMenuType

            Case ContextMenueType.DockPanel

                'WPF用のコンテキストメニューアイテムのインスタンスを作成
                mContextMenu = _CreateContextMenuInstance(ContextMenueType.DockPanelItem)

                'WPF用のコンテキストメニュー名をセット
                mContextMenu.Header = pMenuName

                'リスト数分繰り返す（階層メニューの項目を追加していく）
                For Each mItemName As String In pMenuList

                    mContextMenu.Items.Add(_CreateContextMenuNormalItem(mItemName, ContextMenueType.DockPanel))

                Next

            Case ContextMenueType.TrayIcon

                'トレイアイコン用のコンテキストメニューアイテムのインスタンスを作成
                mContextMenu = _CreateContextMenuInstance(ContextMenueType.TrayIconItem)

                'トレイアイコン用のコンテキストメニュー名をセット
                mContextMenu.Text = pMenuName

                'リスト数分繰り返す（階層メニューの項目を追加していく）
                For Each mItemName As String In pMenuList

                    mContextMenu.DropDownItems.Add(_CreateContextMenuNormalItem(mItemName, ContextMenueType.TrayIcon))

                Next

            Case Else

                mContextMenu = Nothing

        End Select

        Return mContextMenu

    End Function

#End Region

#Region "コンテキストメニュー状態変更関連"

    ''' <summary>コンテキストメニュー左横のチェックマークをセットする</summary>
    ''' <param name="pContextMenu">コンテキストメニュー</param>
    ''' <param name="pMenuType">コンテキストメニュータイプ</param>
    ''' <param name="pItemId">コンテキストメニューアイテムID</param>
    ''' <param name="pCheckStatus">チェック状態（True：チェック、False：チェックなし）</param>
    ''' <remarks>
    '''   コンテキストメニュータイプによりチェックが付けられるプロパティが違うためそれぞれで行う
    ''' </remarks>
    Private Sub _SetCheckMarkForContextMenu(ByVal pContextMenu As Object, ByVal pMenuType As ContextMenueType, ByVal pItemId As Integer, ByVal pCheckStatus As Boolean)

        Select Case pMenuType

            Case ContextMenueType.DockPanel, ContextMenueType.DockPanelItem

                DirectCast(pContextMenu.Items(pItemId), System.Windows.Controls.MenuItem).IsChecked = pCheckStatus

            Case ContextMenueType.TrayIcon

                DirectCast(pContextMenu.Items(pItemId), System.Windows.Forms.ToolStripMenuItem).Checked = pCheckStatus

            Case ContextMenueType.TrayIconItem

                DirectCast(pContextMenu.DropDownItems(pItemId), System.Windows.Forms.ToolStripMenuItem).Checked = pCheckStatus

        End Select

    End Sub

    ''' <summary>コンテキストメニューのEnableをセットする</summary>
    ''' <param name="pContextMenu">コンテキストメニュー</param>
    ''' <param name="pMenuType">コンテキストメニュータイプ</param>
    ''' <param name="pItem">コンテキストメニューアイテム</param>
    ''' <param name="pEnabled">Enableの状態（True：使用可能、False：使用不可）</param>
    ''' <remarks>
    '''   コンテキストメニュータイプによりチェックが付けられるプロパティが違うためそれぞれで行う
    ''' </remarks>
    Private Sub _SetEnabledForContextMenu(ByVal pContextMenu As Object, ByVal pMenuType As ContextMenueType, ByVal pItem As ContextMenuItem, ByVal pEnabled As Boolean)

        Select Case pMenuType

            Case ContextMenueType.DockPanel

                DirectCast(pContextMenu.Items(pItem), System.Windows.Controls.MenuItem).IsEnabled = pEnabled

            Case ContextMenueType.TrayIcon

                DirectCast(pContextMenu.Items(pItem), System.Windows.Forms.ToolStripMenuItem).Enabled = pEnabled

        End Select

    End Sub

#End Region

#Region "コンテキストメニュー情報取得関連"

    ''' <summary>クリックされたコンテキストメニュー名を取得</summary>
    ''' <param name="pSender">メニューオブジェクト</param>
    ''' <returns>クリックされたコンテキストメニュー名</returns>
    ''' <remarks>
    '''   メニューの型ごとでクリックされたメニューの名前の取得方法が異なるため、
    '''   型をチェックしその型にあった取得方法でクリックされたメニューの名前を取得する
    ''' </remarks>
    Private Function _GetClickContextMenuItemName(ByVal pSender As Object) As String

        Dim mClickItemName As String = String.Empty

        'メニューの型ごと処理を分岐
        Select Case _GetContextMenuItemType(pSender)

            Case ContextMenueType.DockPanel

                mClickItemName = DirectCast(pSender, System.Windows.Controls.MenuItem).Header

            Case ContextMenueType.TrayIcon

                mClickItemName = DirectCast(pSender, System.Windows.Forms.ToolStripMenuItem).Text

            Case Else

                mClickItemName = String.Empty

        End Select

        Return mClickItemName

    End Function

    ''' <summary>コンテキストメニューアイテムの型を取得</summary>
    ''' <param name="pSender">メニューオブジェクト</param>
    ''' <returns>コンテキストメニューアイテムの型</returns>
    ''' <remarks></remarks>
    Private Function _GetContextMenuItemType(ByVal pSender As Object) As ContextMenueType

        Select Case True

            Case TypeOf pSender Is System.Windows.Controls.MenuItem

                Return ContextMenueType.DockPanel

            Case TypeOf pSender Is System.Windows.Forms.ToolStripMenuItem

                Return ContextMenueType.TrayIcon

            Case Else

                Return ContextMenueType.Other

        End Select

    End Function

    ''' <summary>階層メニューアイテムタイプを取得</summary>
    ''' <param name="pMenuItem">メニューアイテム</param>
    ''' <returns>階層メニューのアイテムタイプ</returns>
    ''' <remarks>
    '''   階層メニューで無い時は「ContextMenueType.Other」を返します
    ''' </remarks>
    Private Function _GetHierarchyMenuItemType(ByVal pMenuItem As Object) As ContextMenueType

        Select Case True

            Case TypeOf pMenuItem Is System.Windows.Controls.MenuItem

                'メニューアイテムのアイテム数が０より大きい時（階層メニューの時）
                If DirectCast(pMenuItem, System.Windows.Controls.MenuItem).Items.Count > 0 Then

                    Return ContextMenueType.DockPanelItem

                Else

                    Return ContextMenueType.Other

                End If

            Case TypeOf pMenuItem Is System.Windows.Forms.ToolStripMenuItem

                'ドロップダウンメニューアイテムのアイテム数が０より大きい時（階層メニューの時）
                If DirectCast(pMenuItem, System.Windows.Forms.ToolStripMenuItem).DropDownItems.Count > 0 Then

                    Return ContextMenueType.TrayIconItem

                Else

                    Return ContextMenueType.Other

                End If

            Case Else

                Return ContextMenueType.Other

        End Select

    End Function

#End Region

#Region "コンテキストメニューその他関連"

    ''' <summary>コンテキストメニューのインスタンスを作成</summary>
    ''' <param name="pMenuType">コンテキストメニュータイプ</param>
    ''' <returns>コンテキストメニューのインスタンス</returns>
    ''' <remarks>
    '''   コンテキストメニュータイプごとそれに対応したコンテキストメニューのインスタンスを返す
    ''' </remarks>
    Private Function _CreateContextMenuInstance(ByVal pMenuType As ContextMenueType) As Object

        Dim mInstanceObject As New Object

        'コンテキストメニュータイプごと処理を分岐
        Select Case pMenuType

            Case ContextMenueType.DockPanel

                'WPF用のコンテキストメニューのインスタンスをセット
                mInstanceObject = New System.Windows.Controls.ContextMenu

            Case ContextMenueType.TrayIcon

                'トレイアイコン用のコンテキストメニューのインスタンスをセット
                mInstanceObject = New System.Windows.Forms.ContextMenuStrip

            Case ContextMenueType.DockPanelItem

                'WPF用のコンテキストメニューアイテムのインスタンスをセット
                mInstanceObject = New System.Windows.Controls.MenuItem

            Case ContextMenueType.TrayIconItem

                'トレイアイコン用のコンテキストメニューアイテムのインスタンスをセット
                mInstanceObject = New System.Windows.Forms.ToolStripMenuItem

            Case ContextMenueType.Other

                'Nothingをセット
                mInstanceObject = Nothing

        End Select

        Return mInstanceObject

    End Function

    ''' <summary>ファイルの実行処理</summary>
    ''' <param name="pPath">実行ファイルのフルパス</param>
    ''' <remarks>
    '''   パスにURLを指定するとそのURLを規定のブラウザでします
    ''' </remarks>
    Private Sub _RunFile(ByVal pPath As String)

        Dim mPsi As New System.Diagnostics.ProcessStartInfo()

        '関連付けで実行するファイルのフルパスを指定
        mPsi.FileName = pPath

        'ファイルの実行処理
        System.Diagnostics.Process.Start(mPsi)

    End Sub

#End Region

#End Region

#Region "画像ファイル関連"

    ''' <summary>美人時計画像をダウンロード</summary>
    ''' <remarks></remarks>
    Private Async Sub _DownloadBijinTokeiImage()

        '-------------------------------
        ' ダウンロード前処理
        '-------------------------------
        '美人時計画像をダウンロード中に変更
        _IsDownloadingForBijinTokeiImages = True

        'トレイアイコンのテキストを無しに設定
        _NotifyIcon.Text = String.Empty

        'トレイアイコンのアイコンをダウンロード中のアイコンに変更
        _NotifyIcon.Icon = My.Resources.DownloadingImage

        '-------------------------------
        ' ダウンロードファイル情報を作成
        '-------------------------------
        Dim mDownloadInfos As New List(Of DownloadForNet.DownloadInfo)

        '時間のリスト分繰り返す
        For Each mTime As String In _CreateTimeList()

            'ダウンロード情報を追加
            mDownloadInfos.Add(New DownloadForNet.DownloadInfo(_GetImagePath(ImagePathKbn.DownloadPath, mTime) _
                                                             , _GetImagePath(ImagePathKbn.SavePath, mTime)))

        Next

        '現在時刻 ～ 画像のダウンロード数までのリストを作成（現在時刻 ～ 画像のダウンロード数以外のリストを削除する）
        mDownloadInfos.RemoveRange(Me.ImageDownloadCount - 1, _CreateTimeList.Count - Me.ImageDownloadCount)

        '-------------------------------
        ' ファイルのダウンロード
        '-------------------------------
        _DownloadForNet = New DownloadForNet(mDownloadInfos.ToArray _
                                           , New Progress(Of DownloadForNet.FileDownloadProgress)(AddressOf _SetTrayIconText))

        'ファイルのダウンロード処理（非同期）
        Await _DownloadForNet.DownloadFileAsync(mDownloadInfos.ToArray)

        '例外がNothingで無かった時（例外が発生した時）、例外内容をバルーンウィンドウに表示する
        If Not _DownloadForNet.Exception Is Nothing Then Call _SetBalloonWindowText(String.Format(cMessage.DownloadFailure, _DownloadForNet.Exception.Message) _
                                                                                  , BalloonWindowDisplayTime.DownloadComplete)

        _DownloadForNet = Nothing

        '-------------------------------
        ' ダウンロード後処理
        '-------------------------------
        '美人時計画像をダウンロード中でないに変更
        _IsDownloadingForBijinTokeiImages = False

        'トレイアイコンのテキスト設定
        Call _SetTrayIconText()

        'トレイアイコンのアイコンを通常のアイコンに戻す
        _NotifyIcon.Icon = My.Resources.App

    End Sub

    ''' <summary>画像ファイルパスを取得</summary>
    ''' <param name="pImagePathKbn">画像ファイルパス区分</param>
    ''' <param name="pNowTime">現在時刻</param>
    ''' <returns>画像ファイルパス区分</returns>
    ''' <remarks></remarks>
    Private Function _GetImagePath(ByVal pImagePathKbn As ImagePathKbn, Optional ByVal pNowTime As String = "0000") As String

        Dim mImagePath As String = String.Empty

        Select Case pImagePathKbn

            Case ImagePathKbn.SavePath

                '「画像ファイル保存パス / 現在時刻文字列 + .jpg」形式をセット
                mImagePath = Me.BijinTokei.GetImagePath(BijinTokeiInfo.ImagePathKbn.SavePath) & pNowTime & ".jpg"

            Case ImagePathKbn.DownloadPath

                '「美人時計画像格納URL / 現在時刻文字列 + .jpg」形式をセット
                mImagePath = Me.BijinTokei.TargetBijinTokei.ImageUrl & pNowTime & ".jpg"

            Case ImagePathKbn.NoImagePath

                'NoImage画像ファイルパスをセット
                mImagePath = Me.BijinTokei.GetImagePath(BijinTokeiInfo.ImagePathKbn.NoImagePath)

                'NoImage画像が存在しない時は例外を発生させる
                If Not File.Exists(mImagePath) Then Throw New Exception(cMessage.NotExistsNoImage)

        End Select

        Return mImagePath

    End Function

    ''' <summary>ファイルが存在しているかまたは使用しているか</summary>
    ''' <param name="pFilePath">ファイルのフルパス</param>
    ''' <returns>True：ファイルが存在しないまたはファイルが使用中、False：ファイルが存在しファイルが使用中でない</returns>
    ''' <remarks>
    '''   指定されたファイルを開いてみて開けるかどうかにより使用中かそうでないか判断する
    '''   指定されたファイルが存在しない時は例外が発生するので返り値はTrueになります（使用中）
    ''' </remarks>
    Private Function _IsNotExistsOrUsingFile(ByVal pFilePath As String) As Boolean

        Dim mFile As New FileInfo(pFilePath)

        '「ファイルが存在しなかった」 または 「ファイルサイズが0バイト」の時、画像ファイルの存在なし
        If Not mFile.Exists OrElse mFile.Length = 0 Then Return True

        mFile = Nothing

        Try

            '対象パスのファイルを読み取り/書き込みアクセスで開く
            Using mStream As FileStream = New FileStream(pFilePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)

                Return False

            End Using

        Catch ex As Exception

            Return True

        End Try

    End Function

    ''' <summary>美人時計画像をセットする</summary>
    ''' <remarks></remarks>
    Private Sub _SetImageBijinTokei()

        '画像ファイルのフルパスを取得
        Dim mImagePath As String = _GetImagePath(ImagePathKbn.SavePath, _GetNowTime(TimeFormat.HourAndMinute))

        '「表示対象ファイル存在しない」または「使用中の時」は、NoImage画像に変更
        If _IsNotExistsOrUsingFile(mImagePath) Then mImagePath = _GetImagePath(ImagePathKbn.NoImagePath)

        '美人時計画像の切り替えが可の時
        If _CanChangeBijinTokeiImage(mImagePath) Then

            'BitmapImageを作成
            Dim mBitmap As New BitmapImage
            mBitmap.BeginInit()
            mBitmap.UriSource = New Uri(mImagePath)
            mBitmap.CacheOption = BitmapCacheOption.OnLoad
            mBitmap.EndInit()

            '画像を美人時計画像表示コントロールにセット
            imgBijinTokei.Source = mBitmap

        End If

    End Sub

    ''' <summary>美人時計画像の切り替えができるか？</summary>
    ''' <param name="pImagePath">画像ファイルパス</param>
    ''' <returns>True：切り替えが可、False：切り替えが不可</returns>
    ''' <remarks></remarks>
    Private Function _CanChangeBijinTokeiImage(ByVal pImagePath As String) As Boolean

        '     美人時計画像表示コントロールのソースがNothingで無く
        'かつ 現在表示している画像パスと今回表示する画像パスが一緒の時(前回と同じ画像を表示している時)
        If Not imgBijinTokei.Source Is Nothing _
        AndAlso DirectCast(imgBijinTokei.Source, BitmapImage).UriSource.LocalPath = pImagePath Then

            Return False

        Else

            Return True

        End If

    End Function

#End Region

#End Region

End Class
