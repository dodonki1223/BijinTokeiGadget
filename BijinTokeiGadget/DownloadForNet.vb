Option Explicit On

Imports System
Imports System.Net.Http
Imports System.IO
Imports System.Linq

''' <summary>ネットからファイルをダウンロードする機能を提供します</summary>
''' <remarks></remarks>
Public Class DownloadForNet

#Region "定数"

    ''' <summary>１度に実行するタスクの最大値</summary>
    ''' <remarks></remarks>
    Public Const cRunTaskMaxCount As Integer = 60

    ''' <summary>再ダウンロード実行待ち時間</summary>
    ''' <remarks>
    '''   サーバーに付加をかけないようにするために処理待ちを行う
    '''   ※３秒間処理待ちをする
    ''' </remarks>
    Public Const cReRunDownloadSleepTime As Integer = 3000

#End Region

#Region "変数"

    ''' <summary>ダウンロード情報</summary>
    ''' <remarks></remarks>
    Private _DownloadInfo() As DownloadInfo

    ''' <summary>ダウンロードの失敗したダウンロード情報</summary>
    ''' <remarks></remarks>
    Private _FailureDownloadInfos As New List(Of DownloadInfo)

    ''' <summary>ファイルがダウンロード中か</summary>
    ''' <remarks></remarks>
    Private _IsFileDownloading As Boolean

    ''' <summary>ダウンロードの状態</summary>
    ''' <remarks></remarks>
    Private _FileDownloadProgress As FileDownloadProgress

    ''' <summary>例外格納変数</summary>
    ''' <remarks></remarks>
    Private _Exception As Exception

    ''' <summary>ファイルダウンロードの進捗状況を報告する変数</summary>
    ''' <remarks>ファイルダウンロードの進捗状況報告用</remarks>
    Private _ProcessProgress As IProgress(Of FileDownloadProgress)

#End Region

#Region "構造体"

    ''' <summary>ダウンロード情報</summary>
    ''' <remarks>ダウンロードする情報を格納する</remarks>
    Public Structure DownloadInfo

        ''' <summary>ダウンロードファイルURL</summary>
        ''' <remarks></remarks>
        Private _FileUrl As String

        ''' <summary>ダウンロードファイルの保存先</summary>
        ''' <remarks></remarks>
        Private _SavePath As String

        ''' <summary>ダウンロードファイルURL</summary>
        ''' <remarks></remarks>
        Public Property FileUrl As String

            Get

                Return _FileUrl

            End Get

            Set(value As String)

                _FileUrl = value

            End Set

        End Property

        ''' <summary>ダウンロードファイルの保存先</summary>
        ''' <remarks></remarks>
        Public Property SavePath As String

            Get

                Return _SavePath

            End Get

            Set(value As String)

                _SavePath = value

            End Set

        End Property

        ''' <summary>コンストラクタ</summary>
        ''' <param name="pFileUrl">ファイルURL</param>
        ''' <param name="pSavePath">ファイル保存先</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pFileUrl As String, ByVal pSavePath As String)

            'ダウンロードファイルURLをセット
            _FileUrl = pFileUrl

            'ダウンロードファイルの保存先パスをセット
            _SavePath = pSavePath

        End Sub

    End Structure

    ''' <summary>ファイルダウンロード進捗率構造体</summary>
    ''' <remarks>ファイルダウンロードの進捗率の内容を保持する構造体</remarks>
    Public Structure FileDownloadProgress

        ''' <summary>処理ファイル</summary>
        ''' <remarks></remarks>
        Private _ProcessingFile As String

        ''' <summary>ダウンロードファイル数</summary>
        ''' <remarks></remarks>
        Private _DownloadFileCount As Integer

        ''' <summary>ダウンロード処理済みファイル数</summary>
        ''' <remarks>処理したファイル数をカウント（ダウンロード失敗も含む）</remarks>
        Private _DownloadProcessedCount As Integer

        ''' <summary>ダウンロード済みファイル数</summary>
        ''' <remarks>実際にダウンロード出来た数をカウント（ダウンロード失敗は含まない）</remarks>
        Private _DownloadedFileCount As Integer

        ''' <summary>進捗率プロパティ</summary>
        ''' <remarks></remarks>
        Public ReadOnly Property Percent As Integer

            Get

                '(ダウンロード済みファイル数 / ダウンロードファイル数) * 100 
                Return (_DownloadedFileCount / _DownloadFileCount) * 100

            End Get

        End Property

        ''' <summary>ダウンロード完了か</summary>
        ''' <remarks>
        '''   ダウンロード処理済みファイル数とダウンロードファイル数が同じ時は
        '''   ダウンロード処理の完了とする
        ''' </remarks>
        Public ReadOnly Property IsComplete As Boolean

            Get

                'ダウンロード処理済みファイル数とダウンロードファイル数が同じ時
                If _DownloadProcessedCount = _DownloadFileCount Then

                    Return True

                Else

                    Return False

                End If

            End Get

        End Property

        ''' <summary>対象ファイルプロパティ</summary>
        ''' <remarks></remarks>
        Public Property ProcessingFile As String

            Set(value As String)

                _ProcessingFile = value

            End Set
            Get

                Return _ProcessingFile

            End Get

        End Property

        ''' <summary>ダウンロードファイル数</summary>
        ''' <remarks></remarks>
        Public Property DownloadFileCount As Integer

            Set(value As Integer)

                _DownloadFileCount = value

            End Set

            Get

                Return _DownloadFileCount

            End Get

        End Property

        ''' <summary>ダウンロード処理済みファイル数</summary>
        ''' <remarks>処理したファイル数をカウント（ダウンロード失敗も含む）</remarks>
        Public Property DownloadProcessedCount As Integer

            Set(value As Integer)

                _DownloadProcessedCount = value

            End Set

            Get

                Return _DownloadProcessedCount

            End Get

        End Property

        ''' <summary>ダウンロード済みファイル数</summary>
        ''' <remarks>実際にダウンロード出来た数をカウント（ダウンロード失敗は含まない）</remarks>
        Public Property DownloadedFileCount As Integer

            Set(value As Integer)

                _DownloadedFileCount = value

            End Set

            Get

                Return _DownloadedFileCount

            End Get

        End Property

        ''' <summary>コンストラクタ</summary>
        ''' <param name="pDownloadFileCount">ダウンロードファイル数</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pDownloadFileCount As Integer)

            'ダウンロードファイル数をセット
            _DownloadFileCount = pDownloadFileCount

        End Sub

    End Structure

#End Region

#Region "プロパティ"

    ''' <summary>ダウンロードするファイル情報プロパティ</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property DownloadsFileInfo As DownloadInfo()

        Get

            Return _DownloadInfo

        End Get

    End Property

    ''' <summary>ファイルがダウンロード中か</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property IsFileDownloading As Boolean

        Get

            Return _IsFileDownloading

        End Get

    End Property

    ''' <summary>例外（ダウンロード中に例外が発生した時にエラー内容を取得）</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Exception As Exception

        Get

            Return _Exception

        End Get

    End Property

    ''' <summary>処理進捗プロパティ</summary>
    ''' <remarks>
    '''   呼び出し元の画面から渡されてくる
    ''' </remarks>
    Public WriteOnly Property ProcessProgress As IProgress(Of FileDownloadProgress)

        Set(value As IProgress(Of FileDownloadProgress))

            _ProcessProgress = value

        End Set

    End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks>引数無しは外部に公開しない</remarks>
    Private Sub New()

    End Sub

    ''' <summary>コンストラクタ</summary>
    ''' <param name="pDownloadInfo">ダウンロード情報</param>
    ''' <remarks>引数付きのコンストラクタのみを公開</remarks>
    Public Sub New(ByVal pDownloadInfo() As DownloadInfo)

        'ダウンロード情報を変数にセット
        _DownloadInfo = pDownloadInfo

    End Sub

    ''' <summary>コンストラクタ</summary>
    ''' <param name="pDownloadInfo">ダウンロード情報</param>
    ''' <param name="pProcessProgress">ファイルダウンロードの進捗状況の報告</param>
    ''' <remarks>引数付きのコンストラクタのみを公開</remarks>
    Public Sub New(ByVal pDownloadInfo() As DownloadInfo, ByVal pProcessProgress As IProgress(Of FileDownloadProgress))

        'ダウンロード情報を変数にセット
        _DownloadInfo = pDownloadInfo

        'ファイルダウンロードの進捗状況を報告する変数にセット
        _ProcessProgress = pProcessProgress

    End Sub

#End Region

#Region "メソッド"

#Region "ファイルダウンロード進捗報告関連"

    ''' <summary>ダウンロード状態を更新</summary>
    ''' <param name="pFileUrl">ファイルURL</param>
    ''' <remarks></remarks>
    Private Sub _UpdateFileDownloadProgress(ByVal pFileUrl As String)

        'ファイルのダウンロード済み数を更新（失敗は含まない）
        _FileDownloadProgress.DownloadedFileCount = _FileDownloadProgress.DownloadedFileCount + 1

        '対象ファイルを更新
        _FileDownloadProgress.ProcessingFile = pFileUrl

    End Sub

    ''' <summary>ファイルダウンロード進捗報告</summary>
    ''' <remarks></remarks>
    Private Sub _ReportProcessProgress()

        'ファイルダウンロードの進捗状況を報告する変数がNothingで無かった時
        If Not _ProcessProgress Is Nothing Then

            '進捗率を報告
            _ProcessProgress.Report(_FileDownloadProgress)

        End If

    End Sub

#End Region

#Region "非同期にファイルをダウンロードする"

    ''' <summary>非同期にファイルをダウンロードする</summary>
    ''' <param name="pHttpClient"></param>
    ''' <param name="pFileUrl">ファイルURL</param>
    ''' <param name="pSavePath">保存先パス</param>
    ''' <returns>Task</returns>
    ''' <remarks></remarks>
    Public Async Function DownloadFileAsync(ByVal pHttpClient As HttpClient, ByVal pFileUrl As String, ByVal pSavePath As String) As Task

        Using mResponse As HttpResponseMessage = Await pHttpClient.GetAsync(pFileUrl, HttpCompletionOption.ResponseHeadersRead)

            'GET要求がHTTP500内部サーバーエラーの時
            If mResponse.StatusCode = Net.HttpStatusCode.InternalServerError Then

                'ダウンロード失敗情報に失敗した情報をセットしダウンロード処理を終了
                _FailureDownloadInfos.Add(New DownloadInfo(pFileUrl, pSavePath))
                Exit Function

            Else

                'ファイルのダウンロード処理済みファイル数を更新
                _FileDownloadProgress.DownloadProcessedCount = _FileDownloadProgress.DownloadProcessedCount + 1

            End If

            '親ディレクトリが存在しなかったらフォルダを作成する
            Call _CreateFolder(pSavePath)

            Using mFileStream As FileStream = File.Create(pSavePath)

                Using mHttpStream As Stream = Await mResponse.Content.ReadAsStreamAsync()

                    '非同期でファイルを書き込む
                    Await mHttpStream.CopyToAsync(mFileStream)

                    'ファイルダウンロードの進捗状況を更新（ダウンロード済みファイル数、対象ファイル）
                    Call _UpdateFileDownloadProgress(pFileUrl)

                    'ファイルダウンロードの進捗状況を報告
                    Call _ReportProcessProgress()

#If DEBUG Then
                    Console.WriteLine("ダウンロード済み：" & pSavePath)
#End If

                End Using

            End Using

        End Using

    End Function

    ''' <summary>非同期にファイルをダウンロードする</summary>
    ''' <param name="pDownloadInfos">ダウンロード情報構造体</param>
    ''' <returns>Task</returns>
    ''' <remarks></remarks>
    Public Async Function DownloadFileAsync(ByVal pDownloadInfos As DownloadInfo()) As Task

        'ファイルをダウンロード中に変更
        _IsFileDownloading = True

        _FileDownloadProgress = New FileDownloadProgress(pDownloadInfos.Count)

        Dim mTasks As New List(Of Task)

        Using mHttpClient As HttpClient = New HttpClient

            'ダウンロード情報分繰り返す
            For Each mDownloadInfo As DownloadInfo In pDownloadInfos

                'タスクを格納するListに対象ダウンロード情報のタスクを追加
                mTasks.Add(DownloadFileAsync(mHttpClient, mDownloadInfo.FileUrl, mDownloadInfo.SavePath))

            Next

            Try

                '失敗したダウンロード情報をすべて削除
                _FailureDownloadInfos.Clear()

                '----------------------------------------
                ' ダウンロードタスクを分割して実行
                '----------------------------------------
                '非同期で実行するタスクの実行回数を取得（小数点以下は切り捨て）
                Dim mRepeatCount As Integer = Math.Truncate(pDownloadInfos.Count / cRunTaskMaxCount)

                '非同期で実行するタスクの実行回数分繰り返す
                For i As Integer = 0 To mRepeatCount

                    'タスク実行範囲の始めを取得
                    Dim mRangeStart As Integer = i * cRunTaskMaxCount

                    'タスク実行範囲の終わりを取得
                    '※「タスク実行範囲の始め」＋「タスク実行範囲の終わり」がダウンロード情報のカウント数を超えていた時は
                    '  「ダウンロード情報のカウント数」ー「タスク実行範囲の始め」セット
                    Dim mRangeEnd As Integer = cRunTaskMaxCount
                    If mRangeStart + cRunTaskMaxCount >= pDownloadInfos.Count Then mRangeEnd = pDownloadInfos.Count - mRangeStart

                    '全てのタスクが終了してから完了するタスクを作成（非同期で行う）
                    Await Task.WhenAll(mTasks.GetRange(mRangeStart, mRangeEnd).ToArray)

                Next

                ''全てのタスクが終了してから完了するタスクを作成（非同期で行う）※１回のタスクで実行するパターン
                'Await Task.WhenAll(mTasks.ToArray)

                '----------------------------------------
                ' 失敗したダウンロードタスクを再実行する
                '----------------------------------------
                '失敗したダウンロード情報が０より多い時
                While (_FailureDownloadInfos.Count > 0)

                    '再ダウンロード前に処理待ちをする
                    System.Threading.Thread.Sleep(cReRunDownloadSleepTime)

                    '失敗したダウンロードを再実行する
                    Await DownloadFileAsync(_FailureDownloadInfos.ToArray)

                End While

            Catch ex As Exception

                '「ダウンロード処理済みファイル数」に「ダウンロードファイル数」をセットする
                '※明示的にダウンロード処理完了とする
                _FileDownloadProgress.DownloadProcessedCount = _FileDownloadProgress.DownloadFileCount

                '一番最初の例外のみをキャッチしてセットする
                '参考URL:https://www.kekyo.net/2015/06/22/5119
                _Exception = ex

                '------------------------
                ' イベントログに書き込み
                '------------------------
                'イベントログに表示するメッセージを取得
                Dim mEventLogMessage As String = Application.GetErrorMessage(ex, Application.ErrorMessageType.EventLog)

                'イベントログに書き込む
                Call Application.WriteEventLog(ex, mEventLogMessage, EventLogEntryType.Error)

            End Try

        End Using

        'ファイルをダウンロード中でないに変更
        _IsFileDownloading = False

    End Function

#End Region

#Region "フォルダ作成"

    ''' <summary>フォルダの作成</summary>
    ''' <param name="pFilePath">ファイルパス</param>
    ''' <remarks></remarks>
    Private Sub _CreateFolder(ByVal pFilePath As String)

        'ファイルの親ディレクトリ情報を取得
        Dim mDi As DirectoryInfo = New DirectoryInfo(pFilePath)
        Dim mParentFolder As DirectoryInfo = mDi.Parent

        '親ディレクトリが存在しなかったらフォルダを作成する
        If Not mParentFolder.Exists Then mParentFolder.Create()

    End Sub

#End Region

#End Region

End Class
