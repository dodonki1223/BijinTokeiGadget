Option Explicit On

Imports BijinTokeiGadget.BijinTokeiDefinition
Imports System.Data

''' <summary>表示時計の情報を提供する</summary>
''' <remarks></remarks>
Public Class BijinTokeiInfo

#Region "定数"

    ''' <summary>美人時計情報を格納しているXMLファイル名</summary>
    ''' <remarks></remarks>
    Private _cBijinTokeiInfoXmlPath As String = "BijinTokeiInfo.xml"

    ''' <summary>公式の美人時計のID</summary>
    ''' <remarks></remarks>
    Private _cOfficialId As Integer = 0

#End Region

#Region "列挙体"

    ''' <summary>美人時計データを格納するDataTableのカラム列挙体</summary>
    Private Enum BijinTokeiColumns

        ''' <summary>ID</summary>
        Id

        ''' <summary>対象時計</summary>
        TargetTokei

        ''' <summary>対象時計URL</summary>
        Url

        ''' <summary>画像URL</summary>
        ImageUrl

        ''' <summary>画像保存パス</summary>
        ImageSavePath

        ''' <summary>NoImage画像パス</summary>
        NoImagePath

        ''' <summary>画像の幅</summary>
        ImageWidth

        ''' <summary>画像の高さ</summary>
        ImageHeight

        ''' <summary>最小サイズ率</summary>
        MinimumSizeRate

    End Enum

    ''' <summary>画像パス区分</summary>
    Public Enum ImagePathKbn

        ''' <summary>美人時計画像保存パス</summary>
        SavePath

        ''' <summary>NoImageパス</summary>
        NoImagePath

    End Enum

#End Region

#Region "変数"

    ''' <summary>美人時計情報が格納されているXMLファイルを読み込むためのDatTable</summary>
    ''' <remarks></remarks>
    Private _BijinTokeiInfo As DataTable

    ''' <summary>対象の美人時計のIDを保持する変数</summary>
    ''' <remarks></remarks>
    Private _TargetBijinTokeiId As Integer

    ''' <summary>EXEの実行パスを格納する変数</summary>
    ''' <remarks></remarks>
    Private _ExePath As System.IO.FileInfo

    ''' <summary>美人時計情報構造体</summary>
    Public Structure BijinTokeiInfoStorage

        ''' <summary>対象時計ID</summary>
        Public Id As Integer

        ''' <summary>対象時計</summary>
        Public TargetTokei As String

        ''' <summary>対象時計URL</summary>
        Public Url As String

        ''' <summary>画像URL</summary>
        Public ImageUrl As String

        ''' <summary>画像保存パス</summary>
        Public ImageSavePath As String

        ''' <summary>NoImage画像パス</summary>
        Public NoImagePath As String

        ''' <summary>画像の幅</summary>
        Public ImageWidth As Integer

        ''' <summary>画像の高さ</summary>
        Public ImageHeight As Integer

        ''' <summary>最小サイズ率</summary>
        Public MinimumSizeRate As Double

        ''' <summary>美人時計情報をセットする</summary>
        ''' <param name="pId">対象時計ID</param>
        ''' <param name="pTargetTokei">対象時計</param>
        ''' <param name="pUrl">対象時計URL</param>
        ''' <param name="pImageUrl">画像URL</param>
        ''' <param name="pImageSavePath">画像保存パス</param>
        ''' <param name="pNoImagePath">NoImage画像パス</param>
        ''' <param name="pImageWidth">画像の幅</param>
        ''' <param name="pImageHeight">画像の高さ</param>
        ''' <param name="pMinimumSizeRate">最小サイズ率</param>
        Public Sub SetBijinTokeiInfo(ByVal pId As Integer, ByVal pTargetTokei As String, ByVal pUrl As String,
                                     ByVal pImageUrl As String, ByVal pImageSavePath As String, ByVal pNoImagePath As String,
                                     ByVal pImageWidth As Integer, ByVal pImageHeight As Integer, ByVal pMinimumSizeRate As Double)

            Id = pId
            TargetTokei = pTargetTokei
            Url = pUrl
            ImageUrl = pImageUrl
            ImageSavePath = pImageSavePath
            NoImagePath = pNoImagePath
            ImageWidth = pImageWidth
            ImageHeight = pImageHeight
            MinimumSizeRate = pMinimumSizeRate

        End Sub

    End Structure

#End Region

#Region "プロパティ"

    ''' <summary>美人時計情報</summary>
    ''' <returns>美人時計情報</returns>
    ''' <remarks>
    '''   Xmlファイルからの読み込みは１回だけになるように美人時計情報格納DataTableが「Nothing」
    '''   かそうでないかで取得方法が違います
    ''' </remarks>
    Private ReadOnly Property BijinTokeiInfo As DataTable

        Get

            '美人時計情報が「Nothing」の時（まだ１度もこのプロパティが呼ばれていない時）
            If _BijinTokeiInfo Is Nothing Then

                '美人時計情報が格納されているXMLファイルのパスを取得（EXE実行パスの親フォルダ + \ + xmlファイル名）
                Dim mXmlPath As String = Me.RunExeFolderPath & "\" & _cBijinTokeiInfoXmlPath

                'XMLファイルの読み込み
                Using mXmlReader As New System.Xml.XmlTextReader(mXmlPath)

                    '美人時計情報をXMLファイルからDataTableにセット
                    _BijinTokeiInfo = New DataTable
                    _BijinTokeiInfo.ReadXml(mXmlReader)

                End Using

                Return _BijinTokeiInfo

            Else

                Return _BijinTokeiInfo

            End If

        End Get

    End Property

    ''' <summary>実行EXEのフォルダーパス</summary>
    ''' <returns>実行EXEのフォルダーパス</returns>
    ''' <remarks>
    '''   実行しているEXEのパスの取得が１回だけになるように実行しているEXEのパス変数が「Nothing」
    '''   かそうでないかで取得方法が違います
    ''' </remarks>
    Private ReadOnly Property RunExeFolderPath As String

        Get

            '実行パス情報が「Nothing」の時（まだ１度もこのプロパティが呼ばれていない時）
            If _ExePath Is Nothing Then

                '実行しているEXEのパスを取得
                _ExePath = New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)

                '実行しているEXEのフォルダーパスを返す
                Return _ExePath.DirectoryName

            Else

                '実行しているEXEのフォルダーパスを返す
                Return _ExePath.DirectoryName

            End If

        End Get

    End Property

    ''' <summary>対象の美人時計情報</summary>
    ''' <returns>対象の美人時計情報</returns>
    ''' <remarks>
    '''   Xmlファイルから読み込んだ美人時計情報を対象の美人時計情報構造体に変換して取得する
    ''' </remarks>
    Public ReadOnly Property TargetBijinTokei As BijinTokeiInfoStorage

        Get

            '対象の美人時計情報（ID）を美人時計情報（xmlファイルから読み込んだDataTable）に渡して美人時計情報構造体に変換する
            Dim mTargetBijinTokeiInfo As New BijinTokeiInfoStorage
            mTargetBijinTokeiInfo.SetBijinTokeiInfo(Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.Id),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.TargetTokei),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.Url),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.ImageUrl),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.ImageSavePath),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.NoImagePath),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.ImageWidth),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.ImageHeight),
                                                    Me.BijinTokeiInfo.Rows(_TargetBijinTokeiId)(BijinTokeiColumns.MinimumSizeRate))

            Return mTargetBijinTokeiInfo

        End Get

    End Property

    ''' <summary>美人時計のリスト</summary>
    ''' <returns>美人時計のリスト</returns>
    ''' <remarks>
    '''   Xmlファイルから読み込んだ美人時計情報から美人時計のリストをList(Of String)で返す
    ''' </remarks>
    Public ReadOnly Property TokeiList As List(Of String)

        Get

            Return Me.BijinTokeiInfo.AsEnumerable().Select(
                                                               Function(row) _
                                                                   row(BijinTokeiColumns.TargetTokei).ToString()
                                                          ).ToList()

        End Get

    End Property

#End Region

#Region "コンストラクタ"

    ''' <summary>コンストラクタ</summary>
    ''' <remarks>引数無しは外部に公開しない</remarks>
    Private Sub New()

    End Sub

    ''' <summary>コンストラクタ</summary>
    ''' <param name="pTargetTokeiId">対象の美人時計ID</param>
    ''' <remarks>引数付きのコンストラクタのみを公開</remarks>
    Public Sub New(ByVal pTargetTokeiId As Integer)

        _TargetBijinTokeiId = pTargetTokeiId

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>対象の美人時計が公式か</summary>
    ''' <param name="pId">対象の美人時計ID</param>
    ''' <returns>True：公式、False：公式でない</returns>
    ''' <remarks>引数付きのコンストラクタのみを公開</remarks>
    Public Function IsOfficial(ByVal pId As Integer) As Boolean

        'Xmlファイルから読み込んだ美人時計情報から引数の美人時計IDに一致するIDが公式の時
        If Me.BijinTokeiInfo.Rows(pId)(BijinTokeiColumns.Id) = _cOfficialId Then

            Return True

        Else

            Return False

        End If

    End Function

    ''' <summary>対象の美人時計が地方か</summary>
    ''' <param name="pId">対象の美人時計ID</param>
    ''' <returns>True：地方、False：地方でない</returns>
    ''' <remarks>引数付きのコンストラクタのみを公開</remarks>
    Public Function IsLocal(ByVal pId As Integer) As Boolean

        'Xmlファイルから読み込んだ美人時計情報から引数の美人時計IDに一致するIDが公式でない時
        If Me.BijinTokeiInfo.Rows(pId)(BijinTokeiColumns.Id) <> _cOfficialId Then

            Return True

        Else

            Return False

        End If

    End Function

    ''' <summary>画像のパスを取得</summary>
    ''' <param name="pPathKbn ">画像パス区分</param>
    ''' <returns>画像へのフルパス</returns>
    ''' <remarks>
    '''   画像パス区分に応じた画像へのフルパスを返す
    '''   「EXEの実行パス + 画像への相対パス」の形式で返す
    ''' </remarks>
    Public Function GetImagePath(ByVal pPathKbn As ImagePathKbn)

        Select Case pPathKbn

            Case ImagePathKbn.SavePath

                Return Me.RunExeFolderPath & Me.TargetBijinTokei.ImageSavePath

            Case ImagePathKbn.NoImagePath

                Return Me.RunExeFolderPath & Me.TargetBijinTokei.NoImagePath

            Case Else

                Return ""

        End Select

    End Function

#End Region

End Class
