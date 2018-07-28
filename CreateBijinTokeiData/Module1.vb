Module Module1

#Region "定数"

    ''' <summary>美人時計定数クラス</summary>
    Private Class cBijinTokeiConst

        ''' <summary>美人時計テーブル名</summary>
        Public Const BijinTokeiTableName = "BijinTokeiInfo"

        ''' <summary>美人時計情報を格納したXMLファイル名</summary>
        Public Const XmlFileName = "BijinTokeiInfo.xml"

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

        ''' <summary>XMLファイルのエンコード</summary>
        Public Shared ReadOnly XmlEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("utf-8")

        ''' <summary>美人時計情報</summary>
        Public Shared ReadOnly BijinTokeiInfo()() As String = {
           New String() {0, "公式", "http://www.bijint.com/", "http://www.bijint.com/assets/pict/jp/pc/", "\Image\jp\", "\Image\NoImages\NoImage_Official.jpg", 960, 540, 3},
           New String() {1, "北海道", "http://www.bijint.com/hokkaido/", "http://www.bijint.com/assets/pict/hokkaido/pc/", "\Image\hokkaido\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {2, "青森", "http://www.bijint.com/aomori/", "http://www.bijint.com/assets/pict/aomori/pc/", "\Image\aomori\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {3, "宮城", "http://www.bijint.com/sendai/", "http://www.bijint.com/assets/pict/sendai/pc/", "\Image\sendai\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {4, "福島", "http://www.bijint.com/fukushima/", "http://www.bijint.com/assets/pict/fukushima/pc/", "\Image\fukushima\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {5, "埼玉", "http://www.bijint.com/saitama/", "http://www.bijint.com/assets/pict/saitama/pc/", "\Image\saitama\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {6, "千葉", "http://www.bijint.com/chiba/", "http://www.bijint.com/assets/pict/chiba/pc/", "\Image\chiba\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {7, "東京", "http://www.bijint.com/tokyo/", "http://www.bijint.com/assets/pict/tokyo/pc/", "\Image\tokyo\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {8, "静岡", "http://www.bijint.com/shizuoka/", "http://www.bijint.com/assets/pict/shizuoka/pc/", "\Image\shizuoka\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {9, "愛知", "http://www.bijint.com/nagoya/", "http://www.bijint.com/assets/pict/nagoya/pc/", "\Image\nagoya\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {10, "石川", "http://www.bijint.com/kanazawa/", "http://www.bijint.com/assets/pict/kanazawa/pc/", "\Image\kanazawa\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {11, "大阪", "http://www.bijint.com/osaka/", "http://www.bijint.com/assets/pict/osaka/pc/", "\Image\osaka\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {12, "奈良", "http://www.bijint.com/nara/", "http://www.bijint.com/assets/pict/nara/pc/", "\Image\nara\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {13, "岡山", "http://www.bijint.com/okayama/", "http://www.bijint.com/assets/pict/okayama/pc/", "\Image\okayama\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {14, "香川", "http://www.bijint.com/kagawa/", "http://www.bijint.com/assets/pict/kagawa/pc/", "\Image\kagawa\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {15, "宮崎", "http://www.bijint.com/miyazaki/", "http://www.bijint.com/assets/pict/miyazaki/pc/", "\Image\miyazaki\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {16, "鹿児島", "http://www.bijint.com/kagoshima/", "http://www.bijint.com/assets/pict/kagoshima/pc/", "\Image\kagoshima\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5},
           New String() {17, "沖縄", "http://www.bijint.com/okinawa/", "http://www.bijint.com/assets/pict/okinawa/pc/", "\Image\okinawa\", "\Image\NoImages\NoImage_Local.jpg", 590, 450, 2.5}
        }

        ''' <summary>美人時計情報をLISTで取得する</summary>
        ''' <returns>美人時計情報構造体のList</returns>
        Public Shared Function GetBijinTokeiInfo() As List(Of BijinTokeiInfoStorage)

            '美人時計情報格納用の配列を宣言する
            Dim mBijinTokeiInfoList As New List(Of BijinTokeiInfoStorage)

            '行データを作成
            For Each mBijinTokeiInfo As String() In cBijinTokeiConst.BijinTokeiInfo

                '美人時計の対象時計のデータを作成する
                Dim mBijinTokeiInfoStorage As New cBijinTokeiConst.BijinTokeiInfoStorage

                mBijinTokeiInfoStorage.SetBijinTokeiInfo(mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.Id),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.TargetTokei),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.Url),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.ImageUrl),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.ImageSavePath),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.NoImagePath),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.ImageWidth),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.ImageHeight),
                                                         mBijinTokeiInfo(cBijinTokeiConst.BijinTokeiColumns.MinimumSizeRate))

                mBijinTokeiInfoList.Add(mBijinTokeiInfoStorage)

            Next

            Return mBijinTokeiInfoList

        End Function

        ''' <summary>美人時計データを格納するDataTableのカラム列挙体</summary>
        Public Enum BijinTokeiColumns

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

    End Class

#End Region

#Region "メイン処理"

    ''' <summary>コンソールアプリケーションのメイン処理</summary>
    Sub Main()

        '美人時計DataTableのカラムを作成
        Dim mDt As DataTable = _CreateBijinTokeiColumns()
        mDt.TableName = cBijinTokeiConst.BijinTokeiTableName

        '行データを作成
        For Each mBijinTokeiInfo As cBijinTokeiConst.BijinTokeiInfoStorage In cBijinTokeiConst.GetBijinTokeiInfo

            mDt.Rows.Add(_AddRowDataToDataTable(mDt, mBijinTokeiInfo))

        Next

        'XMLファイルの出力先フルパスを作成 ※Exe実行パスの親フォルダ＋XMLファイル名
        Dim mExePath As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim mXmlOutputPath As String = mExePath.DirectoryName & "\" & cBijinTokeiConst.XmlFileName

        'XMLファイルの出力処理
        Using mWriter = New System.IO.StreamWriter(mXmlOutputPath, False, cBijinTokeiConst.XmlEncoding)

            mDt.WriteXml(mWriter, XmlWriteMode.WriteSchema, True)

        End Using

    End Sub

#End Region

#Region "メソッド"

    ''' <summary>
    '''   美人時計DataTableのカラムを作成する
    ''' </summary>
    ''' <returns>美人時計のカラム作成後のDataTable</returns>
    Private Function _CreateBijinTokeiColumns() As DataTable

        Dim mBijinTokei As New DataTable

        '美人時計カラム列挙体分繰り返す
        For Each mColumnName As String In System.Enum.GetNames(GetType(cBijinTokeiConst.BijinTokeiColumns))

            'カラムによりデータ型を設定する
            Select Case mColumnName

                Case cBijinTokeiConst.BijinTokeiColumns.Id.ToString, cBijinTokeiConst.BijinTokeiColumns.ImageWidth.ToString, cBijinTokeiConst.BijinTokeiColumns.ImageHeight.ToString

                    '「対象時計ID」、「画像の幅」、「画像の高さ」はInteger型で作成
                    mBijinTokei.Columns.Add(mColumnName, Type.GetType("System.Int32"))

                Case cBijinTokeiConst.BijinTokeiColumns.MinimumSizeRate.ToString

                    '「最小サイズ率」はDouble型で作成
                    mBijinTokei.Columns.Add(mColumnName, Type.GetType("System.Double"))

                Case Else

                    'デフォルトはString型で作成
                    mBijinTokei.Columns.Add(mColumnName, Type.GetType("System.String"))

            End Select

        Next

        Return mBijinTokei

    End Function

    ''' <summary>
    '''   コマンドの行データを作成
    ''' </summary>
    ''' <param name="pDataTable">行データ格納用のDataTable</param>
    ''' <param name="pBijinTokeiInfo">美人時計情報格納用の構造体</param>
    ''' <returns>コマンドの行データ</returns>
    Private Function _AddRowDataToDataTable(ByVal pDataTable As DataTable, ByVal pBijinTokeiInfo As cBijinTokeiConst.BijinTokeiInfoStorage) As DataRow

        Dim mAddDataRow As DataRow = pDataTable.NewRow

        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.Id) = pBijinTokeiInfo.Id
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.TargetTokei) = pBijinTokeiInfo.TargetTokei
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.Url) = pBijinTokeiInfo.Url
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.ImageUrl) = pBijinTokeiInfo.ImageUrl
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.ImageSavePath) = pBijinTokeiInfo.ImageSavePath
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.NoImagePath) = pBijinTokeiInfo.NoImagePath
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.ImageWidth) = pBijinTokeiInfo.ImageWidth
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.ImageHeight) = pBijinTokeiInfo.ImageHeight
        mAddDataRow(cBijinTokeiConst.BijinTokeiColumns.MinimumSizeRate) = pBijinTokeiInfo.MinimumSizeRate

        Return mAddDataRow

    End Function

#End Region

End Module
