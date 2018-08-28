# 美人時計ガジェット

## 美人時計のサイトから画像をダウンロードし、ローカルにある１分ごとの画像を切り替えるだけのデスクトップガジェットです

### フォルダについての説明
- CreateBijinTokeiData
    - 美人時計ガジェットで読み込むためのxmlファイルを作成するプログラム
- BijinTokeiGadget
    - 美人時計ガジェット本体

### 起動方法
- フォルダ構成は下記のようにすること
```
BijinTokeiGadget
 ├ Image
 ｜  └ NoImages
 ｜      ├ NoImage_Local.jpg
 ｜      └ NoImage_Official.jpg
 ├ BijinTokeiGadget.exe
 └ BijinTokeiInfo.xml
```
- BijinTokeiGadget.exeを実行する
- ガジェットを右クリックまたは通知領域のアイコンを右クリック
    - 画像のダウンロード数を設定し画像のダウンロードすること

### その他
- [美人時計サイト](http://www.bijint.com/)
