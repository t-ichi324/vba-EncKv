# vba-EncKv

複数のデータをCollectionで管理して文字列へエンコード・デコードすることができます。
OpenArgsで複数のデータを登録したいときや、隠しフィールドに項目に保存しておきたいときに利用できます。


実行例：
```
'--- データのセット
Dim kv As New EncKv
kv.Item("id") = 123
kv.Item("sub_id") = 456
kv.Item("cate") = "user"
kv.Item("names") = "Tester Name1" & vbCrLf & "Tester Name2" & vbCrLf & "Tester Name2"

'--- エンコード
Dim encString As String
encString = kv.Encode()

Debug.Print "エンコードされた文字列: "
Debug.Print encString & vbCrLf

'--- デコード
Dim dec As New EncKv
Call dec.Decode(encString)

Debug.Print "デコードされたデータ: "
Debug.Print "      id: " & dec.Item("id")
Debug.Print "  sub_id: " & dec.Item("sub_id")
Debug.Print "    cate: " & dec.Item("cate")
Debug.Print "    name: " & vbCrLf & dec.Item("names")
```

実行結果：
```
エンコードされた文字列: 
id:123,sub_id:456,cate:user,names:Tester\sName1\c\lTester\sName2\c\lTester\sName2

デコードされたデータ: 
      id: 123
  sub_id: 456
    cate: user
    name: 
Tester Name1
Tester Name2
Tester Name2

```
