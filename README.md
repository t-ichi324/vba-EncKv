# vba-EncKv

複数のデータをCollectionで管理して文字列へエンコード・デコードすることができます。
Jsonほどではないけど、
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

デコードした文字列を含めることもできます。
実行例２：
```
'--- データのセット
Dim kv As New EncKv

Dim i As Integer
For i = 0 To 3
    Dim item As New EncKv
    item.item("id") = i
    item.item("name") = "User " & i
    
    kv.item(i) = item.Encode
Next

'--- エンコード
Dim encString As String
encString = kv.Encode()

Debug.Print "エンコードされた文字列: "
Debug.Print encString & vbCrLf

'--- デコード
Dim dec As New EncKv
Call dec.Decode(encString)

For i = 0 To dec.Count - 1
    Dim itemEnc As String: itemEnc = dec.item(i)
    Dim itemDec As New EncKv
    Call itemDec.Decode(itemEnc)
    
    Debug.Print "デコード: " & i
    Debug.Print "   id: " & itemDec.item("id")
    Debug.Print " name: " & itemDec.item("name")
Next
```

実行結果２：
```
エンコードされた文字列: 
0:id\10\0name\1User\\s0,1:id\11\0name\1User\\s1,2:id\12\0name\1User\\s2,3:id\13\0name\1User\\s3

デコード: 0
   id: 0
 name: User 0
デコード: 1
   id: 1
 name: User 1
デコード: 2
   id: 2
 name: User 2
デコード: 3
   id: 3
 name: User 3
```
