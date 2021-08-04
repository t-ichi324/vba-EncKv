VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EncKv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------

Private mKeys As New Collection
Private mVals As New Collection

Private Const DELIMT_ROW As String = ","
Private Const DELIMT_KV As String = ":"

Private Sub Class_Terminate(): On Error GoTo ex: Set mKeys = Nothing: Set mVals = Nothing
ex: End Sub

Public Function Init(Optional ByVal encString = "") As Boolean
    If encString <> "" Then: Call Decode(encString): Else: Call Clear
End Function

Public Sub Clear(): Set mKeys = New Collection: Set mVals = New Collection:  End Sub
Public Function Count() As Integer: Count = mKeys.Count: End Function
Public Function IsExists(ByVal pKey As String) As Boolean: IsExists = (mIndexOf(pKey) <> -1): End Function

Public Function GetKeys() As Collection
    Dim cp As New Collection
    Dim v As Variant
    For Each v In mKeys
        Call cp.Add(v)
    Next
    Set GetKeys = cp
End Function

Public Function Remove(ByVal pKey As String) As Boolean
    Remove = False
    Dim i As Integer: i = mIndexOf(pKey)
    If i < 0 Then Exit Function
    Call mKeys.Remove(i)
    Call mVals.Remove(i)
    Remove = True
End Function

'//////////////////////////////////////
Public Property Get Item(ByVal pKey As String) As String
    Item = Me.GetValue(pKey, "")
End Property
Public Property Let Item(ByVal pKey As String, ByVal pVal As String)
    Call Me.SetValue(pKey, pVal)
End Property

Public Sub SetValue(ByVal pKey As String, ByVal pVal As String)
    If pKey = "" Then: Exit Sub
    Call Me.Remove(pKey)
    Call mKeys.Add(pKey)
    Call mVals.Add(pVal)
End Sub
Public Function GetValue(ByVal pKey As String, Optional defaultVal As String = "") As String
On Error GoTo ex:
    Dim i As Integer: i = mIndexOf(pKey)
    If i >= 0 Then
        GetValue = mVals(i + 1)
        Exit Function
    End If
ex:
GetValue = defaultVal
End Function

Public Function Encode() As String
On Error GoTo ex:
    Encode = ""
    Dim ret As String: ret = ""
    Dim i As Integer: i = 0
    Dim max As Integer: max = Me.Count
    For i = 1 To max
        If ret <> "" Then: ret = ret & DELIMT_ROW
        ret = ret & VEnc(CStr(mKeys(i))) & DELIMT_KV & VEnc(CStr(mVals(i)))
    Next
    Encode = ret
ex:
End Function
Public Function Decode(ByVal encString As String) As Boolean
On Error GoTo ex:
    Decode = False
    Call Me.Clear
    Dim rows() As String: rows = Split(encString, DELIMT_ROW)
    Dim r As Variant
    For Each r In rows
        Dim kv() As String: kv = Split(CStr(r), DELIMT_KV)
        Call mKeys.Add(Trim(VDec(kv(0))))
        If UBound(kv) > 0 Then
            Call mVals.Add(VDec(kv(1)))
        Else
            Call mVals.Add("")
        End If
    Next
    Decode = True
ex:
End Function


Private Function mIndexOf(ByVal pKey As String) As Integer
On Error GoTo ex:
    mIndexOf = -1
    If pKey = "" Then: Exit Function
    Dim ix As Integer: ix = -1
    Dim v As Variant
    For Each v In mKeys
        ix = ix + 1
        If CStr(v) = pKey Then: mIndexOf = ix: Exit Function
    Next
ex:
End Function
Private Function VEnc(ByVal str As String) As String
    str = Replace(str, Chr(0), "")
    str = Replace(str, "\", "\\")
    str = Replace(str, Chr(13), "\c")       'vbCr
    str = Replace(str, Chr(10), "\l")       'vbLf
    str = Replace(str, Chr(9), "\t")        'vbTab
    str = Replace(str, " ", "\s")           'スペース
    str = Replace(str, DELIMT_ROW, "\0")    ',
    str = Replace(str, DELIMT_KV, "\1")     ':
    VEnc = str
End Function
Private Function VDec(ByVal str As String) As String
    str = Replace(str, Chr(0), "")
    str = Replace(str, "\\", Chr(0))
    str = Replace(str, "\c", Chr(13))       'vbCr
    str = Replace(str, "\l", Chr(10))       'vbLf
    str = Replace(str, "\t", Chr(9))        'vbTab
    str = Replace(str, "\s", " ")           'スペース
    str = Replace(str, "\0", DELIMT_ROW)    ',
    str = Replace(str, "\1", DELIMT_KV)     ':
    str = Replace(str, Chr(0), "\")
    VDec = str
End Function

