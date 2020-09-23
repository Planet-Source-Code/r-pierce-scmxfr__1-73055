Attribute VB_Name = "Module1"
Global Const datasize As Integer = 8192
Global connected As Boolean
Global listening As Boolean

Type settings ' is datasize in length
flag1 As Byte ' d0=master/slave d1=fetch ext IP
dot As String * 16
nicknam As String * 24
port As Integer
ipadr(1 To 10) As String * 128 ' 10 addresses user can choose to check their external IP from
urlidx As Integer
tmeout As Byte
defsave As String * 128
spare As String * 6738
End Type

Global db0 As settings


