VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mId As Integer
Private mName As String
Private mLastName As String
Private mEmail As String
Private mPassword As String
Private mAge As Byte
Private mPhoto As String

' Getter y Setters Name
Public Property Get Id() As Integer
    Id = mId
End Property

Public Property Let Id(ByVal value As Integer)
    mId = value
End Property

' Getter y Setters Name
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property

' Getter y Setters LastName
Public Property Get LastName() As String
    LastName = mLastName
End Property

Public Property Let LastName(ByVal value As String)
    mLastName = value
End Property

' Getter y Setters Email
Public Property Get Email() As String
    Email = mEmail
End Property

Public Property Let Email(ByVal value As String)
    mEmail = value
End Property

' Getter y Setters Password
Public Property Get Password() As String
    Password = mPassword
End Property

Public Property Let Password(ByVal value As String)
    mPassword = value
End Property

' Getter y Setters Age
Public Property Get Age() As String
    Age = mAge
End Property

Public Property Let Age(ByVal value As String)
    mAge = value
End Property

' Getter y Setters Name
Public Property Get Photo() As String
    Photo = mPhoto
End Property

Public Property Let Photo(ByVal value As String)
    mPhoto = value
End Property




