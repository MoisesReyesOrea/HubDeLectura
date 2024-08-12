VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mId_book As Integer
Private mTitle As String
Private mYear As String
Private mDescription As String
Private mAuthor As String
Private mGenre As String
Private mEditorial As String
Private mCoverImg As String

Public Property Get id_book() As String
    id_book = mId_book
End Property

Public Property Let id_book(ByVal value As String)
    mId_book = value
End Property

Public Property Get Title() As String
    Title = mTitle
End Property

Public Property Let Title(ByVal value As String)
    mTitle = value
End Property

Public Property Get Year() As String
    Year = mYear
End Property

Public Property Let Year(ByVal value As String)
    mYear = value
End Property

Public Property Get Description() As String
    Description = mDescription
End Property

Public Property Let Description(ByVal value As String)
    mDescription = value
End Property

Public Property Get author() As String
    author = mAuthor
End Property

Public Property Let author(ByVal value As String)
    mAuthor = value
End Property

Public Property Get genre() As String
    genre = mGenre
End Property

Public Property Let genre(ByVal value As String)
    mGenre = value
End Property

Public Property Get Editorial() As String
    Editorial = mEditorial
End Property

Public Property Let Editorial(ByVal value As String)
    mEditorial = value
End Property

Public Property Get CoverImg() As String
    CoverImg = mCoverImg
End Property

Public Property Let CoverImg(ByVal value As String)
    mCoverImg = value
End Property
