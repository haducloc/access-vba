' ===== Class Module: XError =====
Option Compare Database
Option Explicit

Private mErrNum  As Long
Private mErrDesc As String
Private mErrSrc  As String

Public Property Get ErrNum() As Long
    ErrNum = mErrNum
End Property

Public Property Let ErrNum(ByVal v As Long)
    mErrNum = v
End Property

Public Property Get ErrDesc() As String
    ErrDesc = mErrDesc
End Property

Public Property Let ErrDesc(ByVal v As String)
    mErrDesc = v
End Property

Public Property Get ErrSrc() As String
    ErrSrc = mErrSrc
End Property

Public Property Let ErrSrc(ByVal v As String)
    mErrSrc = v
End Property
