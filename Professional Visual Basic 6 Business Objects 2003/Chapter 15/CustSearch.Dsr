VERSION 5.00
Begin {90290CCD-F27D-11D0-8031-00C04FB6C701} CustSearch 
   ClientHeight    =   4110
   ClientLeft      =   1815
   ClientTop       =   1545
   ClientWidth     =   7215
   _ExtentX        =   12726
   _ExtentY        =   7250
   SourceFile      =   ""
   BuildFile       =   ""
   BuildMode       =   0
   TypeLibCookie   =   243
   AsyncLoad       =   0   'False
   id              =   "CustSearch"
   ShowBorder      =   -1  'True
   ShowDetail      =   -1  'True
   AbsPos          =   0   'False
   HTMLDocument    =   "CustSearch.dsx":0000
End
Attribute VB_Name = "CustSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Function cmdSearch_onclick() As Boolean
  PutProperty BaseWindow.Document, "SearchName", txtName.Value
  PutProperty BaseWindow.Document, "SearchPhone", txtPhone.Value
  BaseWindow.navigate "VideoDHTML_CustList.html"

End Function
