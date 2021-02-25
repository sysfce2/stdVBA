VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Implemented as a low-key implementation of selenium
'
'Methods
'  [ ] Navigate()
'  [ ] SendMessage()
'  [ ] EmbedAsControl(uf as userform)
'Properties
'  [ ] innerHTML [R/W]
'  [ ] Top
'  [ ] Left
'  [ ] Width
'  [ ] Height
'  [ ] AppMode
'  [ ] url
'Events
'  [ ] OnMessage  - Acts as a bridge between JS and VBA
'
Public Enum EBrowserType
  InternetExplorer
  Edge
  Chrome
  FireFox
  Safari
End Enum


Public Function Create(ByVal browser as EBrowserType) as stdWebBrowser

End Function
Public Sub Init()

End Sub

