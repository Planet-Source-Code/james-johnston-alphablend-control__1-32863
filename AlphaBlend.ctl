VERSION 5.00
Begin VB.UserControl AlphaBlend 
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   InvisibleAtRuntime=   -1  'True
   Picture         =   "AlphaBlend.ctx":0000
   PropertyPages   =   "AlphaBlend.ctx":03F9
   ScaleHeight     =   435
   ScaleWidth      =   435
   ToolboxBitmap   =   "AlphaBlend.ctx":0407
   Windowless      =   -1  'True
End
Attribute VB_Name = "AlphaBlend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Default Property Values:
Const m_def_Enabled = 0
Const m_def_Opacity = 128
'Property Variables:
Dim m_Enabled As Boolean
Dim m_Opacity As Integer

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    
    If m_Enabled = True Then Call SetOpacity(Me.Opacity) Else SetOpacity (255)
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Opacity() As Integer
Attribute Opacity.VB_Description = "This value sets the transparency of it's parent form. (0 = Transparent, 255 = Opaque)"
Attribute Opacity.VB_ProcData.VB_Invoke_Property = "General"
    Opacity = m_Opacity
End Property

Public Property Let Opacity(ByVal New_Opacity As Integer)

    If New_Opacity > 255 Then New_Opacity = 255
    If New_Opacity < 0 Then New_Opacity = 0
    
    m_Opacity = New_Opacity
    PropertyChanged "Opacity"
    
    If m_Enabled = True Then Call SetOpacity(Me.Opacity)
    
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
    m_Enabled = m_def_Enabled
    m_Opacity = m_def_Opacity
    
End Sub

Private Sub UserControl_Paint()
    UserControl.Height = 435
    UserControl.Width = 435
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Opacity = PropBag.ReadProperty("Opacity", m_def_Opacity)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Opacity", m_Opacity, m_def_Opacity)
End Sub

Public Sub SetOpacity(NewOpacity As Long)
    Dim Ret As Long
    'Set the window style to 'Layered'
    Ret = GetWindowLong(UserControl.Parent.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong UserControl.Parent.hWnd, GWL_EXSTYLE, Ret
    'Set the opacity of the layered window
    SetLayeredWindowAttributes UserControl.Parent.hWnd, 0, NewOpacity, LWA_ALPHA
End Sub
