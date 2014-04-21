Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private WithEvents m_oComboBox As ComboBox
Attribute m_oComboBox.VB_VarHelpID = -1

Public Property Set ComboBox(ByVal oComboBox As ComboBox)
    Set m_oComboBox = oComboBox
End Property

Private Sub m_oComboBox_GotFocus()
    'MsgBox "ComboEventHandler Working!"
    Call ComboKeyValueStore(Screen.ActiveControl)
End Sub