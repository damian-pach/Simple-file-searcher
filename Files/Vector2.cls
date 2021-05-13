VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vector2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_x As Integer
Private m_y As Integer

Public Property Get x() As Integer

    x = m_x
    
End Property

Public Property Let x(x_val As Integer)

    m_x = x_val

End Property


Public Property Get y() As Integer

    y = m_y
    
End Property

Public Property Let y(y_val As Integer)

    m_y = y_val

End Property


Public Sub SetProperties(p_x As Integer, p_y As Integer)

    m_x = p_x
    m_y = p_y

End Sub

Public Function Magnitude() As Double

    Magnitude = Math.Sqr(m_x * m_x + m_y * m_y)

End Function

Public Function normalized() As Vector2
    
    Dim m_magnitude As Single
    m_magnitude = Magnitude()
    Dim m_x_norm As Single
    Dim m_y_norm As Single
    
    m_x_norm = m_x / m_magnitude
    m_y_norm = m_y / m_magnitude
        
    Set normalized = Factory.Vector2(m_x_norm, m_y_norm)

End Function
