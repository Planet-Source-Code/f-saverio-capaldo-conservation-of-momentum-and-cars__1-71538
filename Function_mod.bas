Attribute VB_Name = "Module1"
Option Explicit

' Conservation of Momentum
' Original by Hou Xiong 21.08.2006

' A smaller ball with a smaller mass collides with
' a larger ball with a larger mass.  The reaction
' occurs just as expected using Conservation of
' Momentum.  Play around with the variables in Form_Load().

' Modified by F. S. Capaldo 12.12.2008

' Now consider speed variation for elasticity coefficient
' and deceleration for drag factors
' Data and results on form
' Balls autosize from mass value

Public Pig As Double

Public x1 As Single    'position
Public r1 As Single    'radius
Public v1 As Single    'velocity
Public Vc1 As Single   'var velocity
Public m1 As Single    'mass

Public x2 As Single
Public r2 As Single
Public v2 As Single
Public Vc2 As Single   'var velocity
Public m2 As Single

Public Coll As Integer 'Flag Collision
Public Tcol As Single  'Time after collision

Public vt1 As Single
Public vt2 As Single
Public B1 As Integer
Public B2 As Integer

Public Dec As Single 'deceleration for drag factor
Public Ec As Single  'Coefficient of elasticity
Public Xcol As Single
Public Ycol As Single

' Disegno veicolo
Public Veic_01(8, 1) As Single     'Matrice punti disegno veicolo
Public Veic_02(8, 1) As Single     'Matrice punti disegno veicolo
Public Veic_21(18, 1) As Single     'Matrice punti parti veicolo
Public Veic_22(18, 1) As Single     'Matrice punti parti veicolo

Public Sub Marca_Punto(Dove As Form, x As Single, y As Single, Col As Integer)

Dove.DrawWidth = 1
Dove.Line (x - 10, y)-(x + 10, y), QBColor(Col)
Dove.Line (x, y - 10)-(x, y + 10), QBColor(Col)
Dove.DrawWidth = 1

End Sub

Public Sub Punti_Auto_1(Veic() As Single, Sca As Single)

'Inserisce i punti relativi al disegno del contorno
'veicolo (4.50x1.80) rispetto all'origine baricentrica
'BARICENTRO PRESO A CAZZO E DA RIVEDERE

Dim I As Integer
Dim J As Integer

Veic(0, 0) = -0.5
Veic(0, 1) = 2.25
Veic(1, 0) = 0.5
Veic(1, 1) = 2.25
Veic(2, 0) = 0.9
Veic(2, 1) = 1.9
Veic(3, 0) = 0.9
Veic(3, 1) = -1.75
Veic(4, 0) = 0.8
Veic(4, 1) = -2.25
Veic(5, 0) = -0.8
Veic(5, 1) = -2.25
Veic(6, 0) = -0.9
Veic(6, 1) = -1.75
Veic(7, 0) = -0.9
Veic(7, 1) = 1.9
Veic(8, 0) = -0.5
Veic(8, 1) = 2.25

For I = 0 To 8
   For J = 0 To 1
      Veic(I, J) = Veic(I, J) * Sca
   Next J
Next I

End Sub

Public Sub Dis_Auto(Dove As Form, Xd As Single, Yd As Single, Veic() As Single, Altro() As Single, An_D As Single, Colore As Integer)

Dove.DrawWidth = 2

'Draw car from polar coordinates of center
'with rotation and color QBColor(Colore)
Dim I As Integer     'point of car
Dim x As Single      'x draw coordinates
Dim y As Single      'y draw coordinates
Dim X0 As Single
Dim Y0 As Single

'Calcolo delle coordinate del baricentro del veicolo
X0 = Xd
Y0 = Yd

'first car point
x = X0 + Veic(0, 0) * Cos(An_D) - Veic(0, 1) * Sin(An_D)
y = Y0 + Veic(0, 0) * Sin(An_D) + Veic(0, 1) * Cos(An_D)

'set pen
Dove.PSet (x, y), QBColor(Colore)

'car points
For I = 1 To 8
   x = X0 + Veic(I, 0) * Cos(An_D) - Veic(I, 1) * Sin(An_D)
   y = Y0 + Veic(I, 0) * Sin(An_D) + Veic(I, 1) * Cos(An_D)
   Dove.Line -(x, y), QBColor(Colore)
Next I

Dove.DrawWidth = 1

'other car caracteristics
X0 = Xd
Y0 = Yd

'Calcola il primo punto del vetro anteriore del veicolo
x = X0 + Altro(0, 0) * Cos(An_D) - Altro(0, 1) * Sin(An_D)
y = Y0 + Altro(0, 0) * Sin(An_D) + Altro(0, 1) * Cos(An_D)

'Posiziona la penna
Dove.PSet (x, y), QBColor(Colore)

'Calcola e disegna gli altri punti del vetro anteriore del veicolo
For I = 1 To 4
   x = X0 + Altro(I, 0) * Cos(An_D) - Altro(I, 1) * Sin(An_D)
   y = Y0 + Altro(I, 0) * Sin(An_D) + Altro(I, 1) * Cos(An_D)
   Dove.Line -(x, y), QBColor(Colore)
Next I

'Calcola il primo punto del tetto del veicolo
x = X0 + Altro(5, 0) * Cos(An_D) - Altro(5, 1) * Sin(An_D)
y = Y0 + Altro(5, 0) * Sin(An_D) + Altro(5, 1) * Cos(An_D)

'Posiziona la penna
Dove.PSet (x, y), QBColor(Colore)

'Calcola e disegna gli altri punti del tetto del veicolo
For I = 6 To 9
   x = X0 + Altro(I, 0) * Cos(An_D) - Altro(I, 1) * Sin(An_D)
   y = Y0 + Altro(I, 0) * Sin(An_D) + Altro(I, 1) * Cos(An_D)
   Dove.Line -(x, y), QBColor(Colore)
Next I

'Calcola il primo punto del vetro posteriore del veicolo
x = X0 + Altro(10, 0) * Cos(An_D) - Altro(10, 1) * Sin(An_D)
y = Y0 + Altro(10, 0) * Sin(An_D) + Altro(10, 1) * Cos(An_D)

'Posiziona la penna
Dove.PSet (x, y), QBColor(Colore)

'Calcola e disegna gli altri punti del vetro posteriore del veicolo
For I = 11 To 14
   x = X0 + Altro(I, 0) * Cos(An_D) - Altro(I, 1) * Sin(An_D)
   y = Y0 + Altro(I, 0) * Sin(An_D) + Altro(I, 1) * Cos(An_D)
   Dove.Line -(x, y), QBColor(Colore)
Next I

'Disegna le due linee di cofano
x = X0 + Altro(15, 0) * Cos(An_D) - Altro(15, 1) * Sin(An_D)
y = Y0 + Altro(15, 0) * Sin(An_D) + Altro(15, 1) * Cos(An_D)
Dove.PSet (x, y), QBColor(Colore)
x = X0 + Altro(16, 0) * Cos(An_D) - Altro(16, 1) * Sin(An_D)
y = Y0 + Altro(16, 0) * Sin(An_D) + Altro(16, 1) * Cos(An_D)
Dove.Line -(x, y), QBColor(Colore)

x = X0 + Altro(17, 0) * Cos(An_D) - Altro(17, 1) * Sin(An_D)
y = Y0 + Altro(17, 0) * Sin(An_D) + Altro(17, 1) * Cos(An_D)
Dove.PSet (x, y), QBColor(Colore)
x = X0 + Altro(18, 0) * Cos(An_D) - Altro(18, 1) * Sin(An_D)
y = Y0 + Altro(18, 0) * Sin(An_D) + Altro(18, 1) * Cos(An_D)
Dove.Line -(x, y), QBColor(Colore)

End Sub



Public Sub Centra_Form(Quale As Form)

   ' Posiziona il form al centro dello schermo
   Quale.Left = (Screen.Width - Quale.Width) / 2
   Quale.Top = (Screen.Height - Quale.Height) / 2

End Sub


Public Sub Punti_Auto_2(Altro() As Single, Sca As Single)

'Inserisce i punti relativi al disegno del contorno
'dei vetri anteriore e posteriore e del tetto per il
'veicolo (4.50x1.80) rispetto all'origine baricentrica
'BARICENTRO PRESO A CAZZO E DA RIVEDERE

Dim I As Integer
Dim J As Integer

'Vetro Anteriore
Altro(0, 0) = -0.75
Altro(0, 1) = 1
Altro(1, 0) = 0.75
Altro(1, 1) = 1
Altro(2, 0) = 0.7
Altro(2, 1) = 0.25
Altro(3, 0) = -0.7
Altro(3, 1) = 0.25
Altro(4, 0) = -0.75
Altro(4, 1) = 1
'Tetto
Altro(5, 0) = -0.7
Altro(5, 1) = 0.15
Altro(6, 0) = 0.7
Altro(6, 1) = 0.15
Altro(7, 0) = 0.7
Altro(7, 1) = -1.1
Altro(8, 0) = -0.7
Altro(8, 1) = -1.1
Altro(9, 0) = -0.7
Altro(9, 1) = 0.15
'Vetro posteriore
Altro(10, 0) = -0.7
Altro(10, 1) = -1.2
Altro(11, 0) = 0.7
Altro(11, 1) = -1.2
Altro(12, 0) = 0.75
Altro(12, 1) = -1.6
Altro(13, 0) = -0.75
Altro(13, 1) = -1.6
Altro(14, 0) = -0.7
Altro(14, 1) = -1.2
'linea cofano dx
Altro(15, 0) = -0.75
Altro(15, 1) = 1.1
Altro(16, 0) = -0.5
Altro(16, 1) = 2.1
'linea cofano sx
Altro(17, 0) = 0.75
Altro(17, 1) = 1.1
Altro(18, 0) = 0.5
Altro(18, 1) = 2.1


For I = 0 To 18
   For J = 0 To 1
      Altro(I, J) = Altro(I, J) * Sca
   Next J
Next I

End Sub


