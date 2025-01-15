VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProMacro01 
   Caption         =   "Выравнивание объектов"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   OleObjectBlob   =   "ProMacro01.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ProMacro01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'    VBA MACRO For CORELDRAW / Alignement shapes with options
'    Copyright (C) 2020 Fabrice VAN NEER
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
  

Dim Compte As Long
Dim ShapeDejaAlignйes(2000) As Long
  
Private Sub Label11_Click()
    VBA.CreateObject("WScript.Shell").Run "https://corelnaveia.com"
End Sub

Private Sub Label13_Click()
    VBA.CreateObject("WScript.Shell").Run "http://kafard62.free.fr"
End Sub

Private Sub Label14_Click()
    VBA.CreateObject("WScript.Shell").Run "https://github.com/elvin-nsk"
End Sub

Private Sub UserForm_Initialize()
      
    ' Fonction ajoutй le 24/10/16 pour rйcupйrer la version automatiquement depuis le GMS
    Me.Caption = Me.Caption & VersionMacroProMacro
    
    ' permet de rйcupйrer la bonne taille avec le bouton prйvu а cette effet
    ActiveDocument.Unit = cdrMillimeter

    ' Charge les derniers paramиtres utilisйs
    TextBox1.Value = GetSetting("PRO MACRO", "ALIGN AUTO", "Largeur", 610)
    TextBox2.Value = GetSetting("PRO MACRO", "ALIGN AUTO", "Hauteur", 305)
    TextBox3.Value = GetSetting("PRO MACRO", "ALIGN AUTO", "Espacement", 5)
    MargeValue.Value = GetSetting("PRO MACRO", "ALIGN AUTO", "Marges", 5)
    CheckBox1.Value = GetSetting("PRO MACRO", "ALIGN AUTO", "Cadre Autour", "False")
    Checkbox2.Value = GetSetting("PRO MACRO", "ALIGN AUTO", "OptimiZ", "False")
 
End Sub
  
Sub RoutineAlignement()
   
 Dim SR As ShapeRange
 Dim S As Shape
 Dim Espacement As Double
 Dim LargeurSupport As Double
 Dim HauteurSupport As Double
 Dim Temp As Long
 Dim ShTemp As Shape
 Dim ShTrouvй As Boolean
 Dim RecupIdShapePlacй As Long
 Dim CoinX As Double
 Dim CoinY As Double
 Dim Marge As Double
 
' effacement des cadres prйcйdents
'ActiveSelection.Shapes.FindShapes("CADRE").Delete
ActivePage.FindShapes("CADRE").Delete
 
 ' Je parametre...
Compte = 0
Set SR = ActiveSelectionRange
ActiveDocument.Unit = cdrMillimeter
Marge = Val(MargeValue.Value)
Espacement = Val(TextBox3.Value)
EspacementEntreSupport = 20
LargeurSupport = Val(TextBox1.Value) - (2 * Marge) ' 610 'mm
HauteurSupport = Val(TextBox2.Value) - (2 * Marge) ' 300 'mm
CoinX = SR.LeftX
CoinY = SR.TopY

' petite vйrif / si rien de selectionnй -> quitte
If SR.Count = 0 Then Exit Sub

Memorise ' les parametres saisis

'*****************************************************************
 Application.EventsEnabled = False
 ActiveDocument.BeginCommandGroup "Выравнивание объектов"
 Optimization = True 'True
 '*****************************************************************

Set SR = SR.ReverseRange

' si on veux optimiser les placements
'If Checkbox2.Value = True Then
' Trie les shapes par dimensions ( hauteurs )
SR.Sort "@Shape1.Height > @Shape2.Height"
'End If

' on dйfini la premiere shape comme rйfйrence 'FormeDeDemarrage"
Dim FormeDeDemarrage As Shape
Set FormeDeDemarrage = SR(1)

' on positionne la premiere shape en haut a gauche de la selection en cours
FormeDeDemarrage.TopY = CoinY
FormeDeDemarrage.LeftX = CoinX
  
' on dйfini la variable permettant de positionnement X des formes
Dim CalculTranslationX As Double
CalculTranslationX = FormeDeDemarrage.RightX + Espacement  ' FormeDeDemarrage.SizeWidth + FormeDeDemarrage.LeftX
' Boucle de positionnement
For i = 2 To SR.Count
     
    If VerifShapeDejaUtilisй(CLng(i)) = True Then GoTo Fin
     
   ' positionne sur le haut de la shape de demarrage
    SR(i).TopY = FormeDeDemarrage.TopY
      
   If Checkbox2.Value = True Then
    'DEBUT //////// ..... POUR OPTIMISATION ...... ////////
    Do
    ' Si la largeur totale avec la shape a venir est plus grande que la LargeurSupport alors...
    If (CalculTranslationX + Espacement + SR(i).SizeWidth) - FormeDeDemarrage.LeftX > LargeurSupport Then

    ' Tester la fonction de sous alignement
    RecupIdShapePlacй = SousAlignement(SR, CLng(i), CalculTranslationX, FormeDeDemarrage.TopY, LargeurSupport - (CalculTranslationX - FormeDeDemarrage.LeftX), FormeDeDemarrage.SizeHeight)
    
    'If RecupIdShapePlacй <> 0 Then SR(RecupIdShapePlacй).Fill.UniformColor.RGBAssign 0, 0, 200
    If RecupIdShapePlacй <> 0 Then CalculTranslationX = CalculTranslationX + Espacement + SR(RecupIdShapePlacй).SizeWidth
    
         MemoriseIndexTemporairement = RecupIdShapePlacй
        
        If RecupIdShapePlacй <> 0 Then
        
            ' si l'espace en dessous de la shape est assez grand alors....
            If FormeDeDemarrage.SizeHeight - SR(RecupIdShapePlacй).SizeHeight - Espacement > 0 Then
                 
                MemoriseLaLargeurTemporairement = SR(RecupIdShapePlacй).SizeWidth
                 ' execute la fonction de sous alignement
                 RecupIdShapePlacй = SousAlignement(SR, CLng(i), SR(RecupIdShapePlacй).LeftX, SR(RecupIdShapePlacй).BottomY - Espacement, SR(RecupIdShapePlacй).SizeWidth, FormeDeDemarrage.SizeHeight - SR(RecupIdShapePlacй).SizeHeight - Espacement)
                                 
                 'si il a sous-alignй correctement et qu'il y a encore de la place alors....
                 If RecupIdShapePlacй <> 0 Then
                 
                     CALCUL = SR(RecupIdShapePlacй).SizeWidth + Espacement
                     
                     Do
                     
                     If RecupIdShapePlacй <> 0 Then
                     
                         If MemoriseLaLargeurTemporairement - CALCUL >= 0 Then
                         
                             RecupIdShapePlacй = SousAlignement(SR, CLng(i), SR(RecupIdShapePlacй).RightX + Espacement, SR(RecupIdShapePlacй).TopY, MemoriseLaLargeurTemporairement - CALCUL, SR(RecupIdShapePlacй).SizeHeight)
                             
                             'If RecupIdShapePlacй <> 0 Then
                             If RecupIdShapePlacй = 0 Then Exit Do
                             
                             CALCUL = CALCUL + Espacement + SR(RecupIdShapePlacй).SizeWidth ' Espacement
                             'End If
                             Else
                             Exit Do
                         End If
                     End If
                     'Exit Do
                     Loop Until RecupIdShapePlacй = 0
                    
                    End If
                              
               End If
        End If
 
    Else
         Exit Do
    End If
    
    Loop Until MemoriseIndexTemporairement = 0 'RecupIdShapePlacй = 0
    ' FIN //////// ..... POUR OPTIMISATION ...... ////////
   End If
    
   ' Si la largeur totale avec la shape a venir est plus grande que la LargeurSupport alors...
    If (CalculTranslationX + Espacement + SR(i).SizeWidth) - FormeDeDemarrage.LeftX > LargeurSupport Then
        
        'FormeDeDemarrage.Fill.UniformColor.RGBAssign 250, 0, 0
        'SR(i).Fill.UniformColor.RGBAssign 0, 200, 0

          ' Si la hauteur totale avec la shape a venir est plus grande que la HauteurSupport alors...
           If CoinY - (FormeDeDemarrage.BottomY - Espacement - SR(i).SizeHeight) >= HauteurSupport Then
            
                If Checkbox2.Value = True Then
                'DEBUT //////// ..... POUR OPTIMISATION ...... ////////
                
                ' variables pour simplifier
                PosXtmp = FormeDeDemarrage.LeftX
                PosYtmp = FormeDeDemarrage.BottomY - Espacement
                LargeurTmp = LargeurSupport
                HauteurTmp = HauteurSupport - (CoinY - (FormeDeDemarrage.BottomY - Espacement))
                                
                RecupIdShapePlacй = SousAlignement(SR, CLng(i), CDbl(PosXtmp), CDbl(PosYtmp), CDbl(LargeurTmp), CDbl(HauteurTmp))
                'si il a sous-alignй correctement et qu'il y a encore de la place alors....
                      
                If RecupIdShapePlacй <> 0 Then
                    CALCUL = SR(RecupIdShapePlacй).SizeWidth + Espacement
                    Do
                    
                    If RecupIdShapePlacй <> 0 Then
                    
                        If LargeurTmp - CALCUL >= 0 Then
                          
                            RecupIdShapePlacй = SousAlignement(SR, CLng(i), CDbl(PosXtmp + CALCUL), CDbl(PosYtmp), CDbl(LargeurTmp - CALCUL), CDbl(HauteurTmp))
                            
                            'If RecupIdShapePlacй <> 0 Then
                            If RecupIdShapePlacй = 0 Then Exit Do
                            
                            CALCUL = CALCUL + Espacement + SR(RecupIdShapePlacй).SizeWidth ' Espacement
                            'End If
                            Else
                            Exit Do
                        End If
                    End If
                    'Exit Do
                    Loop Until RecupIdShapePlacй = 0
                   
                End If
                
               End If
                'DEBUT //////// ..... POUR OPTIMISATION ...... ////////
            
            ' si on а cochй la case pour crйer un cadre autour alors ....
            If CheckBox1.Value = True Then
              TraceCadre CoinX - Marge, CoinY + Marge, LargeurSupport + (2 * Marge), HauteurSupport + (2 * Marge)
            End If
        
            ' redйfini l'origine un peu plus loin...
            CoinX = CoinX + LargeurSupport + (Marge * 2) + EspacementEntreSupport
            SR(i).LeftX = CoinX
            SR(i).TopY = CoinY
             
            Else
          ' la shape va tout a gauche
            'SR(i).Fill.UniformColor.RGBAssign 250, 0, 0
             
            SR(i).LeftX = FormeDeDemarrage.LeftX
            ' et ensuite va а la ligne suivante
            SR(i).TopY = FormeDeDemarrage.BottomY - Espacement
            
           End If
 
        ' Attribution de la nouvelle shape de demarrage
         Set FormeDeDemarrage = SR(i)
        ' "Remise а zero" de la translation
        'SR(i).Fill.UniformColor.RGBAssign 250, 0, 0
         CalculTranslationX = FormeDeDemarrage.PositionX    ' FormeDeDemarrage.SizeWidth + Espacement
    Else
       ' Si la largeur totale avec la shape a venir est plus PETITE que la LargeurSupport alors...
           SR(i).LeftX = CalculTranslationX '
    
        If Checkbox2.Value = True Then
        'DEBUT //////// ..... POUR OPTIMISATION ...... ////////
           ' si l'espace en dessous de la shape est assez grand alors....
           If FormeDeDemarrage.SizeHeight - SR(i).SizeHeight - Espacement > 0 Then
                 
                ' execute la fonction de sous alignement
                RecupIdShapePlacй = SousAlignement(SR, CLng(i), SR(i).LeftX, SR(i).BottomY - Espacement, SR(i).SizeWidth, FormeDeDemarrage.SizeHeight - SR(i).SizeHeight - Espacement)
                                
                'si il a sous-alignй correctement et qu'il y a encore de la place alors....
                If RecupIdShapePlacй <> 0 Then
                
                    CALCUL = SR(RecupIdShapePlacй).SizeWidth + Espacement
                    
                    Do
                    
                    If RecupIdShapePlacй <> 0 Then
                    
                        If SR(i).SizeWidth - CALCUL >= 0 Then
                        
                            RecupIdShapePlacй = SousAlignement(SR, CLng(i), SR(RecupIdShapePlacй).RightX + Espacement, SR(RecupIdShapePlacй).TopY, SR(i).SizeWidth - CALCUL, SR(RecupIdShapePlacй).SizeHeight)
                            
                            'If RecupIdShapePlacй <> 0 Then
                            If RecupIdShapePlacй = 0 Then Exit Do
                            
                            CALCUL = CALCUL + Espacement + SR(RecupIdShapePlacй).SizeWidth ' Espacement
                            'End If
                            Else
                            Exit Do
                        End If
                    End If
                    'Exit Do
                    Loop Until RecupIdShapePlacй = 0
                   
                   End If
                             
              End If
        End If
 
            If VerifShapeDejaUtilisй(CLng(i)) = True Then GoTo Fin 'Exit For
 
    
    End If
    
     CalculTranslationX = CalculTranslationX + Espacement + SR(i).SizeWidth
 
Fin:

Next i

' si on а cochй la case pour crйer un cadre autour alors ....
If CheckBox1.Value = True Then
  TraceCadre CoinX - Marge, CoinY + Marge, LargeurSupport + (2 * Marge), HauteurSupport + (2 * Marge)
End If
  
SR.AddToSelection

'*****************************************************************
    Optimization = False
    Application.EventsEnabled = True
    ActiveWindow.Refresh
    ActiveDocument.EndCommandGroup
'*****************************************************************
    
 End Sub

Function VerifShapeDejaUtilisй(Index As Long) As Boolean
          
        If Compte = 0 Then VerifShapeDejaUtilisй = False: Exit Function
        
        For b = 1 To Compte ' UBound(ShapeDejaAlignйes) 'Compte
            
            If Index = ShapeDejaAlignйes(b) Then
                VerifShapeDejaUtilisй = True
                Exit Function 'GoTo Fin 'Exit For
            End If
             
        Next b
        
        VerifShapeDejaUtilisй = False

End Function

Function SousAlignement(SR As ShapeRange, IndexDeSR As Long, PosX As Double, PosY As Double, LargeurDispo As Double, HauteurDispo As Double) As Long

'Dim ShTemp As Shape
Dim Trouvй As Boolean
Trouvй = False
 
' boucle pour trouver une shape .....
For i = IndexDeSR + 1 To SR.Count
        
    ' si la largeur est assez grande....
    If SR(i).SizeWidth <= LargeurDispo Then
         ' si la hauteur est assez grande.....
        If SR(i).SizeHeight <= HauteurDispo Then
                  
                ' Si elle n'est pas dйjа utilisй, alors c'est bon
                If VerifShapeDejaUtilisй(CLng(i)) = False Then
                
                    ' positionne sur le haut de la shape de demarrage
                    SR(i).TopY = PosY
                    SR(i).LeftX = PosX '
                     
                    SousAlignement = i
                                    
                    ' memorisation de l'index de la shape sousalignй
                    Compte = Compte + 1
                    ShapeDejaAlignйes(Compte) = SousAlignement
                    
                    ' on quitte la fonction
                    Exit Function
                     
                End If
                 
                
        End If
    End If
    
Next i

SousAlignement = 0

End Function
 

Private Function VersionMacroProMacro() As String
End Function
  
Private Sub CommandButton1_Click()

    Dim b As Boolean
    Dim S As Shape
    Dim Shift As Long
    
    Dim DebutX As Double
    Dim DebutY As Double
    Dim FinX As Double
    Dim FinY As Double
     
     Me.Hide
        
      b = ActiveDocument.GetUserArea(DebutX, DebutY, FinX, FinY, Shift, 10, False, cdrCursorWinUpArrow)
      
      Me.Show
      
      If b = True Then Exit Sub
      
    
    TextBox1.Text = Abs(Fix(Val(FinX - DebutX)))
    TextBox2.Text = Abs(Fix(Val(FinY - DebutY)))

End Sub

Private Sub StartButton_Click()
       
    RoutineAlignement
        
End Sub
  
 Sub TraceCadre(PositionX As Double, PositionY As Double, Largeur As Double, Hauteur As Double)

    Dim Rect As Shape
    Set Rect = ActiveLayer.CreateRectangle2(PositionX, PositionY - Hauteur, Largeur, Hauteur)
    
    Rect.Name = "CADRE"

End Sub

Sub Memorise()

    SaveSetting "PRO MACRO", "ALIGN AUTO", "Largeur", Val(TextBox1.Value)
    SaveSetting "PRO MACRO", "ALIGN AUTO", "Hauteur", Val(TextBox2.Value)
    SaveSetting "PRO MACRO", "ALIGN AUTO", "Espacement", Val(TextBox3.Value)
    SaveSetting "PRO MACRO", "ALIGN AUTO", "Marges", Val(MargeValue.Value)
    
    If CheckBox1.Value = False Then Valeur = "False" Else Valeur = "True"
    SaveSetting "PRO MACRO", "ALIGN AUTO", "Cadre Autour", Valeur
    
    If Checkbox2.Value = False Then Valeur = "False" Else Valeur = "True"
    SaveSetting "PRO MACRO", "ALIGN AUTO", "OptimiZ", Valeur

End Sub
 
 
Private Sub UserForm_Terminate()

'Dim SR As ShapeRange
'Set SR = ActivePage.FindShapes("CADRE")

'If SR.Count = 0 Then Exit Sub
'For i = 1 To SR.Count
'    SR(i).Name = "CadreFixe"
'Next i

End Sub
