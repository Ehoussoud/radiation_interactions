VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Donnees_initiales 
   Caption         =   "Donnees Initiales"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4725
   OleObjectBlob   =   "Donnees_initiales.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Donnees_initiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

    If VBA.IsNumeric(Me.TextBox1.Value) = False Or Me.TextBox1.Value <= 0 Or Me.TextBox1.Value > 2500 Then
        MsgBox "Entrer une Energie correcte", vbCritical
        Exit Sub
    End If
    
     If VBA.IsNumeric(Me.TextBox2.Value) = False Then
        MsgBox "Entrer une valeur de Gammas correcte", vbCritical
        Exit Sub
    End If
    
     If VBA.IsNumeric(Me.TextBox3.Value) = False Then
        MsgBox "Entrer une valeur de Z correcte", vbCritical
        Exit Sub
    End If
    
     If VBA.IsNumeric(Me.TextBox4.Value) = False Then
        MsgBox "Entrer une valeur de A correcte", vbCritical
        Exit Sub
    End If
    
     
     If Me.ComboBox1.Value = "" Then
        MsgBox "Preciser un etat d'ecran de protection valable", vbCritical
        Exit Sub
     End If
    
 Unload Donnees_initiales
 
Dim Dist As Double
 
 
    
    Dim Energie, Gammas, E  As Double
    Dim Z, Z1, A As Integer
    Dim Sim As Worksheet
    
    Set Sim = Worksheets("Simulation")
    
 
 
    Dim m, C, pe, d1, d2, dt, dm, dm1, x, Distepb As Double
    Dim Elt As Integer

    d1 = 0
    d2 = 0
    dt = 0
    Dim Talb As Worksheet
    
    Dim comp, q0 As Double
    
    Dim y, j As Integer
    
    y = 0
   

Energie = Me.TextBox1.Value
Gammas = Me.TextBox2.Value
Z = Me.TextBox3.Value
A = Me.TextBox4.Value

Sim.Range("E6").Value = Energie
Sim.Range("E8").Value = Gammas
Sim.Range("E10").Value = Z
Z1 = Z
Sim.Range("E12").Value = A
    
Result = Me.ComboBox1.Value
    

    If Result = "Plomb" Then

        
        Zpb = InputBox("Entrez le Numero Atomique (Zpb)", "Numero Atomique")
        Sim.Range("E16").Value = Zpb


        Apb = InputBox("Entrez le nombre de masse (Apb)", "Nombre de masse ")
        Sim.Range("E18").Value = Apb

        Distepb = InputBox("Entrez l'épaisseur de l'écran de plomb (Distepb)", "épaisseur écran plomb")
        Sim.Range("E20").Value = Distepb
'     calcul de B et Gammas correspondant lorsqu'on ajoute un ecran
        Dim Bpb, Dpb, UOpb As Double
     
   
        Alpha = Energie / 511
        Dim CEpb As Double
        E = Energie
        If E >= 400 Or E <= 1000 Then
            CEpb = 0.21 * Log(E / 1000) + 0.399 'Mise en équation de CEpb
        Else
            If E >= 1500 Or E <= 10000 Then
                CEpb = 0.529 * Exp(-0.186 * (E / 1000)) 'Mise en équation de CEpb
    
            End If
        End If
    
        If E >= 400 Or E <= 1000 Then
            Dpb = 0.019 * (E / 1000) - 0.051 'Mise en équation de Dpb
        End If
        
        bt1 = (1 + Alpha) / (Alpha ^ 2)
        bt2 = 2 * (1 + Alpha) / (1 + 2 * Alpha)
        bt3 = (Log(1 + 2 * Alpha)) / (Alpha)
        bt4 = (Log(1 + 2 * Alpha)) / (2 * Alpha)
        bt5 = ((1 + 3 * Alpha)) / (1 + 2 * Alpha) ^ 2
    
        UOpb = (0.30052 * Zpb * (bt1 * (bt2 - bt3) + bt4 - bt5)) / Apb
        Bepb = 1 + (UOpb * Distepb * CEpb * Exp(UOpb * Dpb * Distepb))

    'calcul du coefficient d'atténuation linéique'
        Alpha = E / 511
        c1 = 2 * (1 + Alpha) ^ 2 / (Alpha ^ 2 * (1 + 2 * Alpha))
        c2 = -(1 + 3 * Alpha) / ((1 + 2 * Alpha) ^ 2)
        c3 = -((1 + Alpha) * (2 * Alpha ^ 2 - 2 * Alpha - 1)) / (Alpha ^ 2 * (1 + 2 * Alpha) ^ 2)
        c4 = -(4 * Alpha ^ 2) / (3 * (1 + 2 * Alpha) ^ 3)
        p1 = (1 + Alpha) / Alpha ^ 3
        p2 = -1 / (2 * Alpha)
        p3 = 1 / (2 * Alpha ^ 3)
        c5 = -(p1 + p2 + p3)
        c6 = Log(1 + 2 * Alpha)
        comppb = (0.30052 * Zpb * (c1 + c2 + c3 + c4 + c5 * c6)) / Apb

        UPB = comppb * 11.34

        
        Gammas = Gammas * Bepb * Exp(-UPB * Distepb)

    
 
    End If
    
 
    Alpha = Energie / 511
   

'    Calcul du coefficient d'attenuation CE

    Dim CE As Double
    E = Energie
    
    p = 2.5522598 - 3.168602 * (10 ^ (-3) * E) + 2.8136316 * (10 ^ (-3) * E) ^ 2
    p1 = -1.47044 * (10 ^ (-3) * E) ^ 3 + 0.47593 * (10 ^ (-3) * E) ^ 4 - 0.099893 * (10 ^ (-3) * E) ^ 5
    p2 = 0.01325747 * (10 ^ (-3) * E) ^ 6 - 0.0011075941 * (10 ^ (-3) * E) ^ 7
    p3 = 5.2429293 * 10 ^ (-5) * (10 ^ (-3) * E) ^ 8 - 1.0723935 * 10 ^ (-6) * (10 ^ (-3) * E) ^ 9

    CE = p + p1 + p2 + p3
    
    'CE = 1.12 * (E / 1000) ^ -0.498  'Mise en équation de CE
    
'    Sim.Range("L11").Value = CE
    
     
'    Calcul du facteur d'accumulation B

    Dim B, D, UO As Double
    
    
    U = 0.3166718 - 0.39482 * (10 ^ (-3) * E) + 0.27691 * (10 ^ (-3) * E) ^ 2
    V = -0.1364 * (10 ^ (-3) * E) ^ 3 + 0.04681 * (10 ^ (-3) * E) ^ 4 - 0.01075 * (10 ^ (-3) * E) ^ 5
    W = 0.001591296 * (10 ^ (-3) * E) ^ 6 - 1.4436448 * 10 ^ (-4) * (10 ^ (-3) * E) ^ 7
    W1 = 7.2737286 * 10 ^ (-6) * (10 ^ (-3) * E) ^ 8 - 1.555029 * 10 ^ (-7) * (10 ^ (-3) * E) ^ 9

    D = U + V + W + W1
      
    at1 = (1 + Alpha) / (Alpha ^ 2)
    at2 = 2 * (1 + Alpha) / (1 + 2 * Alpha)
    at3 = (Log(1 + 2 * Alpha)) / (Alpha)
    at4 = (Log(1 + 2 * Alpha)) / (2 * Alpha)
    at5 = ((1 + 3 * Alpha)) / (1 + 2 * Alpha) ^ 2
    
    UO = (0.30052 * Z * (at1 * (at2 - at3) + at4 - at5)) / A
    
'    Sim.Range("L17").Value = UO
    
    
    
     E = Energie
     
'    Interactions observees

' cette ligne definit une ligne ou le programme peut se deplacer grace a la syntaxe goto suivi du nom de cette ligne

ConditionE:


    
    If E >= 2000 Then 'debut de verification de la valeur de E
        Sim.Range("O4").Value = "Effet de creation de paire"
        Sim.Range("R10").Value = ""
        Sim.Range("R12").Value = ""
        Sim.Range("R15").Value = dt
    ElseIf E <= 100 Then
        If y = 0 Then
            Sim.Range("O4").Value = "Effet Photoelectrique"
        Else
            Sim.Range("O4").Value = "Effet Compton + Photoelectrique"
        End If
    '   Calcul de la dose absorbee en cas d'effet photoelectrique
        Elt = 5  ' Debut iteration tableau d'Allen B effet photo
        Set Talb = Worksheets("Tableau d'Allen Brodsky")
        Talb.Activate
        Do Until Z = CInt(Talb.Cells(Elt, 2).Value) Or Z <= CInt(Talb.Cells(Elt, 2).Value) Or Talb.Cells(Elt, 2).Value = ""
           Elt = Elt + 1 ' on increment Elt pour faire evoluer la cellule Talb.Cells (Elt,2) progressivement
        Loop ' fin de la boucle
        Select Case True ' Verifier si les conditions apres la syntaxe Case sont verifiees
            Case Z = CInt(Talb.Cells(Elt, 2).Value) ' ici on verifie si Z est egal a la valeur de cellule ligne 5 et plus et colonne 2 du tableau d'Allen B.
                 m = Talb.Cells(Elt, 3).Value ' si oui on prend la valeur de m dans la cellule de la colonne 3
                 C = Talb.Cells(Elt, 4).Value 'si oui on prend la valeur de C dans la cellule de la colonne 4
                 pe = C * E ^ -m
                 d1 = 1.6 * 10 ^ (-10) * Gammas * (E / 1000) * pe
                 Sim.Range("R10").Value = d1
                 Sim.Range("R12").Value = d2
                 dm = d1 + d2
                 Sim.Range("R15").Value = dm1
            Case Z1 <= CInt(Talb.Cells(Elt, 2).Value) ' ici on verifie si Z est inferieur a la valeur de cellule ligne 5 et plus et colonne 2 du tableau d'Allen B.
                
                If Z1 > CInt(Talb.Cells(5, 2).Value) Then ' si tel est le cas on verifie d'abord si Z est superieur a la valeur(6) de la 1ere cellule
                 m = (Talb.Cells(Elt - 1, 3).Value + Talb.Cells(Elt, 3).Value) / 2
                 C = (Talb.Cells(Elt - 1, 4).Value + Talb.Cells(Elt, 4).Value) / 2
                 pe = C * E ^ -m
                 d1 = 1.6 * 10 ^ (-10) * Gammas * (E / 1000) * pe
                 Sim.Range("R10").Value = d1
                 Sim.Range("R12").Value = d2
                 dm = d1 + d2
                 Sim.Range("R15").Value = dm1
                Else  ' sinon si Z est inferieur a la valeur (6) de la 1ere cellule le programme ne peut pas determiner Z (ex: Z=5<6)
                 Sim.Range("R10").Value = ""
                 Sim.Range("R12").Value = ""
                 Sim.Range("R15").Value = ""
                 MsgBox ("Z3 ne peut etre calcule a partir du tableau d'Allen Brodsky ")
                End If
            Case Talb.Cells(Elt, 2).Value = "" ' ici on suppose que le programme a verifie toutes les cellules de la colonne Elements et trouve la cellule vide jusque en bas
                Sim.Range("R10").Value = ""
                Sim.Range("R12").Value = ""
                Sim.Range("R15").Value = ""
                MsgBox ("Z2 ne peut etre calcule a partir du tableau d'Allen Brodsky ")
        End Select
    Else
    
   
        Sim.Range("O4").Value = "Effet Compton"
        Alpha = E / 511
        c1 = 2 * (1 + Alpha) ^ 2 / (Alpha ^ 2 * (1 + 2 * Alpha))
        c2 = -(1 + 3 * Alpha) / ((1 + 2 * Alpha) ^ 2)
        c3 = -((1 + Alpha) * (2 * Alpha ^ 2 - 2 * Alpha - 1)) / (Alpha ^ 2 * (1 + 2 * Alpha) ^ 2)
        c4 = -(4 * Alpha ^ 2) / (3 * (1 + 2 * Alpha) ^ 3)
        p1 = (1 + Alpha) / Alpha ^ 3
        p2 = -1 / (2 * Alpha)
        p3 = 1 / (2 * Alpha ^ 3)
        c5 = -(p1 + p2 + p3)
        c6 = Log(1 + 2 * Alpha)
        comp = (0.30052 * Z * (c1 + c2 + c3 + c4 + c5 * c6)) / A

        Elt = 5 ' Debut iteration tableau d'Allen B effet compton
        Set Talb = Worksheets("Tableau d'Allen Brodsky")
        Talb.Activate
        Do Until Z = CInt(Talb.Cells(Elt, 2).Value) Or Z <= CInt(Talb.Cells(Elt, 2).Value) Or Talb.Cells(Elt, 2).Value = ""
        
           Elt = Elt + 1
          
        Loop
        Select Case True
            Case Z1 = CInt(Talb.Cells(Elt, 2).Value)
           
                 m = Talb.Cells(Elt, 3).Value
                 C = Talb.Cells(Elt, 4).Value
                 pe = C * E ^ (-m)
                 q0 = Int(comp / (comp + pe))
                 d2 = d2 + 1.6 * 10 ^ (-13) * Gammas * E * comp
                 Sim.Range("R10").Value = d1
                 Sim.Range("R12").Value = d2
                 dm1 = d1 + d2
                 Sim.Range("R15").Value = dm
                 
            Case Z1 <= CInt(Talb.Cells(Elt, 2).Value)
           
               
                If Z1 > CInt(Talb.Cells(5, 2).Value) Then
                    m = (Talb.Cells(Elt - 1, 3).Value + Talb.Cells(Elt, 3).Value) / 2
                    C = (Talb.Cells(Elt - 1, 4).Value + Talb.Cells(Elt, 4).Value) / 2
                    pe = C * (E) ^ (-m)
                    q0 = Int(comp / (comp + pe))
                    d2 = d2 + 1.6 * 10 ^ (-13) * Gammas * E * comp
                    Sim.Range("R10").Value = d1
                    Sim.Range("R12").Value = d2
                    dm1 = d1 + d2
                    Sim.Range("R15").Value = dm1
                    
                Else
                    Sim.Range("R10").Value = ""
                    Sim.Range("R12").Value = ""
                    Sim.Range("R15").Value = ""
                    MsgBox ("Z1 ne peut etre calcule a partir du tableau d'Allen Brodsky ")
                    GoTo fin
                End If
            
        End Select
        
        y = y + 1
        Gammas = Gammas * (1 - q0)
        Set Rsj = Worksheets("RSJ")
        Rsj.Activate
        
        j = 1
        
calculE:

        If Rsj.Cells(j + 7, 5).Value <= ((1 + 2 * Alpha) / (9 + 2 * Alpha)) Then
        
           Z = j + 1
        
           x = 1 + 2 * Rsj.Cells(Z + 7, 5).Value * Alpha
        
           
            
            If Rsj.Cells(Z + 7, 5).Value <= 4 * (x - 1) / x ^ 2 Then
            
                E = E / x
            
            Else
            
                j = j + 1
                
                If j = 10000 Then
                
                    GoTo fin
                Else
                
                GoTo calculE
            
                
               End If
                
             End If
             
             
        Else
            
          Z = j + 1
          
          x = (1 + 2 * Alpha) / (1 + 2 * Alpha * Rsj.Cells(Z + 7, 5).Value)
          
          
          
          If 2 * Rsj.Cells(Z + 7, 5).Value <= (((1 - x) / Alpha + 1) ^ 2 + (1 / x)) Then
          
            E = E / x
            
          Else
          
            j = j + 1
            
                If j = 10000 Then
                
                    GoTo fin
                Else
                
                GoTo calculE
            
                End If
            
          End If
            
        End If
            
        
        GoTo ConditionE
    
    
    
    End If ' fin de verification de la valeur de E
   
     
fin:

    Sim.Activate
    
For Dist = 0 To 10

    B = 1 + (UO * Dist * CE * Exp(UO * D * Dist))
    
    dt = dm1 * B * Exp(-UO * Dist)
    Sim.Range("R17").Value = dt
    
    Sim.Cells(8 + Dist, 10).Value = dt
 
Next Dist


 
 
    
End Sub



Private Sub Energie_Click()

End Sub

Private Sub UserForm_Activate()

    With Me.ComboBox1
        .Clear
        .AddItem ""
        .AddItem "Aucun"
        .AddItem "Plomb"
                
    End With
    
     
    
    
End Sub

