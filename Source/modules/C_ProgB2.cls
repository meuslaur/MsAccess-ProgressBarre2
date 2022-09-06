VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ProgB2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Dev")
' ------------------------------------------------------
' Name     : C_ProgB2
' ------------------------------------------------------
' ------------------------------------------------------
' Purpose  : Simulation progress bar
' Author   : Laurent
' Sujet    :
' Objectif :
' Date     : 04/09/2022 - 17:55
' DateMod  :
' ------------------------------------------------------
'// ! IMPORTANT !
'//             Le l'enplacement du label de texte(lbl_M2Texte) définira la positions des autres labels (a partir de sont coin bas/droite).
'//
'//             le lable1 (lbl_M2Texte) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'GAUCHE'
'//             le lable1 (lbl_BarreDeb) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'GAUCHE'
'//             le lable1 (lbl_BarreFin) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'GAUCHE'
'//             Les 3 labels doivent avoir leurs point d'encrage verticale(VerticalAnchor) défini sur 'HAUT'
'//
'// Dans le cas contraire les labels ne se potionnerons pas correctement.
'//
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
    Private ThisFrm As Access.Form

    Private Const LBL_GAUCHE As String = "lbl_PrgB2Debut"
    Private Const LBL_DROITE As String = "lbl_PrgB2Fin"
    Private Const LBL_INFO   As String = "lbl_PrgB2Texte"
    Private Const LBL_ROTA   As String = "lbl_PrgB2Rotation"

    Private Const LBL_WIDTH As Integer = 7000       '// Taille des label barres.
    Private Const DECAL     As Integer = 75         '// Espace entre lbl gauche et lbl droite.
    Private Const LBL_HAUT  As Integer = 100
    Private Const TXT_INFO  As String = " %"        '// Info texte.
    Private Const CHAR_ROTA As String = 183         '// Charactère debut horloge.

    Private Type FrmLabels
         Gauche  As Access.Label
         NomG    As String
         Droite  As Access.Label
         NomD    As String
         Info    As Access.Label
         NomI    As String
         Rota    As Access.Label
         NomR    As String
    End Type
    Private LesLabels As FrmLabels

    Private Type PropLabels
        Width   As Integer
        Height  As Integer
        TxtInfo As String
    End Type
    Private ParamLabels As PropLabels

    Private Type ParamBoucle
        NbBoucle    As Integer              '// Definir le nb total à parcourir, initialise les incréments.
        Increment   As Integer              '// Incrément décalage barre 1.
        Avance      As Double               '// Valeur  pourcentage avancement.
        CptRota     As Integer              '// Compteur Char rotation.
    End Type
    Private LaBoucle As ParamBoucle

    Private m_Fraction  As Double           '// Faction pourcentage avancement.
    Private m_Init      As Boolean
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::

'//----------------------------------       INIT/TERM       ------------------------------
Private Sub Class_Initialize()

    ParamLabels.Width = LBL_WIDTH
    ParamLabels.Height = LBL_HAUT
    ParamLabels.TxtInfo = TXT_INFO
    LesLabels.NomG = LBL_GAUCHE
    LesLabels.NomD = LBL_DROITE
    LesLabels.NomI = LBL_INFO
    LesLabels.NomR = LBL_ROTA

End Sub
Private Sub Class_Terminate()
    ResetLabels
End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================
Public Property Let NomLabelGauche(ByVal LabelNom As String)
    LesLabels.NomG = LabelNom
End Property
Public Property Let NomLabelDroite(ByVal LabelNom As String)
    LesLabels.NomD = LabelNom
End Property
Public Property Let NomLabelInfo(ByVal LabelNom As String)
    LesLabels.NomI = LabelNom
End Property
Public Property Let NomLabelRota(ByVal LabelNom As String)
    LesLabels.NomR = LabelNom
End Property

Public Property Let LabelWidth(ByVal NewLabelWidth As Integer)
    If ((NewLabelWidth < 3000) Or (NewLabelWidth > 9000)) Then Exit Property    '// Pour démo.
    ParamLabels.Width = NewLabelWidth
End Property

Public Property Let LabelHeight(NewLabelHeight As Integer)
    If ((NewLabelHeight < 50) Or (NewLabelHeight > 500)) Then Exit Property     '// Pour démo.
    ParamLabels.Height = NewLabelHeight
End Property
'//====================================== END PROP =======================================

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Sub InitLabels(FrmLabels As Access.Form, NbBoucles As Integer)
On Error GoTo ERR_InitLabels

    LaBoucle.NbBoucle = Nz(NbBoucles, 0)
    If (LaBoucle.NbBoucle < 10) Then MsgBox "10 boucles minimum": Exit Sub

    '// Initialisation
    Set ThisFrm = FrmLabels
    Set LesLabels.Gauche = ThisFrm.Controls(LesLabels.NomG)
    Set LesLabels.Droite = ThisFrm.Controls(LesLabels.NomD)
    Set LesLabels.Info = ThisFrm.Controls(LesLabels.NomI)
    Set LesLabels.Rota = ThisFrm.Controls(LesLabels.NomR)

    '// Positionnement des labels par rapport au label Info 'lblInfo' (c'est lui qui défini la position des autres labels)
    LesLabels.Gauche.Width = 0
    LesLabels.Gauche.Height = ParamLabels.Height
    LesLabels.Gauche.Top = (LesLabels.Info.Top + LesLabels.Info.Height)
    LesLabels.Gauche.Left = LesLabels.Info.Left ' + (LesLabels.Info.Width + 50)   '// Décale légerement la barre du texte de 50.
    
    LesLabels.Droite.Width = ParamLabels.Width - DECAL
    LesLabels.Droite.Height = ParamLabels.Height
    LesLabels.Droite.Left = LesLabels.Gauche.Left
    LesLabels.Droite.Top = LesLabels.Gauche.Top

    LesLabels.Rota.Top = LesLabels.Info.Top - (LesLabels.Rota.Height / 4)
    LesLabels.Rota.Left = (LesLabels.Info.Left + ParamLabels.Width) - (LesLabels.Rota.Width / 2)     '// Positionne a la fin du label.

    '// Calul incrément pour dimentionner les labels.
    LaBoucle.Increment = (ParamLabels.Width / (LaBoucle.NbBoucle)) - 1  '// Déclage dus aux valeurs non entières (ex: 5.26)
    m_Fraction = 100 / LaBoucle.NbBoucle                                '// Texte pourcentage avancement.

    LesLabels.Gauche.Visible = True
    LesLabels.Droite.Visible = True
    LesLabels.Info.Visible = True
    LesLabels.Rota.Visible = True

    m_Init = True
    
SORTIE_InitLabels:
    Exit Sub

ERR_InitLabels:
    m_Init = False
   MsgBox Err.Number & vbCrLf & Err.Description
    Resume SORTIE_InitLabels
End Sub

Public Sub UpdateLabels()
    If (m_Init = False) Then Exit Sub

    Dim Wdr As Integer

    LesLabels.Rota.Caption = ChrW$(LaBoucle.CptRota + CHAR_ROTA)    '// Pour le fun ...
    LaBoucle.CptRota = LaBoucle.CptRota + 1                         '//     ""
    If (LaBoucle.CptRota > 11) Then LaBoucle.CptRota = 0            '//     ""

    Wdr = LesLabels.Gauche.Width + LaBoucle.Increment
    If (Wdr > ParamLabels.Width) Then Wdr = ParamLabels.Width       '// Evite débordement dus aux arrondis.
    LesLabels.Gauche.Width = Wdr

    LesLabels.Droite.Left = (LesLabels.Gauche.Left + LesLabels.Gauche.Width) + DECAL
    Wdr = ParamLabels.Width - (LesLabels.Gauche.Width)
    LesLabels.Droite.Width = Wdr
    
    LaBoucle.Avance = LaBoucle.Avance + m_Fraction
    LesLabels.Info.Caption = Format$(LaBoucle.Avance, "##0") & ParamLabels.TxtInfo   '// Affichage correct du pourcentage.

End Sub

Public Sub CleanLabels()
    If (m_Init = False) Then Exit Sub

    LesLabels.Gauche.Width = ParamLabels.Width
    LesLabels.Droite.Visible = False
    LesLabels.Rota.Visible = False
    LesLabels.Info.Caption = 100 & TXT_INFO
    DoEvents

    LaBoucle.Avance = 0
    LaBoucle.CptRota = 0
    LaBoucle.Increment = 0
    LaBoucle.NbBoucle = 0

    ResetLabels
    
    m_Init = False

End Sub
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'// ################################ PRIVATE SUB/FUNC ####################################
Private Sub ResetLabels()

    Set LesLabels.Gauche = Nothing
    Set LesLabels.Droite = Nothing
    Set LesLabels.Info = Nothing
    Set LesLabels.Rota = Nothing
    Set ThisFrm = Nothing

End Sub
'// ################################# END PRIV. SUB/FUNC #################################
