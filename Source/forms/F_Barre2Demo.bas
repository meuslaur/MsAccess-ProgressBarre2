Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =15533
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =6288
    Top =3300
    Right =21816
    Bottom =6660
    RecSrcDt = Begin
        0x8ff0465d12e1e540
    End
    Caption ="Barre 2 démo,,,"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =60.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =536
            Name ="EntêteFormulaire"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =170
                    Top =56
                    Width =2004
                    Height =480
                    FontSize =18
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="Étiquette8"
                    Caption ="Démo Barre2"
                    GridlineColor =10921638
                    LayoutCachedLeft =170
                    LayoutCachedTop =56
                    LayoutCachedWidth =2174
                    LayoutCachedHeight =536
                End
            End
        End
        Begin Section
            Height =2834
            Name ="Détail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    Left =3798
                    Top =170
                    Width =1872
                    Height =218
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lbl_M2BarreDeb"
                    Caption =" "
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =3798
                    LayoutCachedTop =170
                    LayoutCachedWidth =5670
                    LayoutCachedHeight =388
                    BackThemeColorIndex =4
                    ForeThemeColorIndex =9
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    Left =4478
                    Top =566
                    Width =1584
                    Height =398
                    FontSize =10
                    FontWeight =700
                    BorderColor =8355711
                    Name ="lbl_M2BarreFin"
                    Caption =" "
                    GridlineColor =10921638
                    LayoutCachedLeft =4478
                    LayoutCachedTop =566
                    LayoutCachedWidth =6062
                    LayoutCachedHeight =964
                    ThemeFontIndex =-1
                    BackThemeColorIndex =5
                    BackShade =75.0
                    ForeThemeColorIndex =9
                    ForeTint =100.0
                    ForeShade =75.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3061
                    Top =2097
                    Width =794
                    Height =340
                    ForeColor =4210752
                    Name ="Commande0"
                    Caption ="Lancer"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =3061
                    LayoutCachedTop =2097
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =2437
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =3
                    WebImagePaddingTop =3
                    WebImagePaddingRight =2
                    WebImagePaddingBottom =2
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3120
                    Top =1700
                    Width =624
                    Height =300
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtBoucle"
                    DefaultValue ="50"
                    GridlineColor =10921638

                    LayoutCachedLeft =3120
                    LayoutCachedTop =1700
                    LayoutCachedWidth =3744
                    LayoutCachedHeight =2000
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2211
                            Top =1700
                            Width =804
                            Height =300
                            BorderColor =8355711
                            ForeColor =6710886
                            Name ="Étiquette4"
                            Caption ="Boucle :"
                            GridlineColor =10921638
                            LayoutCachedLeft =2211
                            LayoutCachedTop =1700
                            LayoutCachedWidth =3015
                            LayoutCachedHeight =2000
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =232
                    Top =396
                    Width =768
                    Height =288
                    FontSize =10
                    BorderColor =8355711
                    Name ="lbl_M2Texte"
                    Caption ="100 %"
                    GridlineColor =10921638
                    HorizontalAnchor =1
                    LayoutCachedLeft =232
                    LayoutCachedTop =396
                    LayoutCachedWidth =1000
                    LayoutCachedHeight =684
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextAlign =2
                    TextFontFamily =2
                    Left =510
                    Top =1020
                    Width =420
                    Height =480
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =6710886
                    Name ="lbl_Boucle"
                    Caption ="·"
                    FontName ="Wingdings"
                    LeftPadding =0
                    TopPadding =0
                    RightPadding =0
                    BottomPadding =0
                    GridlineColor =10921638
                    LayoutCachedLeft =510
                    LayoutCachedTop =1020
                    LayoutCachedWidth =930
                    LayoutCachedHeight =1500
                    ThemeFontIndex =-1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="PiedFormulaire"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Demo")
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
'//             le lable1 (lbl_BarreDeb) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'DROITE'
'//             le lable1 (lbl_BarreFin) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'GAUCHE'
'//             Les 2 labels doivent avoir leurs point d'encrage verticale(VerticalAnchor) défini sur 'HAUT'
'//
'// Dans le cas contraire les labels ne se potionnerons pas correctement.
'//
Option Compare Database
Option Explicit

Private Sub Commande0_Click()

    Dim lbl1        As Access.Label
    Dim lbl2        As Access.Label
    Dim lblTxt      As Access.Label

    Const MAX_WIDTH As Integer = 7000       '// Taille des label barres.
    Const DECAL     As Integer = 75         '// Décalage entre lbl1 et lbl2.
    Const LBL_HAUT  As Integer = 100
    Const TXT_P     As String = " %"        '// Info texte.
    Const CHAR_TIME As String = 183         '// Charactère debut horloge.

    Dim Boucle      As Integer              '// Definir le nb total à parcourir, initialise les incréments.

    Dim IncrLbl1    As Double               '// Incrément décalage barre 1.
    Dim Pos         As Integer
    Dim Pourcent    As Double               '// Texte avancement.
    Dim Avance      As Double               '//     ""
    Dim TxtAvance   As String               '//     ""
    Dim CptH        As Integer              '// Compteur horloge...
    

    Boucle = Nz(txtBoucle, 0)  'TODO: val boucle demo.
    If (Boucle < 10) Then Exit Sub

    '// Initialisation
    Set lbl1 = lbl_M2BarreDeb
    Set lbl2 = lbl_M2BarreFin
    Set lblTxt = lbl_M2Texte

    '// Positionnement des labels par rapport au labelTexte 'lblTxt' (c'est lui qui défini la position des autres labels)
    lbl1.Width = 0
    lbl1.Height = LBL_HAUT
    lbl1.Top = lblTxt.Top + (lblTxt.Height / 4)
    lbl1.Left = lblTxt.Left + (lblTxt.Width + 50)   '// Décale légerement la barre du texte de 50.
    
    lbl2.Width = MAX_WIDTH - DECAL
    lbl2.Height = LBL_HAUT
    lbl2.Left = lbl1.Left
    lbl2.Top = lbl1.Top

    lbl_Boucle.Top = lblTxt.Top - (lblTxt.Height / 4)
    lbl_Boucle.Left = lbl_M2BarreFin.Left + lbl_M2BarreFin.Width + (DECAL * 2)  '// Positionne ala fin de la barre 2 + décalage.

    '// Calul incrément pour dimentionner les labels.
    IncrLbl1 = (MAX_WIDTH / Boucle) + 1                 '// Déclage dus aux valeurs non entières (ex: 5.26)
    Pourcent = 100 / Boucle                             '// Texte pourcentage avancement.

    lbl1.Visible = True
    lbl2.Visible = True
    lblTxt.Visible = True
    lbl_Boucle.Visible = True

    '// Boucle démo....
    For Pos = IncrLbl1 To MAX_WIDTH Step IncrLbl1

        lbl_Boucle.Caption = ChrW$(CptH + CHAR_TIME)    '// Pour le fun ...
        CptH = CptH + 1                                 '//     ""
        If CptH > 11 Then CptH = 0                      '//     ""

        Avance = Avance + Pourcent
        TxtAvance = Format$(Avance, "###")

        lbl1.Width = lbl1.Width + IncrLbl1
        lbl2.Left = (lbl1.Left + lbl1.Width) + DECAL
        lbl2.Width = MAX_WIDTH - (lbl1.Width)
        lblTxt.Caption = TxtAvance & TXT_P

        DoEvents

        Sleep (100)

    Next

    lbl2.Visible = False
    lbl_Boucle.Visible = False
    lblTxt.Caption = 100 & TXT_P
    lbl1.Width = MAX_WIDTH
    lbl2.Visible = False
    DoEvents

    Set lbl1 = Nothing
    Set lbl2 = Nothing
    Set lblTxt = Nothing

End Sub
