# MsAccess-ProgressBarrre2

v 1.0 Petite démo de barre de progression.

Idée inspiré de [nala](https://github.com/volitank/nala)

# Formulaire F_Barre2Demo

![Formulaire de démarrage](Doc/F_Demo.gif)

# Utilisation :

- Ouvrir le formulaire `F_Barre2Demo`

# Déclaration de la classe

```VB
    Dim Prb As C_ProgB2
    Set Prb = New C_ProgB2
```
- INFO vous pouvez déclarer le nom des labels (AVANT l'initialisation) avec :
```VB
    Prb.NomLabelGauche = Me.lblxx.Name
    Prb.NomLabelDroite = Me.lblxx.Name
    Prb.NomLabelInfo = Me.lblxx.Name
    Prb.NomLabelRota = Me.lblxx.Name
```
- Sinon la classe utiliseras les noms par defaut avec les Constantes : `LBL_GAUCHE`, `LBL_DROITE`, `LBL_INFO`, `LBL_ROTA`.

- Paramètres optionnels de la barre (a faire AVANT l'initialisation (`InitLabels`))
```VB
    Prb.LabelWidth = Nz(Me.txtTaille)   '// Optionnel(Defaut voir Const LBL_WIDTH)  a definir AVANT l'initialisation.
    Prb.LabelHeight = Nz(Me.txtHauteur) '// Optionnel(Deafut voit Const LBL_HAUT)   a definir AVANT l'initialisation.
```
- Initialisation de la classe :
```VB
    Prb.InitLabels Me, Me.txtBoucle
```
- Actualisation :
```VB
    Prb.UpdateLabels
    DoEvents
```
- Reset des labels à la fin de votre code:
```VB
    Prb.CleanLabels
```

## Important :

- Pour un positionnement correct des labels, suivre les indications ci-dessous
    - L'emplacement du label de texte (`lbl_M2Texte`) définira la positions des autres labels (a partir de sont coin bas/droite).

    - le lable1 (`lbl_BarreDeb`) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'DROITE'
    - le lable1 (`lbl_BarreFin`) doit avoir le point d'encrage horizontal(HorizontalAnchor) défini sur 'GAUCHE'
    - Les 2 labels doivent avoir leurs point d'encrage vertical(VerticalAnchor) défini sur 'HAUT'
- Dans le cas contraire les labels ne se potionnerons pas correctement.

## Résumé

|   Créer le|   2022/09/04|
| - | - |
|   Auteur| [@meuslau](https://github.com/meuslaur)|
|   Catégorie|   MsAccess|
|   Type|   Utilitaire|
|   Langage|   VBA|

### Code exporté avec l'outil de : [@joyfullservice](https://github.com/joyfullservice) - [msaccess-vcs-integration](https://github.com/joyfullservice/msaccess-vcs-integration)

- Créez une base vide et utilisez `msaccess-vcs-integration` pour réimporter le code.
