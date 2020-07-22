VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmTerminator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terminator v1.0"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9825
   Icon            =   "frmTerminator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmTerminator.frx":393A1
   ScaleHeight     =   6525
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H0000FF00&
      Caption         =   "&Apagar: Terminator"
      BeginProperty Font 
         Name            =   "Terminator"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4320
      Top             =   3000
   End
   Begin VB.CommandButton cmdexterminar 
      BackColor       =   &H000000FF&
      Caption         =   "&Exterminar"
      BeginProperty Font 
         Name            =   "Terminator"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtprograma 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   7680
      TabIndex        =   0
      ToolTipText     =   "Nombre de Programa."
      Top             =   3270
      Width           =   375
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp2 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   10440
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   661
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp3 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   10440
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   661
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   10440
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   661
   End
   Begin VB.Image imgsonido 
      Height          =   1920
      Left            =   600
      Picture         =   "frmTerminator.frx":10B943
      Top             =   4560
      Width           =   1920
   End
   Begin VB.Image imgOff 
      Height          =   1920
      Left            =   5280
      Picture         =   "frmTerminator.frx":11C18D
      Top             =   4560
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgOn 
      Height          =   1920
      Left            =   3600
      Picture         =   "frmTerminator.frx":12C9D7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1920
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   10440
      Visible         =   0   'False
      Width           =   375
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   661
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   7
      X1              =   5400
      X2              =   7920
      Y1              =   2280
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   7
      X1              =   4560
      X2              =   7920
      Y1              =   2280
      Y2              =   3480
   End
End
Attribute VB_Name = "frmterminator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*
'* Programa Terminator
'* Autor: Martin Grasso Castrillo
'* Date 14 / 10 / 2017
'* destrulle el proceso del Listado de Memoria
'*
'*********************************************************
Function CierraProceso(StrNombreProceso As String, Optional DecirSINO _
As Boolean = True) As Boolean
On Error GoTo nose
  Dim ListaProcesos      As Object
  Dim ObjetoWMI          As Object
  Dim ProcesoConcreto    As Object
  CierraProceso = False
  Set ObjetoWMI = GetObject("winmgmts:")
  If IsNull(ObjetoWMI) = False Then
  Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
  For Each ProcesoConcreto In ListaProcesos
    If UCase(ProcesoConcreto.Name) = UCase(StrNombreProceso) Then
        If DecirSINO Then
          If MsgBox("¿Detener Programa ? " & _
               ProcesoConcreto.Name & vbNewLine & _
               "...¿Está seguro?", _
               vbYesNo + vbCritical) _
               = vbYes Then
           ProcesoConcreto.Terminate (0)
           CierraProceso = True
          End If
        Else
         ProcesoConcreto.Terminate (0)
         CierraProceso = True
        End If
     End If
    Next
  Else
  'pon aqui un msgbox con el error que se produzca
  End If
  Set ListaProcesos = Nothing
  Set ObjetoWMI = Nothing
nose:
End Function

'Visibiliza o no los Rayos Lasér en el control Lineal
Private Sub verRayos(ByVal r1 As Boolean _
, ByVal r2 As Boolean)
Line1.Visible = r1
Line2.Visible = r1
End Sub

'al precionar el bóton se ejecutan los eventos en el boton exterminar
Private Sub cmdexterminar_Click()
sonidoBoton
If Not (Me.txtprograma.Text = "") Then ' si existen datos en la cadena de texto.
    'crea palabras claves para Paquete Office
    Select Case (Me.txtprograma.Text)
           Case ("word")
           Me.txtprograma.Text = "Win" & "Word" 'destruye el word cargado en memoria de programas.
           Case ("exel")
           Me.txtprograma.Text = "Win" & "Excel"        '//
           Case ("powerpoint")
           Me.txtprograma.Text = "Win" & "PowerPoint"   '//
           Case ("paint")
           Me.txtprograma.Text = "ms" & "paint"         '//
     End Select
     CierraProceso Me.txtprograma.Text & ".exe", False '// ejecuta el destructor de objetos cerrando el proceso
     Me.txtprograma.Text = ""
 End If
 Timer1.Enabled = False
 verRayos False, False
End Sub

' al posisionar el mouse sobre el botón se eschara el audio del efecto
Private Sub cmdexterminar_MouseMove(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
sonidoOver
End Sub

' destruye el programa en memoria actual*
Private Sub cmdSalir_Click()
sonidoBoton
Unload Me
Unload frmAcercade
End
End Sub

' destruye el programa en memoria actual en el audio*
Private Sub cmdSalir_MouseMove(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
sonidoOver
End Sub

'cuando carga el programa
Private Sub Form_Load()
Me.Icon = frmAcercade.Icon ' el icono del programa hacer de es = al programa actual.
verRayos False, False      ' los rayos del lineado permanecen inactivos
Timer1.Enabled = False     ' el control de tiempo se desactiva.
With wmp
.URL = "fondo.mp3"         ' le pasa el audio a la URL en este caso es Terminetor
.settings.volume = 50      ' el volumen es establecido al 50% del audio para que no acople con el tema de fondo
.Controls.play             ' reproduce el audio en el control Windows media Palyer.
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdSalir_Click
End Sub

'al hacer click con el botón se eschucha o se termina un sonido
Private Sub imgsonido_Click()
If imgsonido.Tag = "" Then
   imgsonido.Picture = imgOff.Picture
   imgsonido.Tag = 1
   wmp.settings.mute = True
ElseIf imgsonido.Tag = 1 Then
   imgsonido.Picture = imgOn.Picture
   wmp.settings.mute = False
   imgsonido.Tag = ""
End If
End Sub

'genera el efecto de rayo lasér
Private Sub Timer1_Timer()
If Line1.Visible = False Then
   verRayos True, True
ElseIf Line1.Visible = True Then
   verRayos False, False
End If
End Sub

'sondo de bóton
Private Sub sonidoBoton()
With wmp1
.URL = "click.mp3"
.settings.volume = 100
.Controls.play
End With
End Sub

'sonido sobre bóton
Private Sub sonidoOver()
With wmp2
.URL = "over.mp3"
.settings.volume = 100
.Controls.play
End With
End Sub

'sonido en el texto
Private Sub sonidoTexto()
With wmp3
.URL = "teclado.mp3"
.settings.volume = 100
.Controls.play
End With
End Sub

'al oprimir sobre el botón se eschucha el audio
Private Sub txtprograma_Change()
sonidoTexto
txtprograma.ToolTipText = txtprograma.Text
Timer1.Enabled = True
End Sub

'sonido sobre el texto del programa
Private Sub txtprograma_MouseMove(Button As Integer, Shift _
As Integer, X As Single, Y As Single)
sonidoOver
End Sub
