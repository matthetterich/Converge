VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Dim objAspEmail As ASPEMAILLib.MailSender
    Set objAspEmail = New ASPEMAILLib.MailSender
'    With objAspEmail
'        Call .AddAddress("MattHetterich@cinci.rr.com", "Matt Hetterich")
'        .From = "MattHetterich@Cinci.rr.com"
''        .Subject = "asp email test MYBOX"
'        .Subject = "asp email test FROM TSERVER"
'        .IsHTML = True
'        .Body = "<b>this is a test</b>"
'        .Username = "MattHetterich@Cinci.rr.com"
'        .Password = "mh021000"
'        .Host = "smtp-server.cinci.rr.com"
'        .Port = 25
'        Call .Send
'    End With

    Set objAspEmail = New ASPEMAILLib.MailSender
    With objAspEmail
        Call .AddAddress("MattHetterich@cinci.rr.com", "Matt Hetterich")
        .From = "MHetterich@TimcoRubber.com"
'        .Subject = "asp email test MYBOX"
        .Subject = "asp email test FROM TSERVER"
        .IsHTML = True
        .Body = "<b>this is a test</b>"
        .Username = "MHetterich@TimcoRubber.com"
        .Password = "timco1"
        .Host = "smtp.registeredsite.com"
        .Port = 25
'        .Username = "MattHetterich@Cinci.rr.com"
'        .Password = "mh021000"
'        .Host = "smtp-server.cinci.rr.com"
'        .Port = 25
         Call .Send
    End With
    
    Set objAspEmail = New ASPEMAILLib.MailSender
    With objAspEmail
        Call .AddAddress("mhetterich@timcorubber.com", "Matt Hetterich")
        .From = "MHetterich@TimcoRubber.com"
'        .Subject = "asp email test MYBOX"
        .Subject = "asp email test FROM TSERVER"
        .IsHTML = True
        .Body = "<b>this is a test</b>"
        .Username = "MHetterich@TimcoRubber.com"
        .Password = "timco1"
        .Host = "smtp.registeredsite.com"
        .Port = 25
'        .Username = "MattHetterich@Cinci.rr.com"
'        .Password = "mh021000"
'        .Host = "smtp-server.cinci.rr.com"
'        .Port = 25
         Call .Send
    End With
End Sub
