VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Shortcuts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creating Shortcuts"
   ClientHeight    =   2895
   ClientLeft      =   1905
   ClientTop       =   1470
   ClientWidth     =   4830
   Icon            =   "Creating Shortcuts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2895
   ScaleWidth      =   4830
   Begin VB.Frame Frame 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Sourcepath 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox Destinationpath 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton Browsedestination 
         Caption         =   "Bro&wse..."
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton Browsesource 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Create 
         Caption         =   "Cr&eate shortcut..."
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Shortcutname 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   0
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select A File"
         Flags           =   6148
      End
      Begin VB.Label Sourcelabel 
         Caption         =   "&Source:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Destinationlabel 
         Caption         =   "&Destination:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Shortcutnamelabel 
         Caption         =   "&Name of Shortcut:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Shortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Private Sub Browsedestination_Click()

Dim bi As BROWSEINFO 'declare the needed variables
Dim rtn&, pidl&, path$, pos%

bi.hOwner = Me.hWnd 'centres the dialog on the screen
bi.lpszTitle = "Browse for Destination..." 'set the title text
bi.ulFlags = BIF_RETURNONLYFSDIRS 'the type of folder(s) to return
pidl& = SHBrowseForFolder(bi) 'show the dialog box
  
path = Space(512) 'sets the maximum characters
T = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path

pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
T = Left(path$, pos - 1)

A = Mid(T, InStr(T, ":") + 1) 'splits the drive "C:" and the path eg. "\My Documents" into two, and only returns the path
Destinationpath.Text = "...." + A 'puts the selected path into the destination path text box

'For explanations of why this has been done, read the Create Shortcut buttons code.

End Sub


Private Sub Browsesource_Click()

Dialog.ShowOpen 'show the dialog box
Sourcepath.Text = Dialog.filename 'set the selected file to the source text box

End Sub

Private Sub Cancel_Click()

Unload Me 'exit the program

End Sub









Private Sub Create_Click()

T = fCreateShellLink(Destinationpath.Text, Shortcutname.Text, Sourcepath.Text, "") 'create the shortcut with the data specified

'If you've studied the source code for Visual Basic 4's Setup Kit
'(SETUP1.VBP), you may know that STKIT432.DLL includes the following
'function for adding shortcuts to Windows 95 Start Menu:
'
'Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
'
'You may not know, however, that you can use the above function to create
'a shortcut anywhere on the drive containing the Start Menu.
'The functions first parameter (lpstrFolderName) is relative to the
'Start Menus 'Programs' folder, located on most systems at
'C:\Windows\Start Menu\Programs. This means that if you pass a null string
'as the first parameter, the shortcut will be created in the Programs
'folder itself. Similarly, you can navigate up from the Programs folder
'to any directory on the users hard disk. For example, the following will
'create a shortcut on the users desktop:
'
'lResult = fCreateShellLink("..\..\Desktop", "Shortcut Title", "d:\path\appname.exe", "")
'

End Sub


Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

'This project was downloaded from
'
'http://www.brianharper.demon.co.uk/
'
'Please use this project and all of its source code however you want.
'
'UNZIPPING
'To unzip the project files you will need a 32Bit unzipper program that
'can handle long file names. If you have a latest copy of Winzip installed
'on your system then you may use that. If you however dont have a copy,
'then visit my web site, go into the files section and from there you can
'click on the Winzip link to goto their site and download a copy of the
'program. By doing this you will now beable to unzip the project files
'retaining their proper long file names.
'Once upzipped, load up your copy of Visual Basic and goto
'File/Open Project. Locate the project files to where ever you unzipped
'them, then click Open. The project files will be loaded and are now ready
'for use.
'
'THE PROJECT
'I created this project for a menu system I was once programming. I needed
'to beable to create Windows 95 shortcuts using Visual Basic code. After
'surfing the Internet for quite some time, I came across a small article
'from Visual Basic Tips and Tricks which stated that in amongst your
'Visual basic setup kit, is a file called STKIT432.DLL. Contained in this
'file is a function that enables the user to create shortcuts on the
'Windows Start Menu. My problem was solved.
'
'Read the text in the Create shortcut buttons code for more information.
'
'To get a latest copy of Visual Basic Tips and Tricks point your web
'browser at:
'
'VB Tips & Tricks Home Page: http://www.apexsc.com/vb/davem/vbtt.html
'
'To have every tip or trick ever published in VB Tips & Tricks, you only
'need two files. For all tips up until December '95 you need the
'VBTTARC.EXE). For all '96 tips, you need VBTT.EXE. Those of you
'downloading VBTT for the very first time, download VBTTFULL.EXE.
'
'To see the VB Tip on creating Windows 95 shortcuts, search for:
'
'Creating Win95 Shell Links (Shortcuts).
'
'NOTES
'I have only provided the necessary project files with the zip. This keeps
'the size of the zip files down to a minimum and enables me to upload more
'prjects files to my site.
'
'I hope you find the project usful in what ever you are programming. I
'have tried to write out a small explanation of what each line of code
'does in the project, although most of it is pretty simple to understand.
'
'If you find any bugs in the code then please dont hesitate to Email me and
'I will get back to you as soon as possible. If you however need help on a
'different matter concerning Visual Basic then please please Email me as
'I like to here from people and here what they are programming.
'
'My Email address is:
'Brian@brianharper.demon.co.uk
'
'My web site is:
'http://www.brianharper.demon.co.uk/
'
'Please visit my web site and find many other useful projects like this.
'

End Sub


