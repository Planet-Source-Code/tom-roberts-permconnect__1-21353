VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{E20D3660-C041-11D3-BEE0-83B0ED0DE712}#1.0#0"; "SYSTRAYMS.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "                                 Stay Connected with PermConnect"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   12632256
      ForeColor       =   8388608
      TabCaption(0)   =   "Program"
      TabPicture(0)   =   "frmMain.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Options"
      TabPicture(1)   =   "frmMain.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Start"
      TabPicture(2)   =   "frmMain.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "About"
      TabPicture(3)   =   "frmMain.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   " PermConnect "
         BeginProperty Font 
            Name            =   "BrushScript BT"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3255
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   7335
         Begin VB.CommandButton cmdTray 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Tray"
            Height          =   375
            Left            =   3120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2160
            Width           =   1095
         End
         Begin VB.CommandButton cmdStart 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Start"
            Height          =   375
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2160
            Width           =   975
         End
         Begin VB.CommandButton cmdEnd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Close"
            Height          =   375
            Left            =   4440
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2160
            Width           =   975
         End
         Begin SysTrayCtl.cSysTray Tray 
            Left            =   120
            Top             =   1320
            _ExtentX        =   900
            _ExtentY        =   900
            InTray          =   0   'False
            TrayIcon        =   "frmMain.frx":037A
            TrayTip         =   "Right-Click"
         End
         Begin VB.PictureBox Picture1 
            Height          =   495
            Left            =   960
            Picture         =   "frmMain.frx":0694
            ScaleHeight     =   435
            ScaleWidth      =   435
            TabIndex        =   21
            Top             =   480
            Width           =   495
         End
         Begin SHDocVwCtl.WebBrowser Web 
            Height          =   375
            Left            =   960
            TabIndex        =   22
            Top             =   480
            Width           =   255
            ExtentX         =   450
            ExtentY         =   661
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin VB.Label lblCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   240
            TabIndex        =   23
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label lblCnt 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   1920
            TabIndex        =   19
            Top             =   2760
            Width           =   3495
         End
         Begin VB.Image Image1 
            Height          =   1095
            Left            =   2040
            Picture         =   "frmMain.frx":099E
            Stretch         =   -1  'True
            Top             =   720
            Width           =   3165
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PermConnect "
         BeginProperty Font 
            Name            =   "BrushScript BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   7335
         Begin VB.TextBox Text2 
            BackColor       =   &H00C00000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   3000
            TabIndex        =   28
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   5400
            Picture         =   "frmMain.frx":19A9
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdPrev 
            Height          =   375
            Left            =   1560
            Picture         =   "frmMain.frx":1C96
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Cancel "
            Height          =   375
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   2880
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Add"
            Height          =   375
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   2880
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sa&ve Settings"
            Height          =   375
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C00000&
            DataField       =   "URL"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   285
            Left            =   1920
            TabIndex        =   12
            Top             =   2160
            Width           =   3495
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Delete"
            Height          =   375
            Left            =   5040
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2880
            Width           =   1215
         End
         Begin MSDBCtls.DBList DBList1 
            Bindings        =   "frmMain.frx":1F73
            Height          =   1035
            Left            =   1920
            TabIndex        =   15
            Top             =   840
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   1826
            _Version        =   393216
            BackColor       =   12582912
            ForeColor       =   12632256
            ListField       =   "URL"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Data Data1 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "C:\AA PermConnect\PermCon.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   1560
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "Connect"
            Top             =   3000
            Width           =   4575
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Next Connection"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   1920
            TabIndex        =   27
            Top             =   1920
            Width           =   3495
         End
         Begin VB.Image Image4 
            Height          =   375
            Index           =   1
            Left            =   6120
            Picture         =   "frmMain.frx":1F87
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1095
         End
         Begin VB.Image Image3 
            Height          =   495
            Left            =   2760
            Picture         =   "frmMain.frx":2291
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1725
         End
         Begin VB.Image Image4 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.frx":329C
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Stay Connected!"
         BeginProperty Font 
            Name            =   "BrushScript BT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3255
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   7335
         Begin VB.Image Image2 
            Height          =   1095
            Left            =   2160
            Picture         =   "frmMain.frx":35A6
            Top             =   960
            Width           =   2685
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   " PermConnect"
         BeginProperty Font 
            Name            =   "BrushScript BT"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3255
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   7335
         Begin VB.CommandButton cmdSaveSet 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Save Settings"
            Height          =   375
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H00E0E0E0&
            Caption         =   "&Change Settings"
            Height          =   375
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CheckBox chk5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Five Minute Interval"
            Height          =   375
            Left            =   2640
            TabIndex        =   1
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CheckBox chk4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Four Minute Interval"
            Height          =   375
            Left            =   2640
            TabIndex        =   2
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CheckBox chk3 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Three Minute Interval"
            Height          =   375
            Left            =   2640
            TabIndex        =   3
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox chk2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Two Minute Interval"
            Height          =   375
            Left            =   2640
            TabIndex        =   4
            Top             =   840
            Width           =   1815
         End
         Begin VB.CheckBox chk1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "One Minure Interval"
            Height          =   375
            Left            =   2640
            TabIndex        =   5
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin VB.Menu popup 
      Caption         =   "Pop UP"
      Visible         =   0   'False
      Begin VB.Menu Div1 
         Caption         =   "*******************"
      End
      Begin VB.Menu Open 
         Caption         =   "   Open PermConnect"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "   Close PermConnect"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Div2 
         Caption         =   "*******************"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
  
  'cSysTray ocx came from Microsoft "VB Downloads"
  'The cSysTray OCX needs to be registered
  'If you are not sure how, there are many sample projects
  'at PSC. You can also go to the Common Controls Replacement
  'Project at "http://www.mvps.org/ccrp/". They have an ocx that
  'will put regestering you ocx's on the menu when you right click
  'an ocx.
          
           
  'Most of the code or code ideas have come from PSC as this is where
  'I've learned VB.  The Timer Module should I think should be credited
  'to  Patrick K. Bigley.  The rest of the code is modified from other
  'sources. i.e. "Begining VB6 Database Programming", by John Cornell.
  'Mastering VB6 by Evangelos Petroutsos. Yada Yada Yada.  What it boils
  'down to is learning is a slow process and How I got to Z from A is
  'unknown to me.
  
  'This is may not be an example of good code but it is a utility I built
  'to use.  It works and it will keep your connection. By the way, I call
  'it PermConnect Verson 1.5. Even though its Version 1.0.0.
  'Version 1 just didn't look or sound right.
  
  'My thanks to PCS and I hope someone will benifit either
  'by using the utility or picking out some snippits to use.
'-------------------------------------------------------------------
   'Keep Form on top - not sure where I got this code but it
   'sure beats an ocx. Look at Form Load
   Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
       Const HWND_TOPMOST = -1
       Const HWND_NOTTOPMOST = -2
       Const SWP_NOSIZE = 1
       Const SWP_NOMOVE = 2
       Private Flags As Integer
       Private lResult As Long
 '///////////////////////////////////////////////////////
 'This is setting up the .INI. Not sure where I got the code but
 'I've used it in a number of projects.  I've tried the .dll's and
 'other ini code but I always come back to this. Also, ini6 is not
 'being used.
    
    Dim KeySection1 As String
     Dim KeySection2 As String
      Dim KeySection3 As String
       Dim KeySection4 As String
        Dim KeySection5 As String
         Dim KeySection6 As String
         
    Dim KeyKey1 As String
     Dim KeyKey2 As String
      Dim KeyKey3 As String
       Dim KeyKey4 As String
        Dim KeyKey5 As String
         Dim KeyKey6 As String
         
    Dim KeyValue1 As String
     Dim KeyValue2 As String
      Dim KeyValue3 As String
       Dim KeyValue4 As String
        Dim KeyValue5 As String
         Dim KeyValue6 As String
         Dim TotalRecords As Long
Private Sub loadini()
On Error Resume Next
Dim lngResult1 As Long
 Dim lngResult2 As Long
  Dim lngResult3 As Long
   Dim lngResult4 As Long
    Dim lngResult5 As Long
     Dim lngResult6 As Long

Dim strFileName1
 Dim strFileName2
  Dim strFileName3
   Dim strFileName4
    Dim strFileName5
     Dim strFileName6
     
 Dim strResult1 As String * 50
  Dim strResult2 As String * 50
   Dim strResult3 As String * 50
    Dim strResult4 As String * 50
     Dim strResult5 As String * 50
      Dim strResult6 As String * 50
     
strFileName1 = App.Path & "\Myini1.ini" 'Declare your ini file !

lngResult1 = GetPrivateProfileString(KeySection1, _
 KeyKey1, strFileName1, strResult1, Len(strResult1), _
  strFileName1)

If lngResult1 = 0 Then
'An error has occurred
 Call MsgBox("An error has occurred while calling the API function", vbExclamation)
 Else
     KeyValue1 = Trim(strResult1)


'=================================================================
strFileName2 = App.Path & "\Myini2.ini" 'Declare your ini file !

lngResult2 = GetPrivateProfileString(KeySection2, _
 KeyKey2, strFileName2, strResult2, Len(strResult2), _
  strFileName2)

If lngResult2 = 0 Then

 Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
           KeyValue2 = Trim(strResult2)

'=======================================================================

strFileName3 = App.Path & "\Myini3.ini" 'Declare your ini file !
 lngResult3 = GetPrivateProfileString(KeySection3, _
  KeyKey3, strFileName3, strResult3, Len(strResult3), _
   strFileName3)

If lngResult3 = 0 Then

  Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
          KeyValue3 = Trim(strResult3)

'=================================================================

strFileName4 = App.Path & "\Myini4.ini" 'Declare your ini file !
 lngResult4 = GetPrivateProfileString(KeySection4, _
  KeyKey4, strFileName4, strResult4, Len(strResult4), _
   strFileName4)

If lngResult4 = 0 Then

  Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
            KeyValue4 = Trim(strResult4)

'==============================

strFileName5 = App.Path & "\Myini5.ini" 'Declare your ini file
 lngResult5 = GetPrivateProfileString(KeySection5, _
  KeyKey5, strFileName5, strResult5, Len(strResult5), _
   strFileName5)

If lngResult5 = 0 Then

  Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
           KeyValue5 = Trim(strResult5)

'=================================================

strFileName6 = App.Path & "\Myini6.ini" 'Declare your ini file !

lngResult6 = GetPrivateProfileString(KeySection6, _
 KeyKey6, strFileName6, strResult6, Len(strResult6), _
  strFileName6)

If lngResult6 = 0 Then
'An error has occurred
  Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
             KeyValue6 = Trim(strResult6)

      End If
     End If
    End If
   End If
  End If
 End If
 
End Sub
Private Sub saveini1()
On Error Resume Next
Dim lngResult1 As Long
Dim strFileName1

strFileName1 = App.Path & "\Myini1.ini" 'Declare your ini file !
 lngResult1 = WritePrivateProfileString(KeySection1, _
  KeyKey1, KeyValue1, strFileName1)

If lngResult1 = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If
End Sub
Private Sub saveini2()
On Error Resume Next
Dim lngResult2 As Long
Dim strFileName2

strFileName2 = App.Path & "\Myini2.ini" 'Declare your ini file !
 lngResult2 = WritePrivateProfileString(KeySection2, _
  KeyKey2, KeyValue2, strFileName2)

If lngResult2 = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If
End Sub
Private Sub saveini3()
On Error Resume Next
Dim lngResult3 As Long
Dim strFileName3

strFileName3 = App.Path & "\Myini3.ini" 'Declare your ini file !
 lngResult3 = WritePrivateProfileString(KeySection3, _
  KeyKey3, KeyValue3, strFileName3)

If lngResult3 = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If
End Sub
Private Sub saveini4()
On Error Resume Next
Dim lngResult4 As Long
Dim strFileName4

strFileName4 = App.Path & "\Myini4.ini" 'Declare your ini file !
 lngResult4 = WritePrivateProfileString(KeySection4, _
  KeyKey4, KeyValue4, strFileName4)

If lngResult4 = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If
End Sub
Private Sub saveini5()
On Error Resume Next
Dim lngResult5 As Long
Dim strFileName5

strFileName5 = App.Path & "\Myini5.ini" 'Declare your ini file !
 lngResult5 = WritePrivateProfileString(KeySection5, _
  KeyKey5, KeyValue5, strFileName5)

If lngResult5 = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If
End Sub
Private Sub saveini6()
On Error Resume Next
Dim lngResult6 As Long
Dim strFileName6

strFileName6 = App.Path & "\Myini6.ini" 'Declare your ini file !
 lngResult6 = WritePrivateProfileString(KeySection6, _
  KeyKey6, KeyValue6, strFileName6)

If lngResult6 = 0 Then
'An error has occurred

Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If
End Sub


Private Sub chk1_Click()
'This (and the following five) prevents the user from
'checking more that one checkbox.
If chk1.Value = 1 Then
    chk2.Value = 0
      chk3.Value = 0
        chk4.Value = 0
          chk5.Value = 0
End If
End Sub

Private Sub chk2_Click()
If chk2.Value = 1 Then
    chk1.Value = 0
     chk3.Value = 0
      chk4.Value = 0
       chk5.Value = 0
End If
End Sub

Private Sub chk3_Click()
If chk3.Value = 1 Then
    chk2.Value = 0
     chk1.Value = 0
      chk4.Value = 0
       chk5.Value = 0
End If
End Sub

Private Sub chk4_Click()
If chk4.Value = 1 Then
    chk2.Value = 0
     chk3.Value = 0
      chk1.Value = 0
       chk5.Value = 0
End If
End Sub

Private Sub chk5_Click()
If chk5.Value = 1 Then
    chk2.Value = 0
     chk3.Value = 0
      chk4.Value = 0
       chk1.Value = 0
End If
End Sub

Private Sub Close_Click()
End
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
Data1.Recordset.AddNew
cmdSave.Enabled = True
Call MsgBox("Don't forget to click on ""Save Settings""  ", vbInformation + vbOKOnly + vbDefaultButton1, "Save Your Changes")
End Sub

Private Sub cmdCancel_Click()
With Data1.Recordset
If .EditMode Then
   .CancelUpdate
   .MovePrevious
End If
End With
End Sub

Private Sub cmdDelete_Click()
cmdSave.Visible = True
With Data1.Recordset
          .Delete
          .MoveNext
  DBList1.ReFill
Call MsgBox("Don't forget to click on ""Save Settings""  ", vbInformation + vbOKOnly + vbDefaultButton1, "Save Your Changes")
End With
End Sub

Private Sub cmdEdit_Click()
cmdSaveSet.Enabled = True

chk1.Enabled = True
 chk2.Enabled = True
  chk3.Enabled = True
   chk4.Enabled = True
    chk5.Enabled = True
    
Call MsgBox("Don't forget to click on ""Save Settings""  ", vbInformation + vbOKOnly + vbDefaultButton1, "Save Your Changes")
End Sub

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdNext_Click()
'Provides For continuous scroll
With Data1.Recordset
          .MoveNext
       If .EOF Then
          .MoveFirst
End If
End With
End Sub

Private Sub cmdPrev_Click()
'Provides For continuous scroll
With Data1.Recordset
          .MovePrevious
       If .BOF Then
          .MoveLast
End If
End With
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
With Data1.Recordset
          .Update
          .Bookmark = Data1.Recordset.LastModified
DBList1.ReFill
cmdSave.Enabled = False
End With
End Sub

Private Sub cmdSaveSet_Click()
'Save CheckBox Values to Ini so when the user starts up,
'the previous settings are initiated via saved Ini File
If chk1.Value = 1 Or chk1.Value = 0 Then
   KeySection1 = "Check Value"
    KeyKey1 = "Input"
     KeyValue1 = chk1.Value
saveini1 '(MyIni1.ini)

If chk2.Value = 1 Or chk2.Value = 0 Then
   KeySection2 = "Check Value"
    KeyKey2 = "Input"
     KeyValue2 = chk2.Value
saveini2 '(MyIni1.ini)

If chk3.Value = 1 Or chk3.Value = 0 Then
   KeySection3 = "Check Value"
    KeyKey3 = "Input"
     KeyValue3 = chk3.Value
saveini3 '(MyIni1.ini)

If chk4.Value = 1 Or chk4.Value = 0 Then
   KeySection4 = "Check Value"
    KeyKey4 = "Input"
     KeyValue4 = chk4.Value
saveini4 '(MyIni1.ini)

If chk5.Value = 1 Or chk5.Value = 0 Then
   KeySection5 = "Check Value"
    KeyKey5 = "Input"
     KeyValue5 = chk5.Value
saveini5 '(MyIni1.ini)

End If
 End If
  End If
   End If
    End If

chk1.Enabled = False
 chk2.Enabled = False
  chk3.Enabled = False
   chk4.Enabled = False
    chk5.Enabled = False

cmdSaveSet.Enabled = False

End Sub

Private Sub cmdStart_Click()

Dim chkTime As Integer
   If chk1.Value = 1 Then
      chkTime = 60
Else
   If chk2.Value = 1 Then
      chkTime = 120
Else
   If chk3.Value = 1 Then
      chkTime = 180
Else
   If chk4.Value = 1 Then
      chkTime = 240
Else
   If chk5.Value = 1 Then
      chkTime = 300
End If
 End If
  End If
   End If
    End If
'.........................................................
'The following was the main Loop for a nested loop.  I used it
'because I had problems with the loop (used now)getting stalled
'due to .EOF and I couldn't get a Totalizer.  I left the code in
'to show a nested loop.
'Dim Count As Integer
'Dim A
'Count = A
'For A = 1 To 5000
'..........................................................
Dim Counter As Integer
Dim I 'As Integer
Dim L As Integer
L = 2880 'There has to be an exit or the stack will be a problem
         '2880 is 48 hours of running before you get an error.

For I = 1 To L
Counter = I
lblCnt.Caption = ("Successful Connections:  " & Counter & " Time(s)")

 
   Web.Navigate2 Text1.Text
    Data1.Recordset.MoveNext
     Sleep chkTime
      Text1.Refresh
       
  'When we get to the last record, What's
  'The loop to do
  If Data1.Recordset.EOF Then
   Data1.Recordset.MovePrevious
    Sleep 1
     Data1.Recordset.MoveFirst
      Web.Navigate2 Text1.Text
      Sleep chkTime
       Text1.Refresh

End If


Next
'The is the Outer Loop That I removed.
'................................
'lblCounter.Caption = "Connected:     " & Count + 1 & "   Times"   '" & (Count + 1) * (TotalRecords) & "   Times"
'If Data1.Recordset.EOF = True Then
'Data1.Recordset.MovePrevious
'Sleep 1
'Data1.Recordset.MoveFirst
'Text1.Refresh

'End If

'Next
End Sub


Private Sub cmdTray_Click()
Tray.InTray = True
frmMain.Hide
End Sub

Private Sub Data1_Reposition()
'Show Number of URLs (Records in DB)
Text2.Text = Data1.Recordset.AbsolutePosition + 1 & _
"  of  " & TotalRecords & " URLs"
End Sub

Private Sub Form_Activate()
With Data1.Recordset
.MoveLast
TotalRecords = .RecordCount
.MoveFirst
End With
End Sub

Private Sub Form_Load()
cmdSave.Enabled = False
cmdSaveSet.Enabled = False

Data1.Visible = False
Dim My1 As String
Dim My2 As String
Dim My3 As String
Dim My4 As String
Dim My5 As String
'Keeps Form On Top.
Flags = SWP_NOSIZE Or SWP_NOMOVE
lResult = SetWindowPos(frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)

 'This will place the form in the lower right portion of the screen
 Move (Screen.Width - Width) / 1.5, (Screen.Height - Height) * 0.8




Data1.DatabaseName = App.Path & "\PermCon.mdb"


chk1.Enabled = False
 chk2.Enabled = False
  chk3.Enabled = False
   chk4.Enabled = False
    chk5.Enabled = False
    
'////////////////////
'Get Values you created for the CheckBoxes
KeySection1 = "Check Value"
KeyKey1 = "Input"
loadini '(Myini1.ini)
My1 = KeyValue1
chk1.Value = My1
'//////////////////////
'Get Values you created for the CheckBoxes
KeySection2 = "Check Value"
KeyKey2 = "Input"
loadini
My2 = KeyValue2
chk2.Value = My2
'/////////////////////////
'Get Values you created for the CheckBoxes
KeySection3 = "Check Value"
KeyKey3 = "Input"
loadini
My3 = KeyValue3
chk3.Value = My3
'/////////////////////////
'Get Values you created for the CheckBoxes
KeySection4 = "Check Value"
KeyKey4 = "Input"
loadini
My4 = KeyValue4
chk4.Value = My4
'/////////////////////////
'Get Values you created for the CheckBoxes
KeySection5 = "Check Value"
KeyKey5 = "Input"
loadini
My5 = KeyValue5
chk5.Value = My5

End Sub

Private Sub Open_Click()
frmMain.Show
End Sub

Private Sub Tray_MouseDown(Button As Integer, Id As Long)

If Button = 2 Then
PopupMenu popup
End If

End Sub

