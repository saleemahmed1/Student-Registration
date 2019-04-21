VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   Caption         =   "Student Registration System"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   ForeColor       =   &H8000000B&
   Icon            =   "Student Registration System.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Student Registration System.frx":17D04
   ScaleHeight     =   10350
   ScaleWidth      =   20235
   Begin VB.ComboBox Combo4 
      DataField       =   "Gender"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   46
      Text            =   "Select Gender"
      Top             =   8040
      Width           =   4815
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Religion"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11760
      TabIndex        =   45
      Text            =   "Select Religion"
      Top             =   5880
      Width           =   4815
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Last 
      BackColor       =   &H00C0C000&
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton First 
      BackColor       =   &H00C0C000&
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Next 
      BackColor       =   &H00C0C000&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Previous 
      BackColor       =   &H00C0C000&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Update 
      BackColor       =   &H00C0C000&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton ViewRecord 
      BackColor       =   &H00C0C000&
      Caption         =   "View  Record"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8640
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Class from which Left"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11760
      TabIndex        =   37
      Text            =   "Select Class"
      Top             =   3720
      Width           =   4815
   End
   Begin VB.CommandButton Delete 
      BackColor       =   &H00C0C000&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton SaveBtn 
      BackColor       =   &H00C0C000&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton Addnew 
      BackColor       =   &H00C0C000&
      Caption         =   " Add New"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton UploadImage 
      BackColor       =   &H00808000&
      Caption         =   "Browse Image"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17160
      TabIndex        =   33
      Top             =   5160
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Class in which Admit"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   30
      Text            =   "Select Class"
      Top             =   7320
      Width           =   4815
   End
   Begin VB.TextBox Text15 
      DataField       =   "Image"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   29
      Text            =   "Text3"
      Top             =   5160
      Width           =   4815
   End
   Begin VB.TextBox Text14 
      DataField       =   "Remarks"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   28
      Text            =   "Text3"
      Top             =   4440
      Width           =   4815
   End
   Begin VB.TextBox Text12 
      DataField       =   "Date of Leaving"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   27
      Text            =   "Text3"
      Top             =   3000
      Width           =   4815
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   17160
      ScaleHeight     =   2715
      ScaleWidth      =   2355
      TabIndex        =   21
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      DataField       =   "Progress"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   20
      Text            =   "Text10"
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox Text9 
      DataField       =   "Date of Admission"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Text            =   "Text9"
      Top             =   6600
      Width           =   4815
   End
   Begin VB.TextBox Text8 
      DataField       =   "Last School Attended"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   5880
      Width           =   4815
   End
   Begin VB.TextBox Text7 
      DataField       =   "Place of Birth"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   5160
      Width           =   4815
   End
   Begin VB.TextBox Text6 
      DataField       =   "Date of Birth"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Text            =   "Text6"
      Top             =   4440
      Width           =   4815
   End
   Begin VB.TextBox Text5 
      DataField       =   "Caste"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   3720
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      DataField       =   "Father's Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   3000
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      DataField       =   "Name of Student"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      DataField       =   "GR no"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   5280
      Top             =   8400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Student Registration System\data.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\Student Registration System\data.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17520
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   31
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   26
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   25
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Class form which left"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9480
      TabIndex        =   24
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Leaving"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9600
      TabIndex        =   23
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Class in which Admit"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   7320
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Admission"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Last School Attended"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Place of Birth"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Caste"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Father's Name"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Student"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GR #"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ID #"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Registration System"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Last_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Exit_Click()
Unload Form1
End Sub

Private Sub AddNew_Click()
Adodc1.Recordset.Addnew
clear
End Sub
Sub clear()
Text6.Text = "12/01/1947"
Combo1.Text = "Select Class"
Combo2.Text = "Select Class"
Combo3.Text = "Select Religion"
Combo4.Text = "Select Gender"
End Sub




Private Sub Update_Click()
Adodc1.Recordset.Update
MsgBox "Record Saved Successfully"
End Sub

Private Sub Delete_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Delete Record Confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been Deleted Successfully", vbInformation, "Message"
Else
MsgBox "Record Not Deleted...!!", vbInformation, "Message"

End If


End Sub

Private Sub ViewRecord_Click()
Form2.Show

End Sub

Private Sub Text15_Change()
Picture1.Picture = LoadPicture(Text15.Text)
End Sub

Private Sub SaveBtn_Click()
Adodc1.Recordset.Update
MsgBox "Record Saved Successfully"

End Sub

Private Sub Previous_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Next_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub First_Click()
Adodc1.Recordset.MoveFirst
End Sub

Private Sub Form_Load()

Dim str As String
Combo1.AddItem "One"
Combo1.AddItem "Two"
Combo1.AddItem "Three"
Combo1.AddItem "Four"
Combo1.AddItem "Five"
Combo1.AddItem "Sixth"
Combo1.AddItem "Seventh"
Combo1.AddItem "Eighth"
Combo1.AddItem "Nineth"
Combo1.AddItem "Matric"


Combo4.AddItem "Male"
Combo4.AddItem "Female"


Combo2.AddItem "One"
Combo2.AddItem "Two"
Combo2.AddItem "Three"
Combo2.AddItem "Four"
Combo2.AddItem "Five"
Combo2.AddItem "Sixth"
Combo2.AddItem "Seventh"
Combo2.AddItem "Eighth"
Combo2.AddItem "Nineth"
Combo2.AddItem "Matric"

Combo3.AddItem " Islam"
Combo3.AddItem " Hindu"
Combo3.AddItem " Chiristian"
Combo3.AddItem " Jews"



End Sub

Private Sub UploadImage_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*jpg"
Text15.Text = CommonDialog1.FileName

End Sub

