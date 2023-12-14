VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aplikasi rincian informasi dan spefikasi game PC "
   ClientHeight    =   6555
   ClientLeft      =   6285
   ClientTop       =   2340
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10200
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3120
      Top             =   6600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.Frame Frame5 
      Caption         =   "Pencarian"
      Height          =   6495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   2055
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5655
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   -1  'True
         ColumnHeaders   =   0   'False
         HeadLines       =   0
         RowHeight       =   15
         RowDividerStyle =   6
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   2
            ScrollBars      =   0
            Locked          =   -1  'True
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox carigame 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   4575
      Left            =   6960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4515
      ScaleWidth      =   3075
      TabIndex        =   18
      Top             =   240
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rating"
      Height          =   1695
      Left            =   6960
      TabIndex        =   3
      Top             =   4800
      Width           =   3135
      Begin VB.CommandButton ESRB 
         Caption         =   "Klik untuk mengetahui tentang ESRB"
         Height          =   1095
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         FillColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   2040
         Picture         =   "Form1.frx":0E63
         ScaleHeight     =   1275
         ScaleWidth      =   915
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info Game"
      Height          =   6375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox Text1 
         DataField       =   "sinopsis"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "memo"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   4440
         Width           =   4455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Spefikasi"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   4455
         Begin VB.Label Label22 
            Caption         =   "Label22"
            DataField       =   "hdd"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   1200
            TabIndex        =   31
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label Label21 
            Caption         =   "Label21"
            DataField       =   "gpu"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   1200
            TabIndex        =   30
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label20 
            Caption         =   "Label20"
            DataField       =   "ram"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   1200
            TabIndex        =   29
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label19 
            Caption         =   "Label19"
            DataField       =   "cpu"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   1200
            TabIndex        =   28
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label18 
            Caption         =   "Label18"
            DataField       =   "os"
            DataSource      =   "Adodc1"
            Height          =   255
            Left            =   1200
            TabIndex        =   27
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label11 
            Caption         =   "HDD"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "VGA"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "RAM"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "CPU"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "OS"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         DataField       =   "bahasa"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1320
         TabIndex        =   26
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "Label16"
         DataField       =   "genre"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         DataField       =   "rilis"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         DataField       =   "publisher"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label23 
         Caption         =   "Sinopsis"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         DataField       =   "developer"
         DataSource      =   "Adodc1"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         DataField       =   "nama"
         DataSource      =   "Adodc1"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label6 
         Caption         =   "Bahasa"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Genre"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Rilis"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Publisher"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Developer"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Nama"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label rating 
      Caption         =   "rating"
      DataField       =   "rating"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   8400
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label cover 
      Caption         =   "cover"
      DataField       =   "cover"
      DataSource      =   "Adodc1"
      Height          =   255
      Left            =   7320
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      DataField       =   "nama"
      DataSource      =   "Adodc1"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu TESRB 
         Caption         =   "Tentang ESRB"
      End
      Begin VB.Menu keluar 
         Caption         =   "Keluar"
      End
   End
   Begin VB.Menu Tentang 
      Caption         =   "Tentang"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Koneksi As New ADODB.Connection
Dim RSBarang As ADODB.Recordset
Sub BukaDB()
Set Koneksi = New ADODB.Connection
Set RSBarang = New ADODB.Recordset
Koneksi.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\data.mdb"
End Sub
Private Sub DataGrid1_Click()
Picture1.Picture = LoadPicture(App.Path & "\data\gambar\cover\" & cover & ".jpg")
Picture2.Picture = LoadPicture(App.Path & "\data\gambar\rating\" & rating & ".jpg")
End Sub
Private Sub ESRB_Click()
Form2.Show
End Sub
Private Sub Form_Load()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\data.mdb; "
Adodc1.RecordSource = "rincian_game"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub
Private Sub carigame_Change()
Call BukaDB
        RSBarang.Open "select * from rincian_game where nama like '%" & carigame & "%'", Koneksi
        If Not RSBarang.EOF Then
                Adodc1.RecordSource = "select * from rincian_game where nama like '%" & carigame & "%'"
                Adodc1.Refresh
                Set DataGrid1.DataSource = Adodc1
        End If
End Sub
Private Sub keluar_Click()
Unload Me
End Sub
Private Sub Picture1_Click()
Picture1.Picture = LoadPicture(cover)
End Sub
Private Sub TESRB_Click()
Form2.Show
End Sub
