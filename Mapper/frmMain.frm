VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Maker"
   ClientHeight    =   8055
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSound 
      Caption         =   "Sounds"
      Height          =   1455
      Left            =   3960
      TabIndex        =   252
      Top             =   6480
      Width           =   3735
      Begin VB.ComboBox cmbSoundSpatialNode 
         Height          =   285
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   258
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkSoundSpatial 
         Caption         =   "Spatial"
         Height          =   255
         Left            =   960
         TabIndex        =   256
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdSoundNew 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   254
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmbSound 
         Height          =   285
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   253
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Node:"
         Height          =   255
         Left            =   2040
         TabIndex        =   257
         Top             =   650
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   255
         Top             =   290
         Width           =   735
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "Items"
      Height          =   2295
      Left            =   3960
      TabIndex        =   213
      Top             =   4080
      Width           =   3735
      Begin VB.TextBox txtItemValue 
         Height          =   255
         Left            =   960
         TabIndex        =   249
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox cmbItemNameSound 
         Height          =   285
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   245
         Top             =   1440
         Width           =   1815
      End
      Begin VB.ComboBox cmbItemActionSound 
         Height          =   285
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   244
         Top             =   1800
         Width           =   1815
      End
      Begin VB.ComboBox cmbItemName 
         Height          =   285
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   241
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdItemNew 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   240
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmbItemType 
         Height          =   285
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   239
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Value:"
         Height          =   255
         Left            =   240
         TabIndex        =   248
         Top             =   1100
         Width           =   615
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Name Sound:"
         Height          =   255
         Left            =   360
         TabIndex        =   247
         Top             =   1490
         Width           =   975
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Action Sound:"
         Height          =   255
         Left            =   360
         TabIndex        =   246
         Top             =   1850
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   243
         Top             =   290
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   242
         Top             =   740
         Width           =   735
      End
   End
   Begin VB.Frame fraNPC 
      Caption         =   "Characters"
      Height          =   3855
      Left            =   120
      TabIndex        =   212
      Top             =   4080
      Width           =   3735
      Begin VB.TextBox txtNPCRun 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   2640
         TabIndex        =   260
         Top             =   2520
         Width           =   375
      End
      Begin VB.ComboBox cmbNPCActionSound 
         Height          =   285
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   238
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cmbNPCNameSound 
         Height          =   285
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   237
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox txtNPCHPMax 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   1440
         TabIndex        =   234
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtNPCHPMin 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   840
         TabIndex        =   232
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtNPCDefMax 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   3240
         TabIndex        =   230
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtNPCDefMin 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   2640
         TabIndex        =   228
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtNPCStrMax 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   1440
         TabIndex        =   226
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtNPCStrMin 
         Alignment       =   2  'Center
         Height          =   225
         Left            =   840
         TabIndex        =   224
         Top             =   2160
         Width           =   375
      End
      Begin VB.ListBox lstNPCItem 
         Height          =   720
         Left            =   960
         TabIndex        =   219
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmbNPCType 
         Height          =   285
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   218
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdNPCNew 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   216
         Top             =   240
         Width           =   255
      End
      Begin VB.ComboBox cmbNPCName 
         Height          =   285
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   214
         Top             =   240
         Width           =   2295
      End
      Begin MSComctlLib.Slider sldNPCItem 
         Height          =   735
         Left            =   600
         TabIndex        =   220
         Top             =   1335
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   1296
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   1
         TickFrequency   =   25
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   165
         Left            =   3120
         TabIndex        =   262
         Top             =   2520
         Width           =   105
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Run:"
         Height          =   255
         Left            =   1920
         TabIndex        =   261
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Action Sound:"
         Height          =   255
         Left            =   240
         TabIndex        =   236
         Top             =   3290
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Name Sound:"
         Height          =   255
         Left            =   240
         TabIndex        =   235
         Top             =   2930
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   1200
         TabIndex        =   233
         Top             =   2550
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "HP:"
         Height          =   255
         Left            =   120
         TabIndex        =   231
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   3000
         TabIndex        =   229
         Top             =   2190
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Defense:"
         Height          =   255
         Left            =   1920
         TabIndex        =   227
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   1200
         TabIndex        =   225
         Top             =   2190
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Strength:"
         Height          =   255
         Left            =   120
         TabIndex        =   223
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label lblNPCItemPercent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 %"
         Height          =   165
         Left            =   285
         TabIndex        =   222
         Top             =   1560
         Width           =   210
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Items:"
         Height          =   255
         Left            =   240
         TabIndex        =   221
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   217
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   215
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame fraMap 
      Caption         =   "Map"
      Height          =   3855
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   204
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   205
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   202
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   203
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   200
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   201
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   198
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   199
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   196
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   197
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   194
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   195
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   192
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   193
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   2640
         TabIndex        =   190
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   191
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   188
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   189
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   10
         Left            =   3360
         TabIndex        =   186
         Top             =   240
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   187
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   184
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   185
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   182
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   183
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   13
         Left            =   840
         TabIndex        =   180
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   181
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   14
         Left            =   1200
         TabIndex        =   178
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   179
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   15
         Left            =   1560
         TabIndex        =   176
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   177
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   16
         Left            =   1920
         TabIndex        =   174
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   175
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   17
         Left            =   2280
         TabIndex        =   172
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   173
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   18
         Left            =   2640
         TabIndex        =   170
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   171
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   19
         Left            =   3000
         TabIndex        =   168
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   169
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   20
         Left            =   3360
         TabIndex        =   166
         Top             =   600
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   167
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   21
         Left            =   120
         TabIndex        =   164
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   20
            Left            =   0
            TabIndex        =   165
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   22
         Left            =   480
         TabIndex        =   162
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   21
            Left            =   0
            TabIndex        =   163
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   23
         Left            =   840
         TabIndex        =   160
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   22
            Left            =   0
            TabIndex        =   161
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   24
         Left            =   1200
         TabIndex        =   158
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   23
            Left            =   0
            TabIndex        =   159
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   25
         Left            =   1560
         TabIndex        =   156
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   24
            Left            =   0
            TabIndex        =   157
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   26
         Left            =   1920
         TabIndex        =   154
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   25
            Left            =   0
            TabIndex        =   155
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   27
         Left            =   2280
         TabIndex        =   152
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   26
            Left            =   0
            TabIndex        =   153
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   28
         Left            =   2640
         TabIndex        =   150
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   27
            Left            =   0
            TabIndex        =   151
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   29
         Left            =   3000
         TabIndex        =   148
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   28
            Left            =   0
            TabIndex        =   149
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   30
         Left            =   3360
         TabIndex        =   146
         Top             =   960
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   29
            Left            =   0
            TabIndex        =   147
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   31
         Left            =   120
         TabIndex        =   144
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   30
            Left            =   0
            TabIndex        =   145
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   32
         Left            =   480
         TabIndex        =   142
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   31
            Left            =   0
            TabIndex        =   143
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   33
         Left            =   840
         TabIndex        =   140
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   32
            Left            =   0
            TabIndex        =   141
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   34
         Left            =   1200
         TabIndex        =   138
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   33
            Left            =   0
            TabIndex        =   139
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   35
         Left            =   1560
         TabIndex        =   136
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   34
            Left            =   0
            TabIndex        =   137
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   36
         Left            =   1920
         TabIndex        =   134
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   35
            Left            =   0
            TabIndex        =   135
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   37
         Left            =   2280
         TabIndex        =   132
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   36
            Left            =   0
            TabIndex        =   133
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   38
         Left            =   2640
         TabIndex        =   130
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   37
            Left            =   0
            TabIndex        =   131
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   39
         Left            =   3000
         TabIndex        =   128
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   38
            Left            =   0
            TabIndex        =   129
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   40
         Left            =   3360
         TabIndex        =   126
         Top             =   1320
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   39
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   41
         Left            =   120
         TabIndex        =   124
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   40
            Left            =   0
            TabIndex        =   125
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   42
         Left            =   480
         TabIndex        =   122
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   41
            Left            =   0
            TabIndex        =   123
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   43
         Left            =   840
         TabIndex        =   120
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   42
            Left            =   0
            TabIndex        =   121
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   44
         Left            =   1200
         TabIndex        =   118
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   43
            Left            =   0
            TabIndex        =   119
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   45
         Left            =   1560
         TabIndex        =   116
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   44
            Left            =   0
            TabIndex        =   117
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   46
         Left            =   1920
         TabIndex        =   114
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   45
            Left            =   0
            TabIndex        =   115
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   47
         Left            =   2280
         TabIndex        =   112
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   46
            Left            =   0
            TabIndex        =   113
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   48
         Left            =   2640
         TabIndex        =   110
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   47
            Left            =   0
            TabIndex        =   111
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   49
         Left            =   3000
         TabIndex        =   108
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   48
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   50
         Left            =   3360
         TabIndex        =   106
         Top             =   1680
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   49
            Left            =   0
            TabIndex        =   107
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   51
         Left            =   120
         TabIndex        =   104
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   50
            Left            =   0
            TabIndex        =   105
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   52
         Left            =   480
         TabIndex        =   102
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   51
            Left            =   0
            TabIndex        =   103
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   53
         Left            =   840
         TabIndex        =   100
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   52
            Left            =   0
            TabIndex        =   101
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   54
         Left            =   1200
         TabIndex        =   98
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   53
            Left            =   0
            TabIndex        =   99
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   55
         Left            =   1560
         TabIndex        =   96
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   54
            Left            =   0
            TabIndex        =   97
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   56
         Left            =   1920
         TabIndex        =   94
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   55
            Left            =   0
            TabIndex        =   95
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   57
         Left            =   2280
         TabIndex        =   92
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   56
            Left            =   0
            TabIndex        =   93
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   58
         Left            =   2640
         TabIndex        =   90
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   57
            Left            =   0
            TabIndex        =   91
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   59
         Left            =   3000
         TabIndex        =   88
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   58
            Left            =   0
            TabIndex        =   89
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   60
         Left            =   3360
         TabIndex        =   86
         Top             =   2040
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   59
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   61
         Left            =   120
         TabIndex        =   84
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   60
            Left            =   0
            TabIndex        =   85
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   62
         Left            =   480
         TabIndex        =   82
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   61
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   63
         Left            =   840
         TabIndex        =   80
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   62
            Left            =   0
            TabIndex        =   81
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   64
         Left            =   1200
         TabIndex        =   78
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   63
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   65
         Left            =   1560
         TabIndex        =   76
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   64
            Left            =   0
            TabIndex        =   77
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   66
         Left            =   1920
         TabIndex        =   74
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   65
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   67
         Left            =   2280
         TabIndex        =   72
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   66
            Left            =   0
            TabIndex        =   73
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   68
         Left            =   2640
         TabIndex        =   70
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   67
            Left            =   0
            TabIndex        =   71
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   69
         Left            =   3000
         TabIndex        =   68
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   68
            Left            =   0
            TabIndex        =   69
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   70
         Left            =   3360
         TabIndex        =   66
         Top             =   2400
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   69
            Left            =   0
            TabIndex        =   67
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   71
         Left            =   120
         TabIndex        =   64
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   70
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   72
         Left            =   480
         TabIndex        =   62
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   71
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   73
         Left            =   840
         TabIndex        =   60
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   72
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   74
         Left            =   1200
         TabIndex        =   58
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   73
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   75
         Left            =   1560
         TabIndex        =   56
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   74
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   76
         Left            =   1920
         TabIndex        =   54
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   75
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   77
         Left            =   2280
         TabIndex        =   52
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   76
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   78
         Left            =   2640
         TabIndex        =   50
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   77
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   79
         Left            =   3000
         TabIndex        =   48
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   78
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   80
         Left            =   3360
         TabIndex        =   46
         Top             =   2760
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   79
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   81
         Left            =   120
         TabIndex        =   44
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   80
            Left            =   0
            TabIndex        =   45
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   82
         Left            =   480
         TabIndex        =   42
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   81
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   83
         Left            =   840
         TabIndex        =   40
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   82
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   84
         Left            =   1200
         TabIndex        =   38
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   83
            Left            =   0
            TabIndex        =   39
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   85
         Left            =   1560
         TabIndex        =   36
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   84
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   86
         Left            =   1920
         TabIndex        =   34
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   85
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   87
         Left            =   2280
         TabIndex        =   32
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   86
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   88
         Left            =   2640
         TabIndex        =   30
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   87
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   89
         Left            =   3000
         TabIndex        =   28
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   88
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   90
         Left            =   3360
         TabIndex        =   26
         Top             =   3120
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   89
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   91
         Left            =   120
         TabIndex        =   24
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   90
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   92
         Left            =   480
         TabIndex        =   22
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   91
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   93
         Left            =   840
         TabIndex        =   20
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   92
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   94
         Left            =   1200
         TabIndex        =   18
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   93
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   95
         Left            =   1560
         TabIndex        =   16
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   94
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   96
         Left            =   1920
         TabIndex        =   14
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   95
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   97
         Left            =   2280
         TabIndex        =   12
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   96
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   98
         Left            =   2640
         TabIndex        =   10
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   97
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   99
         Left            =   3000
         TabIndex        =   8
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   98
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   100
         Left            =   3360
         TabIndex        =   6
         Top             =   3480
         Width           =   255
         Begin VB.OptionButton opNode 
            Height          =   255
            Index           =   99
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Frame fraNode 
      Caption         =   "Node Properties"
      Height          =   3855
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdNodeEnd 
         Caption         =   "Make End Node"
         Height          =   375
         Left            =   960
         TabIndex        =   266
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNodeImage 
         Caption         =   "..."
         Height          =   255
         Left            =   3240
         TabIndex        =   265
         Top             =   2880
         Width           =   255
      End
      Begin VB.TextBox txtNodeImage 
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   264
         Top             =   2880
         Width           =   2055
      End
      Begin VB.ListBox lstNodeRequiredItem 
         Height          =   645
         Left            =   1080
         Style           =   1  'Checkbox
         TabIndex        =   250
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ListBox lstNodeNPC 
         Height          =   555
         Left            =   1080
         TabIndex        =   208
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ListBox lstNodeItem 
         Height          =   555
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.ListBox lstNodeMusic 
         Height          =   645
         Left            =   1080
         Style           =   1  'Checkbox
         TabIndex        =   259
         Top             =   240
         Width           =   2415
      End
      Begin MSComctlLib.Slider sldNodeItem 
         Height          =   480
         Left            =   720
         TabIndex        =   207
         Top             =   1095
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   847
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   1
         TickFrequency   =   25
      End
      Begin MSComctlLib.Slider sldNodeNPC 
         Height          =   480
         Left            =   720
         TabIndex        =   211
         Top             =   1690
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   847
         _Version        =   393216
         Orientation     =   1
         Max             =   100
         TickStyle       =   1
         TickFrequency   =   25
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Image:"
         Height          =   255
         Left            =   240
         TabIndex        =   263
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Required:"
         Height          =   255
         Left            =   120
         TabIndex        =   251
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblNodeNPCPercent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 %"
         Height          =   165
         Left            =   480
         TabIndex        =   210
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Characters:"
         Height          =   255
         Left            =   120
         TabIndex        =   209
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Items:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblNodeItemPercent 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 %"
         Height          =   165
         Left            =   480
         TabIndex        =   206
         Top             =   1200
         Width           =   210
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Music:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   8280
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   8880
      TabIndex        =   0
      Top             =   2520
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFilePlay 
         Caption         =   "Save and &Play"
      End
      Begin VB.Menu mnuFileBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' GUI variables
Const MAX_ROW As Integer = 10
Const MAX_COL As Integer = 10
Const MAX_NODE As Integer = MAX_ROW * MAX_COL
Dim blStart As Boolean      ' Start node needs to be selected
Dim current As Integer      ' Current node
Dim ignoreClick As Boolean  ' Ignore click for list boxes
' Node & Sound counters
Dim nodeCounter As Integer, soundCounter As Integer
' Map variables for VB
Dim mapStart As Integer, mapEnd As Integer
Dim mapPath As String, mapName As String, soundPath, imagePath
    ' Node correspondence
    Dim optionToNode(0 To 99) As Integer
    Dim nodeToOption() As Integer
    Dim nodeUsed(0 To 99) As Boolean
    ' Node properties
    Dim clcNodeMusic() As New Collection
    Dim clcNodeItem() As New Collection ' Percentages (0 if not in node)
    Dim clcNodeNPC() As New Collection ' Percentages (0 if not in node)
    Dim clcNodeRequiredItem() As New Collection
    Dim strImage() As String
    ' Characters
    Dim clcNPC As New Collection
    Dim clcNPCItem() As New Collection
    ' Items
    Dim clcItem As New Collection
    ' Sounds
    Dim clcSoundSpatial As New Collection
' Map file variables
    Dim nh() As NODEHEADER
    
Dim mh As MAPHEADER
' Load node properties into upper right panel
Private Sub loadNodeProperties(Index As Integer)
    ignoreClick = True
    Dim i As Integer, n As Integer
    n = optionToNode(Index)
    ' Load music
    For i = 0 To lstNodeMusic.ListCount - 1
        lstNodeMusic.Selected(i) = clcNodeMusic(n).Item(i + 1).mbBoolean
    Next
    ' Load required items
    For i = 0 To lstNodeRequiredItem.ListCount - 1
        lstNodeRequiredItem.Selected(i) = clcNodeRequiredItem(n).Item(i + 1).ibBoolean
    Next
    ' Load image string
    If strImage(n) <> "" Then
        txtNodeImage.Text = "...\" & strImage(n)
    Else
        txtNodeImage.Text = ""
    End If
    ignoreClick = False
    If lstNodeItem.ListCount > 0 Then
        lstNodeItem.ListIndex = 0
        lstNodeItem_Click
    End If
    If lstNodeNPC.ListCount > 0 Then
        lstNodeNPC.ListIndex = 0
        lstNodeNPC_Click
    End If
End Sub
' Load NPC properties into character panel
Private Sub loadCharacterProperties(Index As Integer)
    ignoreClick = True
    ' Stats, sounds, etc.
    cmbNPCType.ListIndex = clcNPC(Index + 1).cType
    txtNPCStrMin.Text = clcNPC(Index + 1).cStrMin
    txtNPCStrMax.Text = clcNPC(Index + 1).cStrMax
    txtNPCDefMin.Text = clcNPC(Index + 1).cDefMin
    txtNPCDefMax.Text = clcNPC(Index + 1).cDefMax
    txtNPCHPMin.Text = clcNPC(Index + 1).cHPMin
    txtNPCHPMax.Text = clcNPC(Index + 1).cHPMax
    txtNPCRun.Text = clcNPC(Index + 1).cRunPerc
    cmbNPCNameSound.ListIndex = clcNPC(Index + 1).cNameSound
    cmbNPCActionSound.ListIndex = clcNPC(Index + 1).cActionSound
    ignoreClick = False
    If lstNPCItem.ListCount > 0 Then
        lstNPCItem.ListIndex = 0
        lstNPCItem_Click
    End If
End Sub
' Load Item properties into item panel
Private Sub loadItemProperties(Index As Integer)
    ignoreClick = True
    ' Type, value, name, etc.
    cmbItemType.ListIndex = clcItem(Index + 1).iType
    txtItemValue.Text = clcItem(Index + 1).iValue
    cmbItemNameSound.ListIndex = clcItem(Index + 1).iNameSound
    cmbItemActionSound.ListIndex = clcItem(Index + 1).iActionSound
    ignoreClick = False
End Sub
' Spatial checkbox has been clicked
Private Sub chkSoundSpatial_Click()
    If cmbSound.ListIndex = -1 Or nodeCounter = 0 Then
        ' If no sounds or no nodes, ignore check
        chkSoundSpatial.Value = 0
        Exit Sub
    End If
    ' If checked, enable spatial sound & reset node source to 0
    If chkSoundSpatial.Value = 1 Then
        cmbSoundSpatialNode.Enabled = True
        cmbSoundSpatialNode.ListIndex = 0
    ' If unchecked, disabled spatial sound & reset node source to -1
    Else
        cmbSoundSpatialNode.Enabled = False
        cmbSoundSpatialNode.ListIndex = -1
        clcSoundSpatial(cmbSound.ListIndex + 1).sSpatial = -1
    End If
End Sub
' Item action sound combo clicked, so change value in collection if necessary
Private Sub cmbItemActionSound_Click()
    If ignoreClick Then Exit Sub
    If cmbItemName.ListIndex = -1 Then
        cmbItemActionSound.ListIndex = -1
        Exit Sub
    End If
    clcItem(cmbItemName.ListIndex + 1).iActionSound = cmbItemActionSound.ListIndex
End Sub
' Item name combo clicked so load properties of item
Private Sub cmbItemName_Click()
    loadItemProperties cmbItemName.ListIndex
End Sub
' Item name sound combo clicked, so change value in collection if necessary
Private Sub cmbItemNameSound_Click()
    If ignoreClick Then Exit Sub
    If cmbItemName.ListIndex = -1 Then
        cmbItemNameSound.ListIndex = -1
        Exit Sub
    End If
    clcItem(cmbItemName.ListIndex + 1).iNameSound = cmbItemNameSound.ListIndex
End Sub
' Item type combo clicked, so change value in collection if necessary
Private Sub cmbItemType_Click()
    If ignoreClick Then Exit Sub
    If cmbItemName.ListIndex = -1 Then
        cmbItemType.ListIndex = -1
        Exit Sub
    End If
    clcItem(cmbItemName.ListIndex + 1).iType = cmbItemType.ListIndex
End Sub
' NPC action sound combo clicked, so change value in collection if necessary
Private Sub cmbNPCActionSound_Click()
    If ignoreClick Then Exit Sub
    If cmbNPCName.ListIndex = -1 Then
        cmbNPCActionSound.ListIndex = -1
        Exit Sub
    End If
    clcNPC(cmbNPCName.ListIndex + 1).cActionSound = cmbNPCActionSound.ListIndex
End Sub
' NPC name combo clicked, so load properties of NPC
Private Sub cmbNPCName_Click()
    loadCharacterProperties cmbNPCName.ListIndex
End Sub
' NPC name sound combo clicked, so change value in collection if necessary
Private Sub cmbNPCNameSound_Click()
    If ignoreClick Then Exit Sub
    If cmbNPCName.ListIndex = -1 Then
        cmbNPCNameSound.ListIndex = -1
        Exit Sub
    End If
    clcNPC(cmbNPCName.ListIndex + 1).cNameSound = cmbNPCNameSound.ListIndex
End Sub
' NPC type combo clicked, so change value in collection if necessary
Private Sub cmbNPCType_Click()
    If ignoreClick Then Exit Sub
    If cmbNPCName.ListIndex = -1 Then
        cmbNPCType.ListIndex = -1
        Exit Sub
    End If
    clcNPC(cmbNPCName.ListIndex + 1).cType = cmbNPCType.ListIndex
End Sub
' Sound combo clicked, so load spatial properties
Private Sub cmbSound_Click()
    ignoreClick = True
    Dim n As Integer
    n = clcSoundSpatial(cmbSound.ListIndex + 1).sSpatial
    If n = -1 Then
        chkSoundSpatial.Value = 0
        cmbSoundSpatialNode.Enabled = False
    Else
        chkSoundSpatial.Value = 1
        cmbSoundSpatialNode.Enabled = True
        cmbSoundSpatialNode.ListIndex = n
    End If
    ignoreClick = False
End Sub
' Sound spatial node combo clicked, so change value in collection if necessary
Private Sub cmbSoundSpatialNode_Click()
    If ignoreClick Then Exit Sub
    If cmbSound.ListIndex = -1 Then
        cmbSoundSpatialNode.ListIndex = -1
        Exit Sub
    End If
    clcSoundSpatial(cmbSound.ListIndex + 1).sSpatial = Val(cmbSoundSpatialNode.ListIndex)
End Sub
' New item + button clicked, so ask for new item name
Private Sub cmdItemNew_Click()
    Dim s As String
    s = InputBox("Item name:", "New Item")
    If s <> "" Then addItem s
End Sub
' Set end node
Private Sub cmdNodeEnd_Click()
    If current = -1 Then Exit Sub
    If mapEnd <> -1 Then
        If opNode(mapEnd).BackColor <> vbGreen Then opNode(mapEnd).BackColor = Option1.BackColor
    End If
    mapEnd = current
End Sub
' Copy BMP to map image folder
Private Sub cmdNodeImage_Click()
    If current = -1 Then Exit Sub
    Dlg.Filter = "BMP Files (*.bmp)|*.bmp"
    Dlg.DialogTitle = "Select an image for node"
    ' Use open file dialog
    On Error GoTo CancelError
    Dlg.ShowOpen
    On Error GoTo 0
    FileCopy Dlg.FileName, imagePath & Dlg.FileTitle
    strImage(optionToNode(current)) = Dlg.FileTitle
    txtNodeImage.Text = "...\" & Dlg.FileTitle
    Exit Sub
' Error handling
SoundExists:
    MsgBox "Sound already part of map!", vbCritical, "Sound Import Error"
CancelError:
End Sub
' New NPC + button clicked, so ask for new NPC name
Private Sub cmdNPCNew_Click()
    Dim s As String
    s = InputBox("Character Name:", "New Character")
    If s <> "" Then addNPC s
End Sub
' New Sound + button clicked, so open file dialog should be shown
Private Sub cmdSoundNew_Click()
    Dlg.FileName = ""
    Dlg.Flags = cdlOFNFileMustExist
    Dlg.Filter = "Sound Files (*.mp3;*.wav)|*.mp3;*.wav"
    Dlg.DialogTitle = "Select a sound to add"
    On Error GoTo CancelError
    Dlg.ShowOpen
    On Error GoTo 0
    Dim i As Integer
    ' If sound part of project, exit due to error
    For i = 0 To cmbSound.ListCount - 1
        If cmbSound.List(i) = Dlg.FileTitle Then GoTo SoundExists
    Next
    ' Copy file to project
    FileCopy Dlg.FileName, soundPath & Dlg.FileTitle
    addSound Dlg.FileTitle
    Exit Sub
SoundExists:
    MsgBox "Sound already part of map!", vbCritical, "Sound Import Error"
CancelError:
End Sub
' Add new node to map with given radio button index
Private Sub addNewNode(optionIndex As Integer)
    ' Add to option to node mapping
    optionToNode(optionIndex) = nodeCounter
    Dim i As Integer
    ' Add to node to option mapping
    ReDim Preserve nodeToOption(0 To nodeCounter)
    nodeToOption(nodeCounter) = optionIndex
    ' Add music to node
    ReDim Preserve clcNodeMusic(0 To nodeCounter)
    For i = 0 To cmbSound.ListCount - 1
        Dim newMB As New MusicBoolean
        newMB.mbBoolean = False
        clcNodeMusic(nodeCounter).Add newMB
        Set newMB = Nothing
    Next
    ' Add item to node
    ReDim Preserve clcNodeItem(0 To nodeCounter)
    ReDim Preserve clcNodeRequiredItem(0 To nodeCounter)
    For i = 0 To cmbItemName.ListCount - 1
        Dim newItemPerc As New ItemPercent
        newItemPerc.ipPercent = 0
        clcNodeItem(nodeCounter).Add newItemPerc
        Set newItemPerc = Nothing
        Dim newItemBool As New ItemBoolean
        newItemBool.ibBoolean = False
        clcNodeRequiredItem(nodeCounter).Add newItemBool
        Set newItemBool = Nothing
    Next
    ' Add NPC to node
    ReDim Preserve clcNodeNPC(0 To nodeCounter)
    For i = 0 To cmbNPCName.ListCount - 1
        Dim newNPCPerc As New NPCPercent
        newNPCPerc.npPercent = 0
        clcNodeNPC(nodeCounter).Add newNPCPerc
        Set newNPCPerc = Nothing
    Next
    ReDim Preserve strImage(0 To nodeCounter)
    ReDim Preserve nh(0 To nodeCounter)
    ' Initialize node header
    Dim n As Integer, s As Integer, e As Integer, w As Integer
    ' Get neighbor indices
    n = getNorth(optionIndex)
    s = getSouth(optionIndex)
    e = getEast(optionIndex)
    w = getWest(optionIndex)
    ' Translate to node index
    If n <> -1 Then
        If nodeUsed(n) Then
            n = optionToNode(n)
            nh(n).nSouth = optionToNode(optionIndex)
        Else
            n = -1
        End If
    End If
    If s <> -1 Then
        If nodeUsed(s) Then
            s = optionToNode(s)
            nh(s).nNorth = optionToNode(optionIndex)
        Else
            s = -1
        End If
    End If
    If e <> -1 Then
        If nodeUsed(e) Then
            e = optionToNode(e)
            nh(e).nWest = optionToNode(optionIndex)
        Else
            e = -1
        End If
    End If
    If w <> -1 Then
        If nodeUsed(w) Then
            w = optionToNode(w)
            nh(w).nEast = optionToNode(optionIndex)
        Else
            w = -1
        End If
    End If
    ' Set node header properties
    With nh(nodeCounter)
        .nEast = e
        .nImage = ""
        .nItemCount = 0
        .nMusicCount = 0
        .nNorth = n
        .nNPCCount = 0
        .nSouth = s
        .nWest = w
    End With
    ' GUI stuff
    cmbSoundSpatialNode.addItem nodeCounter
    nodeCounter = nodeCounter + 1
    mnuFileSave.Enabled = True
    mnuFilePlay.Enabled = True
End Sub
' Adds sound to map
Private Sub addSound(s As String)
    ' Add to all necessary combo & list boxes
    cmbSound.addItem s
    lstNodeMusic.addItem s
    cmbNPCNameSound.addItem s
    cmbNPCActionSound.addItem s
    cmbItemNameSound.addItem s
    cmbItemActionSound.addItem s
    ' Add to node music collection
    Dim i As Integer
    For i = 0 To nodeCounter - 1
        Dim mb As New MusicBoolean
        mb.mbBoolean = False
        clcNodeMusic(i).Add mb
        Set mb = Nothing
    Next
    ' Add to spatial collection
    Dim newSound As New Sound
    newSound.sName = s
    newSound.sSpatial = -1
    clcSoundSpatial.Add newSound
    Set newSound = Nothing
    soundCounter = soundCounter + 1
End Sub
' Adds NPC to map
Private Sub addNPC(s As String)
    ' Add to combo box, list box, & character collection
    cmbNPCName.addItem s
    lstNodeNPC.addItem s
    Dim newNPC As New Character
    ' Set NPC properties
    With newNPC
        .cActionSound = -1
        .cDefMax = 0
        .cDefMin = 0
        .cHPMax = 0
        .cHPMin = 0
        .cHPMax = 0
        .cName = s
        .cNameSound = -1
        .cRunPerc = 0
        .cStrMax = 0
        .cStrMin = 0
        .cType = -1
    End With
    clcNPC.Add newNPC
    Set newNPC = Nothing
    ' Expand item % list of NPC
    ReDim Preserve clcNPCItem(0 To cmbNPCName.ListCount - 1)
    Dim i As Integer
    ' Set all items to start with 0%
    For i = 0 To lstNPCItem.ListCount - 1
        Dim newItemPerc As New ItemPercent
        newItemPerc.ipPercent = 0
        clcNPCItem(cmbNPCName.ListCount - 1).Add newItemPerc
        Set newItemPerc = Nothing
    Next
    ' Each node will start with a 0% chance of having new character
    For i = 0 To nodeCounter - 1
        Dim newNPCPerc As New NPCPercent
        newNPCPerc.npPercent = 0
        clcNodeNPC(i).Add newNPCPerc
        Set newNPCPerc = Nothing
    Next
End Sub
' Add item to map
Private Sub addItem(s As String)
    ' Add to necessary combo & list boxes
    cmbItemName.addItem s
    lstNodeItem.addItem s
    lstNodeRequiredItem.addItem s
    lstNPCItem.addItem s
    ' Set new item properties
    Dim newItem As New Item
    With newItem
        .iActionSound = -1
        .iName = s
        .iNameSound = -1
        .iType = -1
        .iValue = 0
    End With
    ' Add to collection
    clcItem.Add newItem
    Set newItem = Nothing
    Dim i As Integer
    ' Each NPC has a 0% chance of having new item
    For i = 0 To cmbNPCName.ListCount - 1
        Dim newItemPerc As New ItemPercent
        newItemPerc.ipPercent = 0
        clcNPCItem(i).Add newItemPerc
        Set newItemPerc = Nothing
    Next
    ' Each node has a 0% chance of having new item
    For i = 0 To nodeCounter - 1
        Dim newItemPerce As New ItemPercent
        newItemPerce.ipPercent = 0
        clcNodeItem(i).Add newItemPerce
        Set newItemPerce = Nothing
        Dim newItemBool As New ItemBoolean
        newItemBool.ibBoolean = False
        clcNodeRequiredItem(i).Add newItemBool
        Set newItemBool = Nothing
    Next
End Sub
' Form load so setup GUI
Private Sub Form_Load()
    Reset
    With cmbNPCType
        .addItem "Enemy"
        .addItem "Friend"
        .addItem "Vendor"
        .addItem "Leprechaun"
    End With
    With cmbItemType
        .addItem "Weapon"
        .addItem "Armor"
        .addItem "Potion"
        .addItem "Gold"
        .addItem "Special"
    End With
    mnuFileSave.Enabled = False
    mnuFilePlay.Enabled = False
    fraMap(1).Enabled = False
    fraNode.Enabled = False
    fraNPC.Enabled = False
    fraItem.Enabled = False
    fraSound.Enabled = False
End Sub
' Get index # from row & column
Private Function getIndex(row As Integer, col As Integer) As Integer
    If row >= MAX_ROW Or row < 0 Or col >= MAX_COL Or col < 0 Then
        getIndex = -1
    Else
        getIndex = MAX_COL * row + col
    End If
End Function
' Get row # from index
Private Function getRow(Index As Integer) As Integer
    getRow = Int(Index / MAX_COL)
End Function
' Get column # from index
Private Function getCol(Index As Integer) As Integer
    getCol = Index Mod MAX_COL
End Function
' Get index of northern neighbor
Private Function getNorth(Index As Integer) As Integer
    getNorth = getIndex(getRow(Index) - 1, getCol(Index))
End Function
' Get index of southern neighbor
Private Function getSouth(Index As Integer) As Integer
    getSouth = getIndex(getRow(Index) + 1, getCol(Index))
End Function
' Get index of eastern neighbor
Private Function getEast(Index As Integer) As Integer
    getEast = getIndex(getRow(Index), getCol(Index) + 1)
End Function
' Get index of western neighbor
Private Function getWest(Index As Integer) As Integer
    getWest = getIndex(getRow(Index), getCol(Index) - 1)
End Function
' Reset map
Private Sub Reset()
    ' Reset map data
    mapStart = -1
    mapEnd = -1
    ' Next selection will be start node
    blStart = True
    ' Reset all nodes
    Dim i As Integer
    For i = 0 To MAX_NODE - 1
        opNode(i).Value = False
        opNode(i).Visible = True
        opNode(i).BackColor = Option1.BackColor
        nodeUsed(i) = False
    Next
    ' Reset collections
    ReDim clcNodeMusic(0 To 0)
    Set clcNodeMusic(0) = Nothing
    ReDim clcNodeItem(0 To 0)
    Set clcNodeItem(0) = Nothing
    ReDim clcNodeNPC(0 To 0)
    Set clcNodeNPC(0) = Nothing
    ReDim clcNodeRequiredItem(0 To 0)
    Set clcNodeRequiredItem(0) = Nothing
    Set clcNPC = Nothing
    ReDim clcNPCItem(0 To 0)
    Set clcNPCItem(0) = Nothing
    Set clcItem = Nothing
    Set clcSoundSpatial = Nothing
    ' Rest other GUI elements
    ' Sound frame
    cmbSound.Clear
    chkSoundSpatial.Value = 0
    cmbSoundSpatialNode.Clear
    cmbSoundSpatialNode.Enabled = False
    ' NPC frame
    cmbNPCName.Clear
    lstNPCItem.Clear
    cmbNPCNameSound.Clear
    cmbNPCActionSound.Clear
    sldNPCItem.Value = 0
    lblNPCItemPercent.Caption = "0 %"
    txtNPCStrMin.Text = ""
    txtNPCStrMax.Text = ""
    txtNPCDefMin.Text = ""
    txtNPCDefMax.Text = ""
    txtNPCHPMin.Text = ""
    txtNPCHPMax.Text = ""
    txtNPCRun.Text = ""
    ' Node Properties frame
    lstNodeMusic.Clear
    lstNodeItem.Clear
    lstNodeNPC.Clear
    lstNodeRequiredItem.Clear
    sldNodeItem.Value = 0
    sldNodeNPC.Value = 0
    lblNodeItemPercent.Caption = "0 %"
    lblNodeNPCPercent.Caption = "0 %"
    ' Item frame
    cmbItemName.Clear
    txtItemValue.Text = ""
    cmbItemNameSound.Clear
    cmbItemActionSound.Clear
    ' Reset cursor
    nodeCounter = 0
    soundCounter = 0
    current = -1
    ' Reset misc
    fraNode.Caption = "Node Properties"
End Sub
' Node item list box clicked, so update collection
Private Sub lstNodeItem_Click()
    If ignoreClick Then Exit Sub
    Dim i As Integer, n As Integer
    i = lstNodeItem.ListIndex
    If current = -1 Then
        lstNodeItem.ListIndex = -1
        Exit Sub
    End If
    n = optionToNode(current)
    sldNodeItem.Value = clcNodeItem(n).Item(i + 1).ipPercent
    lblNodeItemPercent.Caption = sldNodeItem.Value & " %"
End Sub
' Node music list box clicked, so update collection
Private Sub lstNodeMusic_Click()
    If ignoreClick Then Exit Sub
    Dim i As Integer, n As Integer
    i = lstNodeMusic.ListIndex
    If current = -1 Then
        lstNodeMusic.Selected(i) = False
        Exit Sub
    End If
    n = optionToNode(current)
    clcNodeMusic(n).Item(i + 1).mbBoolean = lstNodeMusic.Selected(i)
End Sub
' Node NPC list box clicked, so load information from collection into slider
Private Sub lstNodeNPC_Click()
    If ignoreClick Then Exit Sub
    Dim i As Integer, n As Integer
    i = lstNodeNPC.ListIndex
    If current = -1 Then
        lstNodeNPC.ListIndex = -1
        Exit Sub
    End If
    n = optionToNode(current)
    sldNodeNPC.Value = clcNodeNPC(n).Item(i + 1).npPercent
    lblNodeNPCPercent.Caption = sldNodeNPC.Value & " %"
End Sub
' Node req item list box clicked, so update collection
Private Sub lstNodeRequiredItem_Click()
    If ignoreClick Then Exit Sub
    Dim i As Integer, n As Integer
    i = lstNodeRequiredItem.ListIndex
    If current = -1 Then
        lstNodeRequiredItem.Selected(i) = False
        Exit Sub
    End If
    n = optionToNode(current)
    clcNodeRequiredItem(n).Item(i + 1).ibBoolean = lstNodeRequiredItem.Selected(i)
End Sub
' NPC item list box clicked, so load information from collection into slider
Private Sub lstNPCItem_Click()
    If ignoreClick Then Exit Sub
    Dim i As Integer, n As Integer
    i = lstNPCItem.ListIndex
    n = cmbNPCName.ListIndex
    If i = -1 Or n = -1 Then
        lstNPCItem.ListIndex = -1
        Exit Sub
    End If
    sldNPCItem.Value = clcNPCItem(n).Item(i + 1).ipPercent
    lblNPCItemPercent.Caption = sldNPCItem.Value & " %"
End Sub
' File->Exit clicked
Private Sub mnuFileExit_Click()
    Unload frmMain
    End
End Sub
' File->New clicked, make new map
Private Sub mnuFileNew_Click()
    ' Setup save dialog
    Dlg.FileName = ""
    Dlg.Flags = cdlOFNOverwritePrompt
    Dlg.Filter = "Maps (*.map)|*.map"
    Dlg.DialogTitle = "Create map file"
    On Error GoTo CancelError
    Dlg.ShowSave
    On Error GoTo 0
    ' Add new map file and directories
    mapPath = Left(Dlg.FileName, InStrRev(Dlg.FileName, "\"))
    mapName = Left(Dlg.FileTitle, InStrRev(Dlg.FileTitle, ".map") - 1)
    soundPath = mapPath & mapName & ".sounds\"
    imagePath = mapPath & mapName & ".images\"
    On Error Resume Next
    MkDir soundPath
    On Error GoTo 0
    On Error Resume Next
    MkDir imagePath
    On Error GoTo 0
    ' Reset GUI
    Reset
    fraMap(1).Enabled = True
    fraNode.Enabled = True
    fraNPC.Enabled = True
    fraItem.Enabled = True
    fraSound.Enabled = True
    Exit Sub
CancelError:
End Sub
' Read map file from disk
Private Sub ReadMap(path As String)
    ' Arbitrary variables
    Dim i As Integer, j As Integer, r As Integer, c As Integer
    
    ' Open file
    Open path For Binary As #1
    
    ' Get map header
    Get #1, , mh
    
    ' Sound data
    If mh.mNodeCount > 0 Then ReDim clcNodeMusic(0 To mh.mNodeCount - 1)
    Dim sd() As SOUNDDATA
    If mh.mSoundCount > 0 Then ReDim sd(0 To mh.mSoundCount - 1)
    Get #1, , sd
    For i = 0 To mh.mSoundCount - 1
        addSound Trim(sd(i).sName)
        clcSoundSpatial(i + 1).sSpatial = sd(i).sNode
        ' Node property
        For j = 0 To mh.mNodeCount - 1
            Dim newMusic As New MusicBoolean
            newMusic.mbBoolean = False
            clcNodeMusic(j).Add newMusic
            Set newMusic = Nothing
        Next
    Next
    
    ' Item data
    If mh.mNodeCount > 0 Then ReDim clcNodeItem(0 To mh.mNodeCount - 1)
    If mh.mNodeCount > 0 Then ReDim clcNodeRequiredItem(0 To mh.mNodeCount - 1)
    Dim id() As ITEMDATA
    If mh.mItemCount > 0 Then ReDim id(0 To mh.mItemCount - 1)
    Get #1, , id
    For i = 0 To mh.mItemCount - 1
        addItem Trim(id(i).iName)
        clcItem(i + 1).iType = id(i).iType
        clcItem(i + 1).iValue = id(i).iValue
        clcItem(i + 1).iNameSound = Trim(id(i).iNameSound)
        clcItem(i + 1).iActionSound = Trim(id(i).iActionSound)
        ' Node property
        For j = 0 To mh.mNodeCount - 1
            Dim newItem As New ItemPercent, newItemB As New ItemBoolean
            newItem.ipPercent = 0
            newItemB.ibBoolean = False
            clcNodeItem(j).Add newItem
            clcNodeRequiredItem(j).Add newItemB
            Set newItem = Nothing
            Set newItemB = Nothing
        Next
    Next
    
    ' NPC data
    If mh.mNodeCount > 0 Then ReDim clcNodeNPC(0 To mh.mNodeCount - 1)
    Dim cd As NPCDATA, ci() As NPCITEM
    For i = 1 To mh.mNPCCount
        Get #1, , cd
        addNPC Trim(cd.cName)
        clcNPC(i).cType = cd.cType
        clcNPC(i).cStrMin = cd.cStrMin
        clcNPC(i).cStrMax = cd.cStrMax
        clcNPC(i).cDefMin = cd.cDefMin
        clcNPC(i).cDefMax = cd.cDefMax
        clcNPC(i).cHPMin = cd.cHPMin
        clcNPC(i).cHPMax = cd.cHPMax
        clcNPC(i).cRunPerc = cd.cRunPerc
        clcNPC(i).cNameSound = Trim(cd.cNameSound)
        clcNPC(i).cActionSound = Trim(cd.cActionSound)
        If cd.cItemCount > 0 Then
            ReDim ci(0 To cd.cItemCount - 1)
            Get #1, , ci
            For j = 0 To cd.cItemCount - 1
                clcNPCItem(i - 1).Item(ci(j).cItem + 1).ipPercent = ci(j).cPercent
            Next
        End If
        ' Node property
        For j = 0 To mh.mNodeCount - 1
            Dim newNPC As New NPCPercent
            newNPC.npPercent = 0
            clcNodeNPC(j).Add newNPC
            Set newNPC = Nothing
        Next
    Next
    
    ' Node data
    nodeCounter = mh.mNodeCount
    If mh.mNodeCount > 0 Then ReDim strImage(0 To mh.mNodeCount - 1)
    If mh.mNodeCount > 0 Then ReDim nodeToOption(0 To mh.mNodeCount - 1)
    If mh.mNodeCount > 0 Then ReDim nh(0 To mh.mNodeCount - 1)
    Dim ns As NODESOUND, ni As NODEITEM, nc As NODENPC, nr As NODEREQITEM
    For i = 0 To mh.mNodeCount - 1
        Get #1, , nh(i)
        strImage(i) = Trim(nh(i).nImage)
        cmbSoundSpatialNode.addItem i
        ignoreClick = True
        opNode(nh(i).nIndex).Visible = True
        Dim n As Integer, s As Integer, e As Integer, w As Integer
        n = getNorth(nh(i).nIndex)
        s = getSouth(nh(i).nIndex)
        e = getEast(nh(i).nIndex)
        w = getWest(nh(i).nIndex)
        If n <> -1 Then opNode(n).Visible = True
        If s <> -1 Then opNode(s).Visible = True
        If e <> -1 Then opNode(e).Visible = True
        If w <> -1 Then opNode(w).Visible = True
        opNode(nh(i).nIndex).Value = True
        ignoreClick = False
        optionToNode(nh(i).nIndex) = i
        nodeToOption(i) = nh(i).nIndex
        nodeUsed(nh(i).nIndex) = True
        ' Music
        For j = 0 To nh(i).nMusicCount - 1
            Get #1, , ns
            clcNodeMusic(i).Item(ns.nSound + 1).mbBoolean = True
        Next
        ' Items
        For j = 0 To nh(i).nItemCount - 1
            Get #1, , ni
            clcNodeItem(i).Item(ni.nItem + 1).ipPercent = ni.nPercent
        Next
        ' NPCs
        For j = 0 To nh(i).nNPCCount - 1
            Get #1, , nc
            clcNodeNPC(i).Item(nc.nNPC + 1).npPercent = nc.nPercent
        Next
        ' Required items
        For j = 0 To nh(i).nReqItemCount - 1
            Get #1, , nr
            clcNodeRequiredItem(i).Item(nr.nReqItem + 1).ibBoolean = True
        Next
    Next
    
    ' Set start/end nodes
    mapStart = nodeToOption(mh.mStart)
    blStart = (mapStart = -1)
    mapEnd = nodeToOption(mh.mEnd)
    
    ' Set cursor
    opNode(mapEnd).BackColor = vbRed
    moveCursor mapStart

    ' Close file
    Close #1
End Sub
' Write map data to disk
Private Sub WriteMap(path As String)
    ' Arbitrary variables
    Dim i As Integer, j As Integer, r As Integer, c As Integer
    
    ' Open file
    Open path For Binary As #1
    
    ' Map header
    mh.mStart = optionToNode(mapStart)
    mh.mEnd = optionToNode(mapEnd)
    mh.mNodeCount = nodeCounter
    mh.mNPCCount = cmbNPCName.ListCount
    mh.mSoundCount = cmbSound.ListCount
    mh.mItemCount = cmbItemName.ListCount
    Put #1, , mh
    
    ' Sound data
    Dim sd As SOUNDDATA
    For i = 0 To mh.mSoundCount - 1
        sd.sName = clcSoundSpatial(i + 1).sName & Chr(0)
        j = clcSoundSpatial(i + 1).sSpatial
        sd.sNode = j
        sd.sSpatial = (j <> -1)
        If sd.sSpatial Then
            sd.sXCoord = getCol(nodeToOption(j)) - getCol(mapStart)
            sd.sYCoord = getRow(mapStart) - getRow(nodeToOption(j))
        Else
            sd.sXCoord = 0
            sd.sYCoord = 0
        End If
        Put #1, , sd
    Next
    
    ' Item data
    Dim id As ITEMDATA
    For i = 1 To mh.mItemCount
        id.iName = clcItem(i).iName & Chr(0)
        id.iType = clcItem(i).iType
        id.iValue = clcItem(i).iValue
        id.iNameSound = clcItem(i).iNameSound
        id.iActionSound = clcItem(i).iActionSound
        Put #1, , id
    Next
    
    ' NPC data
    Dim cd As NPCDATA, ci() As NPCITEM
    For i = 1 To mh.mNPCCount
        cd.cName = clcNPC(i).cName & Chr(0)
        cd.cType = clcNPC(i).cType
        cd.cStrMin = clcNPC(i).cStrMin
        cd.cStrMax = clcNPC(i).cStrMax
        cd.cDefMin = clcNPC(i).cDefMin
        cd.cDefMax = clcNPC(i).cDefMax
        cd.cHPMin = clcNPC(i).cHPMin
        cd.cHPMax = clcNPC(i).cHPMax
        cd.cRunPerc = clcNPC(i).cRunPerc
        cd.cNameSound = clcNPC(i).cNameSound
        cd.cActionSound = clcNPC(i).cActionSound
        ' Items
        c = 0
        For j = 1 To mh.mItemCount
            r = clcNPCItem(i - 1).Item(j).ipPercent
            If r > 0 Then
                ReDim Preserve ci(0 To c)
                ci(c).cItem = j - 1
                ci(c).cPercent = r
                c = c + 1
            End If
        Next
        cd.cItemCount = c
        Put #1, , cd
        If c > 0 Then Put #1, , ci
    Next
    
    ' Node data
    Dim ns() As NODESOUND, ni() As NODEITEM, nc() As NODENPC, nr() As NODEREQITEM
    For i = 0 To mh.mNodeCount - 1
        nh(i).nIndex = nodeToOption(i)
        ' Music
        c = 0
        For j = 1 To mh.mSoundCount
            If clcNodeMusic(i).Item(j).mbBoolean Then
                ReDim Preserve ns(0 To c)
                ns(c).nSound = j - 1
                c = c + 1
            End If
        Next
        nh(i).nMusicCount = c
        ' Items
        c = 0
        For j = 1 To mh.mItemCount
            r = clcNodeItem(i).Item(j).ipPercent
            If r > 0 Then
                ReDim Preserve ni(0 To c)
                ni(c).nItem = j - 1
                ni(c).nPercent = r
                c = c + 1
            End If
        Next
        nh(i).nItemCount = c
        ' NPCs
        c = 0
        For j = 1 To mh.mNPCCount
            r = clcNodeNPC(i).Item(j).npPercent
            If r > 0 Then
                ReDim Preserve nc(0 To c)
                nc(c).nNPC = j - 1
                nc(c).nPercent = r
                c = c + 1
            End If
        Next
        nh(i).nNPCCount = c
        ' Required items
        c = 0
        For j = 1 To mh.mItemCount
            If clcNodeRequiredItem(i).Item(j).ibBoolean Then
                ReDim Preserve nr(0 To c)
                nr(c).nReqItem = j - 1
                c = c + 1
            End If
        Next
        nh(i).nReqItemCount = c
        nh(i).nImage = nh(i).nImage & Chr(0)
        Put #1, , nh(i)
        If nh(i).nMusicCount > 0 Then Put #1, , ns
        If nh(i).nItemCount > 0 Then Put #1, , ni
        If nh(i).nNPCCount > 0 Then Put #1, , nc
        If nh(i).nReqItemCount > 0 Then Put #1, , nr
    Next
    
    ' Close file
    Close #1
End Sub
' File->Open clicked, so open map
Private Sub mnuFileOpen_Click()
    ' Setup open dialog
    Dlg.FileName = ""
    Dlg.Flags = cdlOFNFileMustExist
    Dlg.Filter = "Maps (*.map)|*.map"
    Dlg.DialogTitle = "Open map file"
    On Error GoTo CancelError
    Dlg.ShowOpen
    On Error GoTo 0
    ' Get map path attributes
    mapPath = Left(Dlg.FileName, InStrRev(Dlg.FileName, "\"))
    mapName = Left(Dlg.FileTitle, InStrRev(Dlg.FileTitle, ".map") - 1)
    soundPath = mapPath & mapName & ".sounds\"
    imagePath = mapPath & mapName & ".images\"
    On Error Resume Next
    MkDir soundPath
    On Error GoTo 0
    On Error Resume Next
    MkDir imagePath
    On Error GoTo 0
    ' Reset GUI
    Reset
    fraMap(1).Enabled = True
    fraNode.Enabled = True
    fraNPC.Enabled = True
    fraItem.Enabled = True
    fraSound.Enabled = True
    mnuFileSave.Enabled = True
    mnuFilePlay.Enabled = True
    Dim i As Integer
    For i = 0 To MAX_NODE - 1
        opNode(i).Visible = False
    Next
    ' Load map data
    ReadMap mapPath & mapName & ".map"
    Exit Sub
CancelError:
End Sub
' File->Save & Play clicked, so save map and then play with LastCrusade.exe
Private Sub mnuFilePlay_Click()
    ' Make sure map end has been identified
    If mapEnd = -1 Then
        MsgBox "No end node defined for map!", vbCritical, "Map Error"
        Exit Sub
    End If
    WriteMap mapPath & mapName & ".map"
    Shell "LastCrusade.exe " & mapPath & mapName & ".map", vbNormalFocus
End Sub
' File->Save clicked, so save map data
Private Sub mnuFileSave_Click()
    ' Make sure map end has been identified
    If mapEnd = -1 Then
        MsgBox "No end node defined for map!", vbCritical, "Map Error"
        Exit Sub
    End If
    WriteMap mapPath & mapName & ".map"
End Sub
' Node clicked on map
Private Sub opNode_Click(Index As Integer)
    If ignoreClick Then Exit Sub

    If nodeUsed(Index) Then
        moveCursor Index
        Exit Sub
    Else
        nodeUsed(Index) = True
        addNewNode Index
    End If
    ' Start map
    Dim i As Integer
    If blStart Then
        blStart = False
        mapStart = Index
        opNode(Index).BackColor = vbBlue
        For i = 0 To MAX_NODE - 1
            If i <> Index Then opNode(i).Visible = False
        Next
    End If
    ' Make neighbors visible
    Dim r As Integer, c As Integer
    r = getRow(Index)
    c = getCol(Index)
    i = getIndex(r - 1, c)
    If i <> -1 Then opNode(i).Visible = True
    i = getIndex(r + 1, c)
    If i <> -1 Then opNode(i).Visible = True
    i = getIndex(r, c + 1)
    If i <> -1 Then opNode(i).Visible = True
    i = getIndex(r, c - 1)
    If i <> -1 Then opNode(i).Visible = True
    ' Move cursor
    moveCursor Index
End Sub
' Move cursor to node at index
Private Sub moveCursor(Index As Integer)
    If current = Index Then Exit Sub
    If current <> -1 And current = mapStart Then
        opNode(current).BackColor = vbBlue
    ElseIf current <> -1 And current = mapEnd Then
        opNode(current).BackColor = vbRed
    ElseIf current <> -1 Then
        opNode(current).BackColor = Option1.BackColor
    End If
    current = Index
    opNode(current).BackColor = vbGreen
    ' Load node properties
    fraNode.Caption = "Node #" & optionToNode(current) & " Properties"
    loadNodeProperties Index
End Sub
Private Sub opNode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If opNode(Index).Value = True Then moveCursor (Index)
End Sub
' Node item slider changed so update collection
Private Sub sldNodeItem_Change()
    Dim i As Integer, n As Integer
    i = lstNodeItem.ListIndex
    If i = -1 Or current = -1 Then
        sldNodeItem.Value = 0
        lblNodeItemPercent.Caption = sldNodeItem.Value & " %"
        Exit Sub
    End If
    n = optionToNode(current)
    clcNodeItem(n).Item(i + 1).ipPercent = sldNodeItem.Value
End Sub
' Node item slider scrolled so change label caption
Private Sub sldNodeItem_Scroll()
    lblNodeItemPercent.Caption = sldNodeItem.Value & " %"
End Sub
' Node NPC slider changed so update collection
Private Sub sldNodeNPC_Change()
    Dim i As Integer, n As Integer
    i = lstNodeNPC.ListIndex
    If i = -1 Or current = -1 Then
        sldNodeNPC.Value = 0
        lblNodeNPCPercent.Caption = sldNodeNPC.Value & " %"
        Exit Sub
    End If
    n = optionToNode(current)
    clcNodeNPC(n).Item(i + 1).npPercent = sldNodeNPC.Value
End Sub
' Node NPC slider scrolled so change label caption
Private Sub sldNodeNPC_Scroll()
    lblNodeNPCPercent.Caption = sldNodeNPC.Value & " %"
End Sub
' NPC item slider changed so update collection
Private Sub sldNPCItem_Change()
    Dim i As Integer, n As Integer
    i = lstNPCItem.ListIndex
    n = cmbNPCName.ListIndex
    If i = -1 Or n = -1 Then
        sldNPCItem.Value = 0
        lblNPCItemPercent.Caption = sldNPCItem.Value & " %"
        Exit Sub
    End If
    clcNPCItem(n).Item(i + 1).ipPercent = sldNPCItem.Value
End Sub
' NPC item slider changed so change label caption
Private Sub sldNPCItem_Scroll()
    lblNPCItemPercent.Caption = sldNPCItem.Value & " %"
End Sub

' The rest are all text box management routines
' They do two things:
' 1. Make sure a value has been entered, and then changes it to 0 if not
' 2. Updates the corresponding collection value
Private Sub txtItemValue_Change()
    If cmbItemName.ListIndex = -1 Then
        txtItemValue.Text = ""
        Exit Sub
    End If
    txtItemValue.Text = Val(txtItemValue.Text)
    clcItem(cmbItemName.ListIndex + 1).iValue = Val(txtItemValue.Text)
End Sub

Private Sub txtNPCDefMax_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCDefMax.Text = ""
        Exit Sub
    End If
    txtNPCDefMax.Text = Val(txtNPCDefMax.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cDefMax = Val(txtNPCDefMax.Text)
End Sub

Private Sub txtNPCDefMin_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCDefMin.Text = ""
        Exit Sub
    End If
    txtNPCDefMin.Text = Val(txtNPCDefMin.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cDefMin = Val(txtNPCDefMin.Text)
End Sub

Private Sub txtNPCHPMax_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCHPMax.Text = ""
        Exit Sub
    End If
    txtNPCHPMax.Text = Val(txtNPCHPMax.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cHPMax = Val(txtNPCHPMax.Text)
End Sub

Private Sub txtNPCHPMin_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCHPMin.Text = ""
        Exit Sub
    End If
    txtNPCHPMin.Text = Val(txtNPCHPMin.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cHPMin = Val(txtNPCHPMin.Text)
End Sub

Private Sub txtNPCRun_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCRun.Text = ""
        Exit Sub
    End If
    txtNPCRun.Text = Val(txtNPCRun.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cRunPerc = Val(txtNPCRun.Text)
End Sub

Private Sub txtNPCStrMax_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCStrMax.Text = ""
        Exit Sub
    End If
    txtNPCStrMax.Text = Val(txtNPCStrMax.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cStrMax = Val(txtNPCStrMax.Text)
End Sub

Private Sub txtNPCStrMin_Change()
    If cmbNPCName.ListIndex = -1 Then
        txtNPCStrMin.Text = ""
        Exit Sub
    End If
    txtNPCStrMin.Text = Val(txtNPCStrMin.Text)
    clcNPC(cmbNPCName.ListIndex + 1).cStrMin = Val(txtNPCStrMin.Text)
End Sub
