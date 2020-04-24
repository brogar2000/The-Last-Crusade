VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Character class for storing data in a collection
Option Explicit
Public cName As String, cType As Integer, cStrMin As Integer, cStrMax As Integer, cDefMin As Integer, cDefMax As Integer
Public cHPMin As Integer, cHPMax As Integer, cNameSound As Integer, cActionSound As Integer
Public cRunPerc As Integer
