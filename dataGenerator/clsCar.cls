VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public id As Long
Public model As String
Public make As String
Public cylinders As String
Public fuel As String
Public trim As String
Public color As String
Public year As Integer
Public miles As Long
Public blueBookValue As Long
Public listPrice As Long
Public condition As String
Public transmission As String
Public listDate As Date


Public Function toString() As String

    Dim strCSV As String
    
    strCSV = Me.id & "," & Me.model & "," & Me.make & "," & Me.cylinders & "," & Me.fuel & "," & Me.trim & ","
    strCSV = strCSV & Me.color & "," & Me.year & "," & Me.miles & "," & Me.blueBookValue & "," & Me.listPrice & ","
    strCSV = strCSV & Me.condition & "," & Me.transmission & "," & Me.listDate
    
    toString = strCSV
    
End Function

