Attribute VB_Name = "makeCSVdata"
Option Compare Database
Option Explicit
    
    Public collModels As New Collection
    Public collMakes As New Collection
    Public collCylinders As New Collection
    Public collFuels As New Collection
    Public collTrims As New Collection
    Public collColors As New Collection
    Public collConditions As New Collection
    Public collTransmissions As New Collection
    Public collCars As New Collection

Public Sub GenerateData()

    Dim counter As Long
    Dim carCount As Long
    Dim strPath As String
    
    strPath = "D:\xfer\csv\cars_data_test.csv"
    carCount = 50000
    
    CreateCollections
    CreateCars (carCount)
    writeFile (strPath)
    
    Set collModels = Nothing
    Set collMakes = Nothing
    Set collCylinders = Nothing
    Set collFuels = Nothing
    Set collTrims = Nothing
    Set collColors = Nothing
    Set collConditions = Nothing
    Set collTransmissions = Nothing
    Set collCars = Nothing

End Sub

Public Sub CreateCollections()

'build collection of models

    collModels.Add "sedan"
    collModels.Add "coupe"
    collModels.Add "minivan"
    collModels.Add "pickup"
    collModels.Add "suv"
    collModels.Add "crossover"
    collModels.Add "cargo_van"
    collModels.Add "sports_car"
    collModels.Add "motorcycle"
 
 'build collection of makes
    
    collMakes.Add "ford"
    collMakes.Add "chevrolet"
    collMakes.Add "cadillac"
    collMakes.Add "honda"
    collMakes.Add "nissan"
    collMakes.Add "mercedes_benz"
    collMakes.Add "toyota"
    collMakes.Add "bmw"
    collMakes.Add "dodge"
    collMakes.Add "chrysler"
    
'build collection of cylinders
    
    collCylinders.Add "I4"
    collCylinders.Add "I6"
    collCylinders.Add "V6"
    collCylinders.Add "V8"
    
'build collection of fuels
    
    collFuels.Add "diesel"
    collFuels.Add "gas"
    collFuels.Add "electric"
    
'build collection of trims

    collTrims.Add "basic"
    collTrims.Add "touring"
    collTrims.Add "luxury"
    collTrims.Add "sport"
    collTrims.Add "special_edition"
    
 'build collection of colors
    
    collColors.Add "white"
    collColors.Add "black"
    collColors.Add "gray"
    collColors.Add "silver"
    collColors.Add "gold"
    collColors.Add "blue"
    collColors.Add "red"
    collColors.Add "yellow"
    collColors.Add "green"
    collColors.Add "brown"
    collColors.Add "orange"
    
 'build collection of conditions
 
    collConditions.Add "new"
    collConditions.Add "like_new"
    collConditions.Add "good"
    collConditions.Add "fair"
    collConditions.Add "poor"
    
'build collection of transmissions
    
    collTransmissions.Add "manual"
    collTransmissions.Add "automatic"

End Sub

Public Sub CreateCars(recordCount As Long)

    Dim car As clsCar
    Dim markUpPercent As Integer
    Dim counter As Long
    
    markUpPercent = 30
    
    For counter = 1 To recordCount
    
        Set car = New clsCar
        
        With car
        
            .id = counter
            .model = GetRandomFromCollection(collModels)
            .make = GetRandomFromCollection(collMakes)
            .cylinders = GetRandomFromCollection(collCylinders)
            .fuel = GetRandomFromCollection(collFuels)
            .trim = GetRandomFromCollection(collTrims)
            .color = GetRandomFromCollection(collColors)
            .condition = GetRandomFromCollection(collConditions)
            .transmission = GetRandomFromCollection(collTransmissions)
            .blueBookValue = LongRandBetween(3000, 50000)
            .miles = LongRandBetween(25, 200000)
            .year = RandBetween(2000, 2020)
            .listPrice = Round(car.blueBookValue * (1 + (markUpPercent / 100)), 2)
            .listDate = GetRndDate("12/31/2017", "9/1/2020")
            
        End With
        
        collCars.Add car
        
    Next
    
    Set car = Nothing

End Sub
Public Sub writeFile(strPath)

    Dim fso As Object
    Dim oFile As Object
    Dim counter As Long
    Dim strHeaders As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(strPath)
    
    strHeaders = "id,model,make,cylinders,fuel,trim,color,year,miles,blue_book,list_price,condition,transmission,list_date"
    
    oFile.WriteLine strHeaders
    
    For counter = 1 To collCars.Count
    
        oFile.WriteLine collCars.Item(counter).toString
        
    Next
    
    oFile.Close
      
    Set fso = Nothing
    Set oFile = Nothing

End Sub

Private Function GetRandomFromCollection(coll As Collection) As String

    Dim random As Integer
    
    random = RandBetween(1, coll.Count)
    
    GetRandomFromCollection = coll(random)

End Function

Public Function RandBetween(lower As Integer, upper As Integer) As Integer

    RandBetween = Int(lower + Rnd * (upper - lower + 1))

End Function

Public Function LongRandBetween(lower As Long, upper As Long) As Long

    LongRandBetween = (lower + Rnd * (upper - lower + 1))

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRndDate
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Get a random date between 2 dates
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' dtStartDate : Minimum Date value
' dtEndDate   : Maximum Date value
'
' Usage:
' ~~~~~~
' dtRnd = GetRndDate(#12/01/2002#, #01/05/2015#)
'           Will return a random date between #12/01/2002# and #01/05/2015#
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-02-07              Initial Release
' 2         2017-02-09              Added check that Start is < to End
'                                   Added Comments to code
' 3         2018-09-20              Updated Copyright
'---------------------------------------------------------------------------------------
Function GetRndDate(dtStartDate As Date, dtEndDate As Date) As Date
    On Error GoTo Error_Handler
    Dim dtTmp As Date
 
    'Swap the dates if dtStartDate is after dtEndDate
    If dtStartDate > dtEndDate Then
        dtTmp = dtStartDate
        dtStartDate = dtEndDate
        dtEndDate = dtTmp
    End If
 
    Randomize
    GetRndDate = DateAdd("d", Int((DateDiff("d", dtStartDate, dtEndDate) + 1) * Rnd), dtStartDate)
 
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: GetRndDate" & vbCrLf & _
           "Error Description: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

