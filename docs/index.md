## IsInArray

Diese Funktion überprüft, ob ein gegebener String (`val`) in einem Array (`arr`) enthalten ist. Gibt `True` zurück, wenn der String im Array gefunden wurde, andernfalls `False`.

```vbnet
Public Function IsInArray(val As String, arr As Variant) As Boolean
    'Prüft, ob der String val im Array arr enthalten ist.
    'Argumente:
    '    - val (String): Zu suchender Wert.
    '    - arr (Variant): Array, in dem gesucht wird.
    'Rückgabewert: True (Boolean), wenn val im Array gefunden wird, sonst False.
    Dim n As Integer
    For n = 0 To UBound(arr)
        If UCase(val) = UCase(arr(n)) Then
            IsInArray = True
            Exit Function
        End If
    Next n
    IsInArray = False
End Function
```

## ExtractFileNameFromPath

Extrahiert aus einem vollständigen Dateipfad den Dateinamen und gibt optional die Dateiendung mit zurück.

```vbnet
Public Function ExtractFileNameFromPath(path As String, includeFileExtension As Boolean) As String
    'Extrahiert den Dateinamen aus einem vollständigen Pfad.
    'Argumente:
    '    - path (String): Vollständiger Dateipfad.
    '    - includeFileExtension (Boolean): True, wenn die Dateiendung enthalten sein soll.
    'Rückgabewert: Dateiname als String (mit oder ohne Endung).

    Dim fileNameExtension As String
    fileNameExtension = Mid(path, InStrRev(path, "\") + 1) 'Dateiname inkl. Dateiendung extrahieren

    'Prüfen, ob Dateiendung auch berücksichtigt werden soll
    If includeFileExtension = True Then
        ExtractFileNameFromPath = fileNameExtension
    Else
        'Dateiendung entfernen, indem 7 Zeichen von Rechts entfernt werden
        'Alle möglichen SolidWorks Dateiendungen sind 6 Zeichen lang plus der Punkt
        ExtractFileNameFromPath = Left(fileNameExtension, Len(fileNameExtension) - 7)
    End If

End Function
```

## SortDictionaryByValue

Sortiert ein `Scripting.Dictionary` nach seinen Werten auf- oder absteigend und gibt das sortierte Dictionary zurück.

```vbnet
Public Function SortDictionaryByValue(dict As Object _
                    , Optional sortorder As XlSortOrder = xlAscending) As Object
    'Sortiert ein Dictionary nach seinen Werten auf- oder absteigend.
    'Argumente:
    '    - dict (Object): Zu sortierendes Dictionary.
    '    - sortorder (XlSortOrder, optional): Sortierreihenfolge (Standard: aufsteigend).
    'Rückgabewert: Sortiertes Dictionary (Object).
    
    On Error GoTo eh
    
    Dim arrayList As Object
    Set arrayList = CreateObject("System.Collections.ArrayList")
    
    Dim dictTemp As Object
    Set dictTemp = CreateObject("Scripting.Dictionary")
   
    ' Put values in ArrayList and sort
    ' Store values in tempDict with their keys as a collection
    Dim key As Variant, value As Variant, coll As Collection
    For Each key In dict
    
        value = dict(key)
        
        ' if the value doesn't exist in dict then add
        If dictTemp.Exists(value) = False Then
            ' create collection to hold keys
            ' - needed for duplicate values
            Set coll = New Collection
            dictTemp.Add value, coll
            
            ' Add the value
            arrayList.Add value
            
        End If
        
        ' Add the current key to the collection
        dictTemp(value).Add key
    
    Next key
    
    ' Sort the value
    arrayList.Sort
    
    ' Reverse if descending
    If sortorder = xlDescending Then
        arrayList.Reverse
    End If
    
    dict.RemoveAll
    
    ' Read through the ArrayList and add the values and corresponding
    ' keys from the dictTemp
    Dim item As Variant
    For Each value In arrayList
        Set coll = dictTemp(value)
        For Each item In coll
            dict.Add item, value
        Next item
    Next value
    
    Set arrayList = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByValue = dict
        
Done:
    Exit Function
eh:
    If Err.number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue" _
                , "Cannot sort the dictionary if the value is an object"
    End If
End Function
```

## PrintDictionary

Gibt alle Keys und zugehörigen Values eines `Scripting.Dictionary` in der Debug-Konsole aus.

```vbnet
Public Sub PrintDictionary(dict As Scripting.Dictionary)
    'Gibt alle Schlüssel und Werte des Dictionary in der Debug-Konsole aus.
    'Argumente:
    '    - dict (Scripting.Dictionary): Dictionary, dessen Inhalt ausgegeben wird.
    'Rückgabewert: Keiner (Sub).
    Dim key As Variant
    Dim file As Object

    ' Loop through each key in the dictionary
    For Each key In dict.Keys
        ' Retrieve the IEdmFile5 object associated with each key
        
        ' Print the file name
        Debug.Print "Key: " & key.Name & ", Value: " & dict(key)
    Next key
End Sub
```

## ConvertIntegerToLetter

Diese Funktion wandelt eine positive Ganzzahl in die entsprechende Buchstabenfolge um, wie sie beispielsweise für Spaltenbezeichnungen in Excel verwendet wird (A, B, …, Z, AA, AB, …). Dabei entspricht 1 dem Buchstaben „A“, 2 dem Buchstaben „B“ usw. Nach „Z“ wird mit „AA“ fortgesetzt. Ist die übergebene Zahl kleiner als 1, gibt die Funktion ein Minuszeichen („-“) zurück.

```vbnet
Public Function ConvertIntegerToLetter(ByVal num As Integer) As String
    'Wandelt eine positive Ganzzahl in eine entsprechende Buchstabenfolge um, wobei 1 dem Buchstaben „A“, 2 dem Buchstaben „B“ usw. entspricht. Nach „Z“ wird mit „AA“ fortgesetzt. Dies ist nützlich für die Umwandlung von Zahlen in Buchstabenfolgen nach dem Prinzip des 26er-Systems.
    'Argumente:
    '    - num (Integer): Zu konvertierende Zahl.
    'Rückgabewert: Buchstabenfolge als String, "-" falls num < 1.
    Dim letters As String
    letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
    If num < 1 Then
        ConvertIntegerToLetter = "-"
        Exit Function
    End If
    
    Dim result As String
    result = ""
    
    Do While num > 0
        Dim remainder As Long
        remainder = (num - 1) Mod 26
        result = result & Mid(letters, remainder + 1, 1)
        num = (num - 1) \ 26
    Loop
    
    ConvertIntegerToLetter = result
End Function
```

## IsElementUnique

Überprüft, ob der angegebene Wert (`searchValue`) genau einmal in den Werten eines Dictionaries vorkommt. Gibt `True` zurück, wenn der Wert einzigartig ist, andernfalls `False`.
```vbnet
Public Function IsElementUnique(dict As Object, searchValue As String) As Boolean
    'Prüft, ob searchValue genau einmal als Wert im Dictionary vorkommt.
    'Argumente:
    '    - dict (Object): Dictionary, in dem gesucht wird.
    '    - searchValue (String): Zu prüfender Wert.
    'Rückgabewert: True (Boolean), wenn Wert eindeutig ist, sonst False.

    Dim key As Variant
    Dim valueArray As Variant
    Dim count As Long
    
    count = 0
    
    ' Loop through each key in the dictionary
    For Each key In dict.Keys
        valueArray = dict(key)
        
        ' Check if the first element matches the search value
        If valueArray(0) = searchValue Then
            count = count + 1
        End If
    Next key
    
    ' Return True if the element exists only once, otherwise False
    If count = 1 Then
        IsElementUnique = True
    Else
        IsElementUnique = False
    End If
End Function
```

## IsInCollection
Überprüft, ob ein Element mit dem Namen `sItem` bereits in einer `Collection` vorhanden ist, und gibt `True` zurück, wenn das Element existiert, andernfalls `False`.

```vbnet
Public Function IsInCollection(oCollection As Collection, sItem As String) As Boolean
    'Prüft, ob ein Element mit dem Namen sItem in der Collection vorhanden ist.
    'Argumente:
    '    - oCollection (Collection): Collection, in der gesucht wird.
    '    - sItem (String): Name des gesuchten Elements.
    'Rückgabewert: True (Boolean), wenn Element existiert, sonst False.

    Dim n As Integer
    For n = 1 To oCollection.count
        If oCollection(n).Name = sItem Then
            IsInCollection = True
            Exit Function
        End If
    Next n
    IsInCollection = False
End Function
```

## GetFilePathFromIEdmFile

Gibt den lokalen Speicherpfad einer Datei zurück, basierend auf einem übergebenen `IEdmFile5`-Objekt. Die Funktion ermittelt dazu den zugehörigen Ordner und liefert den vollständigen Pfad der Datei im lokalen Dateisystem.

```vbnet
Public Function GetFilePathFromIEdmFile(file As IEdmFile5) As String
        'Gibt den lokalen Speicherpfad einer Datei anhand eines IEdmFile5-Objekts zurück.
    'Argumente:
    '    - file (IEdmFile5): Dateiobjekt, dessen Pfad ermittelt werden soll.
    'Rückgabewert: Dateipfad als String.

    'Position von erstem, der Datei, übergeordnetem Ordner
    Dim folderPos As IEdmPos5
    Set folderPos = file.GetFirstFolderPosition
    
    'Eltern-Ordner der Datei holen
    Dim parentFolder As IEdmFolder5
    Set parentFolder = file.GetNextFolder(folderPos)

    GetFilePathFromIEdmFile = file.GetLocalPath(parentFolder.ID)
End Function
```

## FormatNumberForReadability

Formatiert große Zahlen so, dass sie mit dem Suffix „k“ (für Tausender) lesbar ausgegeben werden (z.B. 1500 → 1.50k).  
Der Rückgabewert ist ein String mit der formatierten Zahl und dem Suffix „k“.

```vbnet
Public Function FormatNumberForReadability(num As Long) As String
    'Formatiert große Zahlen mit "k"-Suffix für Tausender (z.B. 1500 → 1.50k).
    'Argumente:
    '    - num (Long): Zu formatierende Zahl.
    'Rückgabewert: Formatierte Zahl als String.

    Dim absNum As Double
    absNum = Abs(num)
    
    Select Case absNum
        Case Is >= 1000
            FormatNumberForReadability = Format(num / 1000, "0.00") & "k"
        Case Else
            FormatNumberForReadability = CStr(num)
    End Select
End Function
```

## GetGraphicsTriangles

Diese Funktion ermittelt die Anzahl der Dreiecke (Tessellation) eines SolidWorks-Modells, unabhängig davon, ob es sich um ein Teil oder eine Baugruppe handelt.  
**Hinweis:** Das Logging-Modul muss vor der Ausführung dieser Funktion initialisiert sein, da andernfalls das Logging im Error Handler nicht funktioniert.

```vbnet
Public Function GetGraphicsTriangles(model As ModelDoc2) As Long
    'Ermittelt die Anzahl der Grafikdreiecke (Tessellation) eines SolidWorks-Modells.
    'Argumente:
    '    - model (ModelDoc2): SolidWorks-Modell (Teil oder Baugruppe).
    'Rückgabewert: Anzahl Dreiecke als Long, -1 bei Fehler.

    On Error GoTo ErrorHandler
    
    Dim swPartDoc As SldWorks.PartDoc
    
    If model.GetType = swDocPART Then
        Set swPartDoc = model
        
        GetGraphicsTriangles = swPartDoc.GetTessTriangleCount
    ElseIf model.GetType = swDocASSEMBLY Then
        Dim swAsmDoc As SldWorks.AssemblyDoc
        Set swAsmDoc = model
        
        Dim vComponents As Variant
        vComponents = swAsmDoc.GetComponents(False)

        Dim totalTriangleCount As Long
        totalTriangleCount = 0
        
        Dim component As Variant
        For Each component In vComponents
            
            Dim suppressionState As Integer
            suppressionState = component.GetSuppression2

            'Wenn die Komponente reduziert geladen ist, scheitert das ermitteln der Grafikdreiecke.
            'Daher werden reduzierte Komponenten übersprungen.
            If suppressionState = swComponentFullyResolved Or suppressionState = swComponentResolved Then
                Dim swModelDoc As SldWorks.ModelDoc2
                Set swModelDoc = component.GetModelDoc2
                
                If swModelDoc.GetType = swDocPART Then
                    Set swPartDoc = swModelDoc
                    totalTriangleCount = totalTriangleCount + swPartDoc.GetTessTriangleCount
                End If
            End If
        Next component
        GetGraphicsTriangles = totalTriangleCount
    End If
    
    Exit Function
    
ErrorHandler:
    Logger.logWarn "Ermittlung der Anzahl Grafikdreiecke für " & model.GetTitle & " gescheitert", "GetGraphicsTriangles"
    GetGraphicsTriangles = -1 'Identifikator für fehlgeschlagene Funktion zurückgeben
End Function
```

## FileExistsInPDM
Prüft, ob eine Datei im PDM (Product Data Management) existiert und liefert dabei `True` zurück, wenn mindestens eine Datei mit dem angegebenen Namen gefunden wurde, ansonsten `False`.

```vbnet
Public Function FileExistsInPDM(ByVal article As String) As Boolean
    ' Überprüft, ob eine Datei mit dem Namen `article` im PDM vorhanden ist.
    '
    ' Argumente:
    '   article As String       - Zu suchender Dateiname (mit oder ohne Endung).

    Dim pdm As IEdmVault5
    Dim pdmSearch As IEdmSearch5
    Dim searchResult As IEdmSearchResult5
    Dim fileItem As IEdmFile5

    ' PDM-Verbindung aufbauen
    Set pdm = New EdmVault5
    pdm.LoginAuto "00_Reiden", 0

    ' Suchobjekt konfigurieren
    Set pdmSearch = pdm.CreateSearch
    pdmSearch.FindFiles = True
    pdmSearch.FileName = article

    ' Erstes Suchergebnis abrufen
    Set searchResult = pdmSearch.GetFirstResult

    ' Alle Suchergebnisse durchlaufen
    Do While Not searchResult Is Nothing
        If searchResult.ObjectType = EdmObject_File Then
            Set fileItem = searchResult
            FileExistsInPDM = True
            Exit Function
        End If
        Set searchResult = pdmSearch.GetNextResult
    Loop

    ' Keine Datei gefunden
    FileExistsInPDM = False

    ' Objekte freigeben
    Set pdm = Nothing
    Set pdmSearch = Nothing
    Set searchResult = Nothing
    Set fileItem = Nothing
End Function
```

## ChangeFileExtension
Ändert die Dateiendung eines vollständigen Pfads und gibt den neuen Pfad als `String` zurück.

```vbnet
Public Function ChangeFileExtension(ByVal filePath As String, ByVal newExtension As String) As String
    ' Ändert die Dateiendung eines Pfads.
    '
    ' Argumente:
    '   filePath As String      - Vollständiger Dateipfad (mit oder ohne Endung).
    '   newExtension As String   - Neue Endung (ohne Punkt).

    Dim dotPosition As Long
    Dim basePath As String

    ' Position des letzten Punkts finden
    dotPosition = InStrRev(filePath, ".")

    ' Basis-Pfad ohne Endung extrahieren
    If dotPosition > 0 Then
        basePath = Left(filePath, dotPosition - 1)
    Else
        basePath = filePath
    End If

    ' Neue Endung anhängen
    ChangeFileExtension = basePath & "." & newExtension
End Function
```

## RoundUp
Rundet eine `Double`-Zahl auf die nächsthöhere Ganzzahl auf und gibt diese als `Integer` zurück.

```vbnet
Public Function RoundUp(ByVal Number As Double) As Integer
    ' Rundet eine Zahl auf die nächsthöhere Ganzzahl.
    '
    ' Argumente:
    '   Number As Double    - Zu rundende Zahl.

    If Number = Int(Number) Then
        RoundUp = Number    ' Bereits Ganzzahl
    Else
        RoundUp = Int(Number) + 1
    End If
End Function
```

## MAX
Ermittelt das Maximum von zwei `Double`-Werten und gibt den größeren Wert zurück.

```vbnet
Public Function MAX(val1 As Double, val2 As Double) As Double
    ' Ermittelt das Maximum zweier Werte.
    '
    ' Argumente:
    '   val1 As Double    - Erster Wert.
    '   val2 As Double    - Zweiter Wert.

    If val1 >= val2 Then
        MAX = val1
    Else
        MAX = val2
    End If
End Function
```

## MIN
Ermittelt das Minimum von zwei `Double`-Werten und gibt den kleineren Wert zurück.

```vbnet
Public Function MIN(val1 As Double, val2 As Double) As Double
    ' Ermittelt das Minimum zweier Werte.
    '
    ' Argumente:
    '   val1 As Double    - Erster Wert.
    '   val2 As Double    - Zweiter Wert.

    If val1 <= val2 Then
        MIN = val1
    Else
        MIN = val2
    End If
End Function
```

## FileIsReleased
Überprüft, ob eine Datei freigegeben ist und gibt True zurück, wenn das der Fall ist.
Ansonsten wird False zurückgegeben
**Hinweis:**  Diese Funktion verwendet `GetIEdmFileFromPath`

```vbnet
Public Function FileIsReleased(modelDoc As SldWorks.ModelDoc2) As Boolean
    'Überprüft, ob eine Datei Freigegeben ist
    '
    'Argumente: modelDoc (ModelDoc2) der zu überprüfenden Datei
    'Rückgabewert: True = Freigegeben, ansonsten False (Boolean)
    
    Dim schemaFile As IEdmFile5
    Set schemaFile = GetIEdmFileFromPath(modelDoc.GetPathName)
    
    'Wenn das Schema Freigegeben ist, Fehler ausgeben und Makro beenden
    If schemaFile.currentState.Name = "Freigegeben" Then
        FileIsReleased = True
    Else
        FileIsReleased = False
    End If
End Function
```

## EnsureFileIsCheckedOut
Stellt sicher, dass eine Datei ausgecheckt wird und gibt 0 zurück, wenn das auschecken fehlgeschlagen ist.
**Hinweis:**  Diese Funktion verwendet `GetIEdmFileFromPath` und `GetParentFolderFromPath`

```vbnet
Public Function EnsureFileIsCheckedOut(modelDoc As SldWorks.ModelDoc2) As Integer
    'Diese Funktion stellt sicher, dass eine Datei ausgecheckt wird, falls dies nicht
    'bereits der Fall ist.
    '
    'Argumente: modelDoc (ModelDoc2) der Datei, die ausgecheckt werden soll
    'Rückgabewerte: 0 = auschecken nicht erfolgreich, 1 = erfolgreich (Integer)

    'Geöffnete Zeichnungsdatei vom PDM entkoppeln, sodass sie im nächsten Schritt via PDM ausgecheckt werden kann.
    'Ohne ForceReleaseLocks ist das aus -und einchecken von files via PDM API blockiert
    modelDoc.ForceReleaseLocks
    
    Dim edmFile As IEdmFile5
    Set edmFile = GetIEdmFileFromPath(modelDoc.GetPathName)
    
    'Wenn Datei noch nicht ausgecheckt ist, auschecken
    If Not edmFile.IsLocked Then
        edmFile.LockFile GetParentFolderFromPath(modelDoc.GetPathName).ID, 0
    End If
    
    'Wenn Datei noch immer nicht ausgecheckt ist, 0 für Error zurückgeben
    If edmFile.IsLocked Then
        EnsureFileIsCheckedOut = 0
    Else
        EnsureFileIsCheckedOut = 1 '1, wenn alles gut gelaufen ist
    End If
End Function
```



