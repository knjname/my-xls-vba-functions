Attribute VB_Name = "knjname_MyXlsFunc_Core"
Option Explicit

' Depends on "Microsoft Scripting Runtime"
' Depends on "Micsoroft VBScript Regular Expressions 5.5"

Private xlsVbaFuncFSO As New FileSystemObject

Function increment&(ByRef value&, Optional ByVal added = 1)
    increment = value
    value = value + added
End Function

Function prefixIncrement&(ByRef value&, Optional ByVal added = 1)
    value = value + added
    prefixIncrement = value
End Function

Function unifyToCrLf$(ByVal content$, Optional ByVal lineFeedChar = vbCrLf)
    unifyToCrLf = regexpReplace(content, "\r?\n|\r\n?|\n", lineFeedChar)
End Function

Function dict(ParamArray keyAndValues() As Variant) As Dictionary
    Dim d As Dictionary
    Set dict = addDictEntries(d, keyAndValues)
End Function

Function addDictEntries(ByRef d As Dictionary, keyAndValues() As Variant) As Dictionary
    If d Is Nothing Then
        Set d = New Dictionary
    End If
    
    Dim i&
    For i = LBound(keyAndValues) To UBound(keyAndValues)
        If IsObject(keyAndValues(i + 1)) Then
            Set d(keyAndValues(i)) = keyAndValues(i + 1)
        Else
            d(keyAndValues(i)) = keyAndValues(i + 1)
        End If
    Next
End Function

Function seekFilesRecursively( _
    ByVal baseDir$, _
    ByVal testerRegexp As regexp, _
    Optional ByVal minDepth& = -1, _
    Optional ByVal maxDepth& = -1, _
    Optional ByVal collectFile As Boolean = True, _
    Optional ByVal collectFolder As Boolean = False, _
    Optional ByVal terminateWhenCollectedFolderFound As Boolean = True, _
    Optional ByVal ignoreDotFolderOnRecursion As Boolean = True, _
    Optional ByVal ignoreTemporaryOfficeFile As Boolean = True, _
    Optional ByVal testWithFullPath As Boolean = False, _
    Optional ByVal skipStartFrom& = 0, _
    Optional ByVal limitResultNumber& = -1, _
    Optional ByVal collectAsString As Boolean = False, _
    Optional ByVal noErrorEvenBaseDirIsMissing As Boolean = True) As Collection
    
    Dim result As New Collection
    Set seekFilesRecursively = result
    
    If minDepth > 0 And maxDepth > 0 And minDepth > maxDepth Then
        Err.Raise 7777, "seekFileRecursively()", "Invalid minDepth(" & minDepth & ") - maxDepth(" & maxDepth & ") arguments."
    End If
    
    If xlsVbaFuncFSO.FolderExists(baseDir) Then
        
        seekFilesRecursively_rec _
            0, _
            result, _
            xlsVbaFuncFSO.GetFolder(baseDir), _
            testerRegexp, _
            minDepth, _
            maxDepth, _
            collectFile, _
            collectFolder, _
            terminateWhenCollectedFolderFound, _
            ignoreDotFolderOnRecursion, _
            ignoreTemporaryOfficeFile, _
            testWithFullPath, _
            skipStartFrom, _
            limitResultNumber, _
            collectAsString
        
    Else
        If Not noErrorEvenBaseDirIsMissing Then
            ' TODO assign proper error number
            Err.Raise 7777, "seekFileRecursively()", "No basedir is found at " & baseDir & "."
        End If
    End If
 
End Function

' Workhorse function
Private Sub seekFilesRecursively_rec( _
    ByVal depth&, _
    ByVal result As Collection, _
    ByVal baseDir As Folder, _
    ByVal testerRegexp As regexp, _
    ByVal minDepth&, _
    ByVal maxDepth&, _
    ByVal collectFile As Boolean, _
    ByVal collectFolder As Boolean, _
    ByVal terminateWhenCollectedFolderFound As Boolean, _
    ByVal ignoreDotFolderOnRecursion As Boolean, _
    ByVal ignoreTemporaryOfficeFile As Boolean, _
    ByVal testWithFullPath As Boolean, _
    ByRef skipStartFrom&, _
    ByVal limitResultNumber&, _
    ByVal collectAsString As Boolean)
    
    Dim shouldBeCollected As Boolean
    shouldBeCollected = (minDepth < 0 Or minDepth <= depth)
    
    Dim f As File
    Dim d As Folder
     
    If shouldBeCollected And collectFile Then
        For Each f In baseDir.Files
            If seekFilesRecursively_testPath(testerRegexp, f, ignoreTemporaryOfficeFile, testWithFullPath) Then
                If ignoreTemporaryOfficeFile Imp Not isOfficeTemporaryFile(f.Name) Then
                    seekFilesRecursively_addResult result, f, collectAsString, skipStartFrom
                    
                    If seekFilesRecursively_exceedsLimitCount(result, limitResultNumber) Then Exit Sub
                End If
            End If
        Next
    End If
    For Each d In baseDir.SubFolders
        Dim foundDir As Boolean
        foundDir = False
        
        If shouldBeCollected And collectFolder Then
            If seekFilesRecursively_testPath(testerRegexp, d, ignoreTemporaryOfficeFile, testWithFullPath) Then
                seekFilesRecursively_addResult result, d, collectAsString, skipStartFrom
                If seekFilesRecursively_exceedsLimitCount(result, limitResultNumber) Then Exit Sub
                foundDir = True
            End If
        End If
        
        If terminateWhenCollectedFolderFound Imp Not foundDir Then
            If maxDepth < 0 Or depth < maxDepth Then
                If ignoreDotFolderOnRecursion Imp Left(d.Name, 1) <> "." Then
                    seekFilesRecursively_rec _
                        depth + 1, _
                        result, _
                        d, _
                        testerRegexp, _
                        minDepth, _
                        maxDepth, _
                        collectFile, _
                        collectFolder, _
                        terminateWhenCollectedFolderFound, _
                        ignoreDotFolderOnRecursion, _
                        ignoreTemporaryOfficeFile, _
                        testWithFullPath, _
                        skipStartFrom, _
                        limitResultNumber, _
                        collectAsString
                    If seekFilesRecursively_exceedsLimitCount(result, limitResultNumber) Then Exit Sub
                End If
            End If
        End If
    Next
     
 
End Sub

Private Function seekFilesRecursively_exceedsLimitCount(ByVal result As Collection, ByVal limitNumber&)
    If limitNumber& > 0 Then seekFilesRecursively_exceedsLimitCount = result.Count = limitNumber&
End Function

Private Function isOfficeTemporaryFile(ByVal fileName$) As Boolean
    If Len(isOfficeTemporaryFile) > 3 Then
        If Left(isOfficeTemporaryFile, 2) = "$~" Then
            isOfficeTemporaryFile = True
        End If
    End If
End Function

Private Function seekFilesRecursively_testPath(ByVal r As regexp, ByVal fileOrDir As Object, ByVal ignoreTemporaryOfficeFile As Boolean, ByVal testWithFullPath As Boolean) As Boolean
    If testWithFullPath Then
        seekFilesRecursively_testPath = r.Test(fileOrDir)
    Else
        seekFilesRecursively_testPath = r.Test(fileOrDir.Name)
    End If
End Function

Private Sub seekFilesRecursively_addResult(ByVal resultCol As Collection, ByVal result As Object, ByVal collectAsString As Boolean, ByRef skipStartFrom&)
    If skipStartFrom <= 0 Then
        If collectAsString Then
            resultCol.Add CStr(result)
        Else
            resultCol.Add result
        End If
    End If
    
    skipStartFrom = skipStartFrom - 1
End Sub

Function newRegExp(ByVal pattern$, _
    Optional ByVal ignoreCase As Boolean = False, _
    Optional ByVal multiLine As Boolean = False, _
    Optional ByVal matchesGlobally As Boolean = True) _
    As regexp

    Set newRegExp = New regexp
    With newRegExp
        .pattern = pattern
        .ignoreCase = ignoreCase
        .multiLine = multiLine
        .Global = matchesGlobally
    End With

End Function

Function regexpTest(ByVal tested$, _
    ByVal pattern$, _
    Optional ByVal ignoreCase As Boolean = False, _
    Optional ByVal multiLine As Boolean = False) As Boolean

    With newRegExp(pattern, ignoreCase:=ignoreCase, multiLine:=multiLine)
        regexpTest = .Test(tested)
    End With

End Function

Function regexpReplace$(ByVal replaced$, _
    ByVal pattern$, _
    ByVal replacement$, _
    Optional ByVal ignoreCase As Boolean = False, _
    Optional ByVal multiLine As Boolean = False, _
    Optional ByVal matchesGlobally As Boolean = True)

    With newRegExp(pattern, ignoreCase:=ignoreCase, multiLine:=multiLine, matchesGlobally:=matchesGlobally)
        regexpReplace = .Replace(replaced, replacement)
    End With

End Function

