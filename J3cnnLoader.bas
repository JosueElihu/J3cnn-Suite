Attribute VB_Name = "J3cnnLoader"
'--------------------------------------------------------------------------------
'    Component  : J3cnnLoader
'    Autor      : J. Elihu
'    Description: Activex Dll loader (without Regsvr32)
'--------------------------------------------------------------------------------

Option Explicit

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal szModule As String) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Any, ByVal oVft As Long, ByVal cc As Integer, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Any, ByRef prgpvarg As Any, ByRef pvargResult As Variant) As Long
Private Declare Function LoadTypeLibEx Lib "oleaut32" (ByVal szFile As Long, ByVal REGKIND As Long, ByRef pptlib As IUnknown) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long ' Always ignore the returned value, it's useless.

Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

Private Const CC_STDCALL          As Long = 4
Private Const REGKIND_NONE        As Long = 2
Private Const TKIND_COCLASS       As Long = 5

Private m_Helper        As Helper   '/*  J3cnn.Helper, J3cnn_c.Helper, J3cnn_mc.Helper    */
Private m_lcnnLib       As Long


' /* J3cnn.dll, J3cnn_c.dll, J3cnn_mc.dll  Constructor Object   */
'=====================================================================================================================
Property Get SQLite() As Constructor
On Error GoTo e_
    If m_Helper Is Nothing Then Call BuildHelper
    Set SQLite = m_Helper.SQLite
e_:
End Property

Private Function BuildHelper() As Boolean
Dim lMod As Long
Dim tmp  As String

        '/*  Check if load library was used     */
        '-----------------------------------------
        lMod = GetModuleHandleA("J3cnn.dll")
        If lMod = 0 Then lMod = GetModuleHandleA("J3cnn_c.dll")
        If lMod = 0 Then lMod = GetModuleHandleA("J3cnn_mc.dll")
        
        If lMod <> 0 Then
            BuildHelper = mvNewHelperInstance(lMod)
            If Not BuildHelper Then MsgBox "Error: SQlite-Helper object not found or invalid object in: " & mvModuleFileName(lMod, True) Else GoTo exit_
        Else
        
            lMod = mvLoadJ3cnnLib(App.Path & "\J3cnn.dll")
            If lMod = 0 Then lMod = mvLoadJ3cnnLib(App.Path & "\J3cnn_c.dll")
            If lMod = 0 Then lMod = mvLoadJ3cnnLib(App.Path & "\J3cnn_mc.dll")
            If lMod Then GoTo exit_
            
            '*/ Search J3cnn dll's in App Sub-folders   */
            '---------------------------------------------
            tmp = Dir(App.Path & "\", vbDirectory)
            Do While tmp > ""
                If PathIsDirectory(App.Path & "\" & tmp) And (tmp <> "." And tmp <> "..") Then
                    lMod = mvLoadJ3cnnLib(App.Path & "\" & tmp & "\J3cnn.dll")
                    If lMod = 0 Then lMod = mvLoadJ3cnnLib(App.Path & "\" & tmp & "\J3cnn_c.dll")
                    If lMod = 0 Then lMod = mvLoadJ3cnnLib(App.Path & "\" & tmp & "\J3cnn_mc.dll")
                    If lMod Then GoTo exit_
                End If
                tmp = Dir()
            Loop
            Debug.Print "J3cnn dll's not found " & vbNewLine & String$(79, "•") & vbNewLine & "J3cnn.dll" & vbNewLine & "J3cnn_c.dll" & vbNewLine & "J3cnn_mc.dll" & vbNewLine & String$(79, "-") & vbNewLine & "Use 'J3cnnLoader.LoadLib' to load it from a custom path."
        End If
exit_:
    If lMod Then m_lcnnLib = lMod
End Function
Private Function mvLoadJ3cnnLib(sLib As String) As Long
    If PathFileExists(sLib) = 0 Then Exit Function
    mvLoadJ3cnnLib = LoadLibraryW(StrPtr(sLib))
    If mvLoadJ3cnnLib Then
        If Not mvNewHelperInstance(mvLoadJ3cnnLib) Then
            FreeLibrary mvLoadJ3cnnLib
            mvLoadJ3cnnLib = 0
        End If
    End If
End Function
Private Function mvNewHelperInstance(lMod As Long) As Boolean
On Error GoTo e_
  Set m_Helper = mvNewObj(lMod, "Helper")
  mvNewHelperInstance = Not (m_Helper Is Nothing)
e_:
End Function


' Activex DLL's Loader
'=====================================================================================================================

Public Function LoadLib(ByVal sDll As String, Optional Search As Boolean) As Long
Dim tmp   As String
Dim sName As String

    LoadLib = GetModuleHandleW(StrPtr(sDll))
    If LoadLib = 0 Then LoadLib = LoadLibraryW(StrPtr(sDll))
    If LoadLib = 0 And Search Then
        sName = Right(sDll, Len(sDll) - InStrRev(sDll, "\"))
        tmp = Dir(App.Path & "\", vbDirectory)
        Do While tmp > ""
             If PathIsDirectory(App.Path & "\" & tmp) And (tmp <> "." And tmp <> "..") Then
                 tmp = App.Path & "\" & tmp & "\"
                 LoadLib = LoadLibraryW(StrPtr(tmp & sName))
                 If LoadLib Then Exit Do
             End If
             tmp = Dir()
        Loop
        If LoadLib = 0 Then Debug.Print "Search library error: " & sName & " not found"
    End If
End Function
Public Function FreeLib(ByVal Lib As String) As Long
Dim lMod    As Long
Dim lpAddr  As Long
Dim Ret     As Long

    If Not IsNumeric(Lib) Then lMod = GetModuleHandleW(StrPtr(Lib)) Else lMod = Lib
    If lMod = 0 Then Exit Function
    If lMod = m_lcnnLib Then Set m_Helper = Nothing
    
    lpAddr = GetProcAddress(lMod, "DllCanUnloadNow")
    If lpAddr <> 0 Then
        If DispCallFunc(0&, lpAddr, CC_STDCALL, vbLong, 0, ByVal 0&, ByVal 0&, Ret) = 0 Then
            If Ret = 0 Then FreeLib = FreeLibrary(lMod)
        End If
    End If
    
End Function
Property Get NewObj(ByVal sClassName As String) As IUnknown
Dim lMod    As Long
Dim sLib    As String
    If InStr(sClassName, ".") Then
        Call mvLib(sClassName, sLib, sClassName)
        lMod = GetModuleHandleW(StrPtr(sLib))
        If lMod = 0 Then lMod = LoadLib(sLib, True)
    Else
        If m_lcnnLib = 0 Then
            m_lcnnLib = GetModuleHandleA("J3cnn.dll")
            If m_lcnnLib = 0 Then m_lcnnLib = GetModuleHandleA("J3cnn_c.dll")
            If m_lcnnLib = 0 Then m_lcnnLib = GetModuleHandleA("J3cnn_mc.dll")
        End If
        lMod = m_lcnnLib
    End If
    Set NewObj = mvNewObj(lMod, sClassName)
End Property
Property Get LibHandle(ByVal sLib As String) As Long
     LibHandle = GetModuleHandleW(StrPtr(sLib))
End Property


Private Property Get mvNewObj(lMod As Long, ByRef sClassName As String) As IUnknown
Dim oTypeLib    As IUnknown
Dim oTypeInfo   As IUnknown
Dim TYPEKIND    As Long
Dim pAttr       As Long
Dim mvRet       As Variant
Dim lCLSID(3)   As Long

    If lMod = 0 Then Exit Property
    If LoadTypeLibEx(StrPtr(mvModuleFileName(lMod)), REGKIND_NONE, oTypeLib) <> 0 Then Exit Property
    Call mvITypeLib_FindName(oTypeLib, sClassName, 0, oTypeInfo, 0, 1)
    If oTypeInfo Is Nothing Then Exit Property
    
    ' GetTypeAttr:
    mvRet = VarPtr(pAttr)
    If DispCallFunc(oTypeInfo, &HC, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(mvRet), 0) <> 0 Then Exit Property
    GetMem4 ByVal pAttr + &H28, TYPEKIND
    If TYPEKIND <> TKIND_COCLASS Then Exit Property
    
    ' GetCLSID:
    CopyMemory lCLSID(0), ByVal pAttr, 16&
    
    ' ReleaseTypeAttr:
    DispCallFunc oTypeInfo, &H4C, CC_STDCALL, vbEmpty, 1, vbLong, VarPtr(CVar(pAttr)), 0
    
    Dim Out                 As IUnknown
    Dim lpAddr              As Long
    Dim CLSID_Factory(3)    As Long
    Dim CLSID_IUnknown(3)   As Long

    lpAddr = GetProcAddress(lMod, "DllGetClassObject")
    If lpAddr = 0 Then Exit Property
    
    'CLSIDFromString StrPtr("{00000001-0000-0000-C000-000000000046}"), CLSID_Factory(0)
    'CLSIDFromString StrPtr("{00000000-0000-0000-C000-000000000046}"), CLSID_IUnknown(0)

    CLSID_Factory(0) = &H1
    CLSID_Factory(2) = &HC0
    CLSID_Factory(3) = &H46000000
    CLSID_IUnknown(2) = &HC0
    CLSID_IUnknown(3) = &H46000000
    
    If mvGetClassObject(lpAddr, lCLSID(0), CLSID_Factory(0), Out) <> 0 Then Exit Property
    If mvIClassFactory_CreateInstance(Out, 0, CLSID_IUnknown(0), mvNewObj) <> 0 Then Set mvNewObj = Nothing
    Set Out = Nothing
    
End Property


'ITypeLib:GetClassObject using a pointer.
'---------------------------------------------------------------------------------------------------------------------------------
Private Function mvGetClassObject(ByVal funcAddr As Long, ByRef CLSID As Long, ByRef IID As Long, ByRef Out As IUnknown) As Long
Dim Params(2)   As Variant
Dim Types(2)    As Integer
Dim List(2)     As Long
Dim mvRet     As Variant
Dim i           As Long

    Params(0) = VarPtr(CLSID)
    Params(1) = VarPtr(IID)
    Params(2) = VarPtr(Out)
    For i = 0 To UBound(Params)
        List(i) = VarPtr(Params(i)):   Types(i) = VarType(Params(i))
    Next
    mvGetClassObject = DispCallFunc(0&, funcAddr, CC_STDCALL, vbLong, 3, Types(0), List(0), mvRet)
    If mvGetClassObject = 0 Then mvGetClassObject = mvRet
    
End Function

'ITypeLib:CreateInstance
'---------------------------------------------------------------------------------------------------------------------------------
Private Function mvIClassFactory_CreateInstance(ByVal obj As IUnknown, ByVal pUnkOuter As Long, ByRef riid As Long, ByRef Out As IUnknown) As Long
Dim Params(2)   As Variant
Dim Types(2)    As Integer
Dim List(2)     As Long
Dim mvRet     As Variant
Dim i      As Long

    Params(0) = pUnkOuter
    Params(1) = VarPtr(riid)
    Params(2) = VarPtr(Out)
    For i = 0 To UBound(Params)
        List(i) = VarPtr(Params(i)):   Types(i) = VarType(Params(i))
    Next
    mvIClassFactory_CreateInstance = DispCallFunc(obj, &HC, CC_STDCALL, vbLong, 3, Types(0), List(0), mvRet)
    If mvIClassFactory_CreateInstance = 0 Then mvIClassFactory_CreateInstance = mvRet

End Function

'ITypeLib:FindName
'---------------------------------------------------------------------------------------------------------------------------------
Private Function mvITypeLib_FindName(ByVal obj As IUnknown, ByRef szNameBuf As String, ByVal lHashVal As Long, ByRef ppTInfo As IUnknown, ByRef rgMemId As Long, ByRef pcFound As Integer) As Long
Dim Params(4) As Variant
Dim Types(4)  As Integer
Dim List(4)   As Long
Dim mvRet     As Variant
Dim i         As Long

    Params(0) = StrPtr(szNameBuf)
    Params(1) = lHashVal
    Params(2) = VarPtr(ppTInfo)
    Params(3) = VarPtr(rgMemId)
    Params(4) = VarPtr(pcFound)
    
    For i = 0 To UBound(Params)
        List(i) = VarPtr(Params(i)):  Types(i) = VarType(Params(i))
    Next
    mvITypeLib_FindName = DispCallFunc(obj, &H2C, CC_STDCALL, vbLong, 5, Types(0), List(0), mvRet)
    If mvITypeLib_FindName = 0 Then mvITypeLib_FindName = mvRet
    
End Function

Private Function mvModuleFileName(lMod As Long, Optional OnlyTitle As Boolean, Optional RmvExt As Boolean) As String
On Error GoTo e_
Dim lLen As Long
    mvModuleFileName = String(260, 0)
    lLen = GetModuleFileName(lMod, mvModuleFileName, 260)
    If lLen Then
        mvModuleFileName = Left(mvModuleFileName, lLen)
        If OnlyTitle Then mvModuleFileName = Right(mvModuleFileName, Len(mvModuleFileName) - InStrRev(mvModuleFileName, "\"))
        If RmvExt Then mvModuleFileName = Left$(mvModuleFileName, InStrRev(mvModuleFileName, ".") - 1)
    Else
        mvModuleFileName = vbNullString
    End If
e_:
End Function
Private Function mvLib(sArgs As String, sLib As String, sClass As String) As Long
    mvLib = InStrRev(sArgs, ".")
    If mvLib Then sLib = Left$(sArgs, mvLib - 1): sClass = Right$(sArgs, Len(sArgs) - mvLib) Else sLib = sArgs
End Function
