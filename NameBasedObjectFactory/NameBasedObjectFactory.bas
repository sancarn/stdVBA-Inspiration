Attribute VB_Name = "NameBasedObjectFactory"
Option Explicit
' ******************************************************************************
' *     Fire-Lines © 2011                                                      *
' *     Модуль:                                                                *
' *         NameBasedObjectFactory                                             *
' *     Описание:                                                              *
' *         Этот модуль обеспечивает возможность создавать экземпляры классов  *
' *         по строковому имени соответствующего класса.                       *
' *     Автор:                                                                 *
' *         Владислав Петровский (firehacker)                                  *
' *     История изменений:                                                     *
' *         *   2011-05-21  firehacker  Файл создан.                           *
' *     Примечания:                                                            *
' *         Переименование модуля вызовет его неработоспосбность!              *
' *     Текущая версия: 1.0.0                                                  *
' *                                                                            *
' ******************************************************************************

#Const SUPPORT_CRYPTED_EXECUTABLE = False
#Const DENY_USERCONTROLS = True

Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Foo1 As Long, ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long

Private Declare Function lstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal s1 As String, ByVal s2 As Long) As Long
Private Declare Function NbofRtlSimpleNew Lib "msvbvm60" Alias "__vbaNew" (lpObjectInformation As Any) As IUnknown
Private Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (ary() As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal lpAddress As Long, dst As Any)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal lpAddress As Long, ByVal nv As Long)
        
Private Type EXEPROJECTINFO
    Signature                       As Long
    RuntimeVersion                  As Integer
    BaseLanguageDll(0 To 13)        As Byte
    ExtLanguageDll(0 To 13)         As Byte
    RuntimeRevision                 As Integer
    BaseLangiageDllLCID             As Long
    ExtLanguageDllLCID              As Long
    lpSubMain                       As Long
    lpProjectData                   As Long
    ' < Дальше есть другие поля, но я их не объявляю, потому что нам они не нужны >
End Type

Private Type ProjectData
    Version                         As Long
    lpModuleDescriptorsTableHeader  As Long
    ' < Дальше есть другие поля, но я их не объявляю, потому что нам они не нужны >
End Type

Private Type MODDESCRTBL_HEADER
    Reserved0                       As Long
    lpProjectObject                 As Long
    lpProjectExtInfo                As Long
    Reserved1                       As Long
    Reserved2                       As Long
    lpProjectData                   As Long
    guid(0 To 15)                   As Byte
    Reserved3                       As Integer
    TotalModuleCount                As Integer
    CompiledModuleCount             As Integer
    UsedModuleCount                 As Integer
    lpFirstDescriptor               As Long
    ' < Дальше есть другие поля, но я их не объявляю, потому что нам они не нужны >
End Type

Private Enum MODFLAGS
    mfBasic = 1
    mfNonStatic = 2
    mfUserControl = &H42000
End Enum

Private Type MODDESCRTBL_ENTRY
    lpObjectInfo                    As Long
    FullBits                        As Long
    Placeholder0(0 To 15)           As Byte
    lpszName                        As Long
    MethodsCount                    As Long
    lpMethodNamesArray              As Long
    Placeholder1                    As Long
    ModuleType                      As MODFLAGS
    Placeholder2                    As Long
End Type

Public Function CreateObjectPrivate(ByVal Class As String) As IUnknown
    Dim IDE_MODE As Boolean: Debug.Assert LetTrue(IDE_MODE)
    '
    ' При работе в скомпилированном виде и под IDE будут использоваться разные механизмы.
    '
    
    If IDE_MODE Then
        Set CreateObjectPrivate = NbofDbgCreateInstance(Class)
    Else
        Set CreateObjectPrivate = NbofRtCreateInstance(Class)
    End If
End Function

Private Function NbofRtCreateInstance(ByVal Class As String) As IUnknown
    Dim lpObjectInformation As Long
    
    '
    ' Получаем адрес блока информации о классе. Если такой класс не найден, будет
    ' сгенерирована ошибка. А затем создаём экземпляр указанного класса.
    '
    
    If Not NbofGetOiOfClass(Class, lpObjectInformation) Then
        Err.Raise 8, , "Specified class '" + Class + "' does not defined."
        Exit Function
    End If
    
    Set NbofRtCreateInstance = NbofRtlSimpleNew(ByVal lpObjectInformation)
End Function

Private Function NbofGetOiOfClass(ByVal Class As String, ByRef lpOi As Long) As Boolean
    Static Modules()        As NameBasedObjectFactory.MODDESCRTBL_ENTRY
    Static bModulesSet      As Boolean
    Dim i                   As Long
    
    #If DENY_USERCONTROLS Then
        Const mfBadFlags As Long = mfUserControl
    #Else
        Const mfBadFlags As Long = 0
    #End If
    
    If Not bModulesSet Then
        ReDim Modules(0)
        If NbofLoadDescriptorsTable(Modules) Then bModulesSet = True Else Exit Function
    End If
    
    '
    ' Ищем дескриптор, соответствующий указанному классу.
    '
    
    For i = LBound(Modules) To UBound(Modules)
        With Modules(i)
        If lstrcmpi(Class, .lpszName) = 0 And _
            CBool(.ModuleType And mfNonStatic) And Not _
            CBool(.ModuleType And mfBadFlags) Then
                lpOi = .lpObjectInfo
                NbofGetOiOfClass = True: Exit Function
            End If
        End With
    Next i
End Function

Private Function NbofLoadDescriptorsTable(dt() As MODDESCRTBL_ENTRY) As Boolean
    Dim lpEPI               As Long
    Dim EPI(0)              As NameBasedObjectFactory.EXEPROJECTINFO
    Dim ProjectData(0)      As NameBasedObjectFactory.ProjectData
    Dim ModDescrTblHdr(0)   As NameBasedObjectFactory.MODDESCRTBL_HEADER
    
    '
    ' WARNING: Эта процедура вызывается ТОЛЬКО ОДИН раз за всё время работы проекта.
    ' Получаем адрес EPI.
    '

    If Not NbofFindEpiSimple(lpEPI) Then
        #If SUPPORT_CRYPTED_EXECUTABLE Then
            If Not NbofFindEpiFull(lpEPI) Then
                Err.Raise 17, , "Failed to locate EXEPROJECTINFO structure in process address space."
                Exit Function
            End If
        #Else
            Err.Raise 17, , "Failed to locate EXEPROJECTINFO structure in process module image."
            Exit Function
        #End If
    End If
    
    '
    ' Из EPI находим местоположение PROJECTDATA, из PROJECTDATA получаем местоположение
    ' заголовка таблицы дескрипторов, из заголовка получаем количество дескрипторов и
    ' адрес начала последовательности.
    '
    
    SaMap AryPtr(EPI), lpEPI
    SaMap AryPtr(ProjectData), EPI(0).lpProjectData: SaUnmap AryPtr(EPI)
    SaMap AryPtr(ModDescrTblHdr), ProjectData(0).lpModuleDescriptorsTableHeader: SaUnmap AryPtr(ProjectData)
    SaMap AryPtr(dt), ModDescrTblHdr(0).lpFirstDescriptor, ModDescrTblHdr(0).TotalModuleCount: SaUnmap AryPtr(ModDescrTblHdr)
   
    NbofLoadDescriptorsTable = True
End Function

Private Function NbofFindEpiSimple(ByRef lpEPI As Long) As Boolean
    Dim DWords()            As Long: ReDim DWords(0)
    Dim PotentionalEPI(0)   As NameBasedObjectFactory.EXEPROJECTINFO
    Dim PotentionalPD(0)    As NameBasedObjectFactory.ProjectData
    Dim i                   As Long
    
    Const EPI_Signature     As Long = &H21354256 ' "VB5!"
    Const PD_Version        As Long = &H1F4
    
    '
    ' Пытаемся получить указатель на структуру EXEPROJECTINFO. Её адрес нигде не хранится,
    ' поэтому единственный способ найти структуру — найти её по сигнатуре, которая у неё,
    ' к счастью, имеется (наверное специально для этой цели ;) ).
    '
    
    '
    ' Текущая реализация поиска просто отвратительна: она ищет сигнатуру с самого начала
    ' образа, включая и те места, в которых её быть заведомо не может. И поиск выходит
    ' за границы образа, если найти сигнатуру внутри виртуального образа не удалось.
    ' Это скорее всего приведёт к AV-исключению. Но она (реализация) компактна :)
    '
    
    SaMap AryPtr(DWords), App.hInstance
    Do
        If DWords(i) = EPI_Signature Then
            SaMap AryPtr(PotentionalEPI), VarPtr(DWords(i))
            SaMap AryPtr(PotentionalPD), PotentionalEPI(0).lpProjectData
            
            If PotentionalPD(0).Version = PD_Version Then
                lpEPI = VarPtr(DWords(i))
                NbofFindEpiSimple = True
            End If
            
            SaUnmap AryPtr(PotentionalPD)
            SaUnmap AryPtr(PotentionalEPI)
            
            If NbofFindEpiSimple Then Exit Do
        End If
        
        i = i + 1
    Loop
    SaUnmap AryPtr(DWords)
End Function

Private Function NbofDbgCreateInstance(ByVal Class As String) As IUnknown
    EbExecuteLine StrPtr("NameBasedObjectFactory.NbofOneCellQueue new " + Class), 0, 0, 0
    Set NbofDbgCreateInstance = NbofOneCellQueue(Nothing)
    If NbofDbgCreateInstance Is Nothing Then
        Err.Raise 8, , "Specified class '" + Class + "' does not defined."
        Exit Function
    End If
End Function

Private Function NbofOneCellQueue(ByVal refIn As IUnknown) As IUnknown
    Static cell As IUnknown
    Set NbofOneCellQueue = cell
    Set cell = refIn
End Function

Private Sub SaMap(ByVal ppSA As Long, ByVal pMemory As Long, Optional ByVal NewSize As Long = -1)
    Dim pSA As Long: GetMem4 ppSA, pSA:
    PutMem4 pSA + 12, ByVal pMemory: PutMem4 pSA + 16, ByVal NewSize
End Sub

Private Sub SaUnmap(ByVal ppSA As Long)
    Dim pSA As Long: GetMem4 ppSA, pSA
    PutMem4 pSA + 12, ByVal 0: PutMem4 pSA + 16, ByVal 0
End Sub

Private Function LetTrue(b As Boolean) As Boolean: b = True: LetTrue = True: End Function

