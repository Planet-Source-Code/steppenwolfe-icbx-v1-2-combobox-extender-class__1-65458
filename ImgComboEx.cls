VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCBExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  ICBExtender  - ImageComboBox Extender Class                                        *
'*                                                                                     *
'*  Created:     May 24, 2006                                                          *
'*  Purpose:     ComboBox Extender Class                                               *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.1.0                                                                 *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*  Credits to:  M. Novaro for the Imagebox class extender (txtCodeId=34778),          *
'*               and Paul Caton, for the great self subclass routines                  *
'*                                                                                     *
'***************************************************************************************

' ~*** Notes ***~

'/~ Couldn't find a combobox with the nice features of the M$ version(Darn that M$!)
'/~ so here you go..
'/~ Added a directory in list browse, through root level folders.
'/~ Added a browse button to the list that invokes a standard dialog.
'/~ Just add a picturebox, and an imagelist to the form, and wire it up.
'/~ You could build on this.. add font, color picker, etc.. and if
'/~ you want to hack the dropdown list some more, you could do things like
'/~ expand the window dynamically, and make it so that all folders could be browsed,
'/~ doing away with the need for directory/drive list couplings.
'/~ Also, styles could be manipulated, scrollbar and button colors/shape..
'/~ It's up to you..

'/~ May 24, 2006 - Ver. 1.1
'/~ Applied a couple of fixes/improvements, and some basic error handling.
'/~ John Underhill (Steppenwolfe) steppenwolfe_2000@yahoo.com

'/~ Properties ~/
'/~ p_Picture                   - set the picturebox
'/~ p_ImageList                 - set the imagelist
'/~ p_AutoComplete              - autocomplete a phrase
'/~ p_AutoSelect                - select the entry
'/~ p_AutoSize                  - size list to largest entry
'/~ p_DirectoryList             - enable directory view
'/~ p_ItemSelected              -
'/~ p_KeyedDescent              - lower list with keystroke
'/~ p_ListWidth                 - set list start width
'/~ p_RestrictItems             - allow only list items

'/~ Routines ~/
'/~ SetReference                - attatch the combobox
'/~ ListDrives                  - select display type

Private Const CB_GETEDITSEL                 As Long = &H140
Private Const CB_LIMITTEXT                  As Long = &H141
Private Const CB_SETEDITSEL                 As Long = &H142
Private Const CB_ADDSTRING                  As Long = &H143
Private Const CB_DELETESTRING               As Long = &H144
Private Const CB_DIR                        As Long = &H145
Private Const CB_GETCOUNT                   As Long = &H146
Private Const CB_GETCURSEL                  As Long = &H147
Private Const CB_GETLBTEXT                  As Long = &H148
Private Const CB_GETLBTEXTLEN               As Long = &H149
Private Const CB_INSERTSTRING               As Long = &H14A
Private Const CB_RESETCONTENT               As Long = &H14B
Private Const CB_FINDSTRING                 As Long = &H14C
Private Const CB_SELECTSTRING               As Long = &H14D
Private Const CB_SETCURSEL                  As Long = &H14E
Private Const CB_SHOWDROPDOWN               As Long = &H14F
Private Const CB_GETITEMDATA                As Long = &H150
Private Const CB_SETITEMDATA                As Long = &H151
Private Const CB_GETDROPPEDCONTROLRECT      As Long = &H152
Private Const CB_SETITEMHEIGHT              As Long = &H153
Private Const CB_GETITEMHEIGHT              As Long = &H154
Private Const CB_SETEXTENDEDUI              As Long = &H155
Private Const CB_GETEXTENDEDUI              As Long = &H156
Private Const CB_GETDROPPEDSTATE            As Long = &H157
Private Const CB_FINDSTRINGEXACT            As Long = &H158
Private Const CB_SETLOCALE                  As Long = &H159
Private Const CB_GETTOPINDEX                As Long = &H15B
Private Const CB_SETTOPINDEX                As Long = &H15C
Private Const CB_GETHORIZONTALEXTENT        As Long = &H15D
Private Const CB_SETHORIZONTALEXTENT        As Long = &H15E
Private Const CB_GETDROPPEDWIDTH            As Long = &H15F
Private Const CB_SETDROPPEDWIDTH            As Long = &H160
Private Const CB_INITSTORAGE                As Long = &H161
Private Const BIF_RETURNONLYFSDIRS          As Integer = 1
Private Const BIF_DONTGOBELOWDOMAIN         As Integer = 2
Private Const MAX_PATH                      As Integer = 260
Private Const BASIC_SHGFI_FLAGS             As Double = _
    &H4 Or &H200 Or &H400 Or &H2000 Or &H4000

Private Const ALL_MESSAGES                  As Long = -1
Private Const MSG_ENTRIES                   As Long = 32
Private Const WNDPROC_OFF                   As Long = &H38
Private Const GWL_WNDPROC                   As Long = -4
Private Const IDX_SHUTDOWN                  As Long = 1
Private Const IDX_HWND                      As Long = 2
Private Const IDX_WNDPROC                   As Long = 9
Private Const IDX_BTABLE                    As Long = 11
Private Const IDX_ATABLE                    As Long = 12
Private Const IDX_PARM_USER                 As Long = 13

Private Type FT
    lLD                                     As Long
    lHD                                     As Long
End Type

Private Type WFD
    lFA                                     As Long
    fCT                                     As FT
    fLA                                     As FT
    fLWT                                    As FT
    lFSH                                    As Long
    lFSL                                    As Long
    lR0                                     As Long
    lR1                                     As Long
    cFN                                     As String * 260
    cAL                                     As String * 14
End Type

Private Type BROWSEINFO
    hwndOwner                               As Long
    pidlRoot                                As Long
    pszDisplayName                          As Long
    lpszTitle                               As Long
    ulFlags                                 As Long
    lpfnCallback                            As Long
    lParam                                  As Long
    iImage                                  As Long
End Type

Private Type SHFILEINFO
    hIcon                                   As Long
    iIcon                                   As Long
    dwAttributes                            As Long
    szDisplayName                           As String * 260
    szTypeName                              As String * 80
End Type

Private Type POINTAPI
    X                                       As Long
    Y                                       As Long
End Type

Private Type RECT
    left                                    As Long
    tOp                                     As Long
    Right                                   As Long
    Bottom                                  As Long
End Type

Public Enum LIST_FORMAT
    Drives_Only = 0
    Special_Folders = 1
End Enum

Private Type SHITEMID
    cb                                      As Long
    abID                                    As Byte
End Type

Private Type ITEMIDLIST
    mkid                                    As SHITEMID
End Type

Private Enum eMsgWhen
    MSG_BEFORE = 1
    MSG_AFTER = 2
    MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER
End Enum

Public Enum eDirectories
    CSIDL_DESKTOPDIRECTORY = 0
    CSIDL_START_PROGRAMS = 2
    CSIDL_MYDOCUMENTS = 5
    CSIDL_FAVORITES = 6
    CSIDL_STARTUP = 7
    CSIDL_RECENT = 8
    CSIDL_SENDTO = 9
    CSIDL_START_MENU = 11
    CSIDL_MYMUSIC = 13
    CSIDL_MYVIDEO = 14
    CSIDL_DESKTOP = 16
    CSIDL_NETHOOD = 19
    CSIDL_FONTS = 20
    CSIDL_TEMPLATES = 21
    CSIDL_COMMON_STARTMENU = 22
    CSIDL_COMMON_PROGRAMS = 23
    CSIDL_COMMON_STARTUP = 24
    CSIDL_COMMON_DESKTOP = 25
    CSIDL_APPDATA = 26
    CSIDL_PRINTHOOD = 27
    CSIDL_SETTINGS_APPDATA = 28
    CSIDL_COMMON_FAVORITES = 31
    CSIDL_INTERNET_CACHE = 32
    CSIDL_COOKIES = 33
    CSIDL_HISTORY = 34
    CSIDL_COMMON_APPDATA = 35
    CSIDL_WINDOWS = 36
    CSIDL_SYSTEM = 37
    CSIDL_PROGRAM_FILES = 38
    CSIDL_MYPICTURES = 39
    CSIDL_PROFILE = 40
    CSIDL_COMMON_SYSTEM = 42
    CSIDL_COMMON_FILES = 43
    CSIDL_COMMON_TEMPLATES = 45
    CSIDL_COMMON_DOCUMENTS = 46
    CSIDL_COMMON_MUSIC = 53
    CSIDL_COMMON_PICTURES = 54
    CSIDL_COMMON_VIDEO = 55
    CSIDL_RESOURCES = 56
    CSIDL_CD_BURN_AREA = 56
End Enum

Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessageAsLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                              ByVal wMsg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Function GetVolumeInformation Lib "Kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                                                                                ByVal lpVolumeNameBuffer As String, _
                                                                                                ByVal nVolumeNameSize As Long, _
                                                                                                lpVolumeSerialNumber As Long, _
                                                                                                lpMaximumComponentLength As Long, _
                                                                                                lpFileSystemFlags As Long, _
                                                                                                ByVal lpFileSystemNameBuffer As String, _
                                                                                                ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                                ByVal lpBuffer As String) As Long

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, _
                                                            ByVal i As Long, _
                                                            ByVal hDCDest As Long, _
                                                            ByVal X As Long, _
                                                            ByVal Y As Long, _
                                                            ByVal Flags As Long) As Long

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, _
                                                                       ByVal csidl As Long, _
                                                                       ByRef ppidl As ITEMIDLIST) As Long

Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lFN As String, _
                                                                              lFD As WFD) As Long

Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal lHD As Long, _
                                                                            lFD As WFD) As Long

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hwnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, _
                                                                lpdwProcessId As Long) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, _
                                                      ByVal dwSize As Long, _
                                                      ByVal flAllocationType As Long, _
                                                      ByVal flProtect As Long) As Long

Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, _
                                                     ByVal dwSize As Long, _
                                                     ByVal dwFreeType As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, _
                                                  ByVal Source As Long, _
                                                  ByVal Length As Long)

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lLongPath As String, _
                                                                                    ByVal lShortPath As String, _
                                                                                    ByVal lBuffer As Long) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal lHD As Long) As Long

Public Event eFilePath(sPath As String)

Private WithEvents m_oComboBox              As ImageCombo
Attribute m_oComboBox.VB_VarHelpID = -1
Private m_bAutoSize                         As Boolean
Private m_bAutoComplete                     As Boolean
Private blnAuto                             As Boolean
Private m_bRestrictItems                    As Boolean
Private m_bAutoSelect                       As Boolean
Private m_bDirectoryList                    As Boolean
Private m_bExtended                         As Boolean
Private z_ScMem                             As Long
Private z_Sc(64)                            As Long
Private m_lMenuHandle                       As Long
Private m_ExtDrive                          As String
Private m_cSIcon                            As Collection
Private z_Funk                              As Collection
Private m_oImageList                        As Object
Private m_oPicture                          As Object
Private tShInfo                             As SHFILEINFO
Private m_cMem                              As Collection
Private m_eListMem                          As LIST_FORMAT


Public Property Get p_AutoComplete() As Boolean
    p_AutoComplete = m_bAutoComplete
End Property

Public Property Let p_AutoComplete(ByVal PropVal As Boolean)
    m_bAutoComplete = PropVal
End Property

Public Property Get p_AutoSelect() As Boolean
    p_AutoSelect = m_bAutoSelect
End Property

Public Property Let p_AutoSelect(ByVal PropVal As Boolean)
    m_bAutoSelect = PropVal
End Property

Public Property Get p_AutoSize() As Boolean
    p_AutoSize = m_bAutoSize
End Property

Public Property Let p_AutoSize(ByVal PropVal As Boolean)
    m_bAutoSize = PropVal
End Property

Public Property Get p_DirectoryList() As Boolean
    p_DirectoryList = m_bDirectoryList
End Property

Public Property Let p_DirectoryList(ByVal PropVal As Boolean)
    m_bDirectoryList = PropVal
End Property

Public Property Get p_ImageList() As Object
    Set p_ImageList = m_oImageList
End Property

Public Property Set p_ImageList(ByVal PropVal As Object)
    Set m_oImageList = PropVal
    Set m_oComboBox.ImageList = m_oImageList
End Property

Public Property Get p_KeyedDescent() As Boolean
    p_KeyedDescent = (SendMessageLong(m_oComboBox.hwnd, CB_GETEXTENDEDUI, 0, 0) <> 0)
End Property

Public Property Let p_KeyedDescent(ByVal PropVal As Boolean)
    SendMessageLong m_oComboBox.hwnd, CB_SETEXTENDEDUI, Abs(PropVal), 0
End Property

Private Property Get p_ListWidth() As Long
    p_ListWidth = SendMessage(m_oComboBox.hwnd, CB_GETDROPPEDWIDTH, 0, ByVal 0)
End Property

Private Property Let p_ListWidth(PropVal As Long)
    SendMessage m_oComboBox.hwnd, CB_SETDROPPEDWIDTH, PropVal, ByVal 0
End Property

Public Property Get p_Picture() As Object
    Set p_Picture = m_oPicture
End Property

Public Property Set p_Picture(ByVal PropVal As Object)
    Set m_oPicture = PropVal
End Property

Public Property Get p_RestrictItems() As Boolean
    p_RestrictItems = m_bRestrictItems
End Property

Public Property Let p_RestrictItems(ByVal PropVal As Boolean)
    m_bRestrictItems = PropVal
End Property


Private Sub Class_Initialize()
    Set m_cMem = New Collection
End Sub

Private Sub BrowseForFolder()
'/* standard folder browsing dialog

Dim lList           As Long
Dim sTitle          As String
Dim tBrowseInfo     As BROWSEINFO
Dim sBuffer         As String
Dim sKey            As String

On Error Resume Next

    '/* title
    sTitle = "Select a Directory to Scan: "
    '/* fill struct
    With tBrowseInfo
        .hwndOwner = m_oComboBox.Container.hwnd
        .lpszTitle = lstrcat(sTitle, vbNullString)
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lList = SHBrowseForFolder(tBrowseInfo)
    '/* call dialog
    If lList Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lList, sBuffer
        sBuffer = left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        sKey = Get_SmallIcon(sBuffer)
        ComboAddItem sBuffer, sKey
        With m_oComboBox
            .ComboItems.Item(.ComboItems.Count).Selected = True
            RaiseEvent eFilePath(sBuffer)
            .Refresh
        End With
    End If
    
On Error GoTo 0
    
End Sub

Private Sub ComboAddItem(ByVal sItem As String, _
                        ByVal sKey As String, _
                        Optional ByVal lIndex As Long, _
                        Optional lIndent As Long)

'/* add an item

    If lIndex > 0 Then
        m_oComboBox.ComboItems.Add lIndex, , sItem, sKey, , lIndent
    Else
        m_oComboBox.ComboItems.Add , , sItem, sKey, , 1
    End If

End Sub

Private Sub ComboRemoveItem(ByVal lIndex As Long)
'/* remove an item

    m_oComboBox.ComboItems.Remove lIndex

End Sub

Public Function Directory_Exists(ByVal sDir As String) As Boolean
'/* can test any, file, any directory..

Dim lRes                As Long
Dim sPath               As String

    '/* prep var
    sPath = String$(255, 0)
    '/* get short path
    lRes = GetShortPathName(sDir, sPath, 254)
    '/* success
    Directory_Exists = lRes > 0

End Function

Private Sub Directory_Remove()
'/* clear the directory tree

Dim vI As Variant

    If m_cMem.Count > 0 Then
        For Each vI In m_cMem
            ComboRemoveItem CLng(vI)
        Next vI
    End If
    Set m_cMem = New Collection
    m_ExtDrive = ""
    
End Sub

Private Sub Directory_Return(ByVal sPath As String)
'/* dump folders to list

Dim bV              As Boolean
Dim lH              As Long
Dim sE              As String
Dim W32             As WFD
Dim lCount          As Long
Dim sKey            As String

On Error Resume Next

    '/* store the drive letter
    m_ExtDrive = sPath
    lH = FindFirstFile(sPath + "*.*", W32)
    bV = True
    Do
        With W32
            sE = left$(.cFN, InStr(1, .cFN, Chr$(0)) - 1)
        End With
        If Not LenB(sE) > 0 Then GoTo Skip
        If AscW(sE) = 46 Then GoTo Skip
        If W32.lFA And &H10 Then
            '/* get the index
            lCount = m_oComboBox.SelectedItem.Index + 1
            '/* icon from key
            sKey = Get_SmallIcon(sPath + sE)
            '/* add to list
            If Not InStr(1, sE, Chr$(58)) > 0 Then
                ComboAddItem sE, sKey, lCount, 3
            End If
            '/* remember
            m_cMem.Add lCount
            lCount = lCount + 1
        End If
Skip:
        bV = FindNextFile(lH, W32)
    Loop While bV
    FindClose lH
    
On Error GoTo 0

End Sub

Private Function Drive_Check() As Collection
'/* get drive strings

Dim lBuffer     As Long
Dim lCount      As Long
Dim sDrives     As String
Dim aDrives()   As String
Dim cTemp       As Collection

On Error Resume Next

    Set cTemp = New Collection
    '//get the buffer size
    lBuffer = GetLogicalDriveStrings(0, sDrives)
    '//set string len
    sDrives = String$(lBuffer, 0)
    '//get the drive list
    GetLogicalDriveStrings lBuffer, sDrives
    '/* split
    sDrives = left(sDrives, Len(sDrives) - 2)
    aDrives = Split(sDrives, vbNullChar)
    
    '/* loop through drives and add info
    For lCount = 0 To UBound(aDrives)
        If Not LCase$(aDrives(lCount)) = "a:\" Then
            cTemp.Add aDrives(lCount) + Drive_Label(aDrives(lCount))
        Else
            cTemp.Add aDrives(lCount)
        End If
    Next lCount
    
    '/* success
    If cTemp.Count > 0 Then
        Set Drive_Check = cTemp
    End If

On Error GoTo 0

End Function

Private Function Drive_Label(ByVal sPath As String) As String
'/* get drive label and fs type

Dim lFlags          As Long
Dim lMaxLen         As Long
Dim lSerial         As Long
Dim sName           As String * 256
Dim sType           As String * 256
Dim sTemp           As String

On Error Resume Next

    '/* test and shorten string
    If Len(sPath) > 3 Then
        sPath = left(sPath, 3)
    End If
    '/* get volume flags
    GetVolumeInformation sPath, sName, Len(sName), lSerial, lMaxLen, lFlags, sType, Len(sType)
    
    sTemp = left(sName, InStr(1, sName, vbNullChar) - 1)
    If Len(sTemp) > 0 Then
        Drive_Label = sTemp
    End If

    sTemp = left(sType, InStr(1, sType, vbNullChar) - 1)
    If Len(sTemp) > 0 Then
        If Len(Drive_Label) > 0 Then
            Drive_Label = "  [" & Drive_Label & " - " & sTemp & "]"
        Else
            Drive_Label = "  [" & sTemp & "]"
        End If
    End If

On Error GoTo 0

End Function

Private Function Get_Folder(SHFlag As eDirectories) As String
'/* get default folder locations

Dim lRes            As Long
Dim sPath           As String
Dim ItemIdL         As ITEMIDLIST

On Error GoTo Handler

    lRes = SHGetSpecialFolderLocation(100, SHFlag, ItemIdL)

    If lRes Then
        Get_Folder = vbNullString
    Else
        sPath = Space$(512)
        lRes = SHGetPathFromIDList(ByVal ItemIdL.mkid.cb, ByVal sPath)
        Get_Folder = left(sPath, InStr(sPath, vbNullChar) - 1)
    End If

Handler:
On Error GoTo 0

End Function

Private Function Get_SmallIcon(ByVal sFile As String) As String
'/* get listview icons
'/* names image keys with file extensions,
'/* for reusable image items in imagelist

Dim hSIcon      As Long
Dim imgObj      As ListImage
Dim sKey        As String

On Error Resume Next

    '/* get associated extension
    '/* get handle to icon
    hSIcon = SHGetFileInfo(sFile, 0&, tShInfo, Len(tShInfo), BASIC_SHGFI_FLAGS Or &H1)
    '/* load icon to picturebox
    If Not hSIcon = 0 Then
        With m_oPicture
            Set .Picture = LoadPicture("")
            .AutoRedraw = True
            ImageList_Draw hSIcon, tShInfo.iIcon, .hdc, 0&, 0&, &H1
            .Refresh
        End With
        '/* test for icon presence in collection
        sKey = m_cSIcon.Item(sFile)
        '/* if not present, add to collection
        If LenB(sKey) = 0 Then
            m_cSIcon.Add 1, sFile
            '/* add icon to image list
            '/* use file extension as image key
            Set imgObj = m_oImageList.ListImages.Add(Key:=sFile, Picture:=m_oPicture.Image)
        End If
        Get_SmallIcon = sFile
    Else
        '/* no icon, use default image
        Get_SmallIcon = "dft"
    End If

On Error GoTo 0

End Function

Public Sub ListDrives(Optional ByVal eFormat As LIST_FORMAT = Drives_Only)
'/* select the items displayed (change to your liking)

Dim vItem       As Variant
Dim sKey        As String
Dim sPath       As String

    m_eListMem = eFormat
    m_oComboBox.ComboItems.Clear
    
    Select Case eFormat
    Case 0
        ComboAddItem "Browse..", "dft"
        For Each vItem In Drive_Check
            sKey = Get_SmallIcon(left(CStr(vItem), 3))
            ComboAddItem vItem, sKey
        Next vItem

    Case 1
        ComboAddItem "Browse..", "dft"
        '/* docs
        sPath = Get_Folder(CSIDL_MYDOCUMENTS)
        sKey = Get_SmallIcon(sPath)
        ComboAddItem "My Documents", sKey
        '/* music
        sPath = Get_Folder(CSIDL_MYMUSIC)
        sKey = Get_SmallIcon(sPath)
        ComboAddItem "My Music", sKey
        '/* pictures
        sPath = Get_Folder(CSIDL_MYPICTURES)
        sKey = Get_SmallIcon(sPath)
        ComboAddItem "My Pictures", sKey
        '/* video
        sPath = Get_Folder(CSIDL_MYVIDEO)
        sKey = Get_SmallIcon(sPath)
        ComboAddItem "My Video", sKey
        '/* profile
        sPath = Get_Folder(CSIDL_PROFILE)
        sKey = Get_SmallIcon(sPath)
        ComboAddItem "Profile", sKey
        '/* programs
        sPath = Get_Folder(CSIDL_PROGRAM_FILES)
        sKey = Get_SmallIcon(sPath)
        ComboAddItem "Program Files", sKey

        '/* drive list
        For Each vItem In Drive_Check
            sKey = Get_SmallIcon(left(CStr(vItem), 3))
            ComboAddItem vItem, sKey
        Next vItem
    End Select
    
    '/* select the first item
    m_oComboBox.ComboItems.Item(1).Selected = True
    
End Sub

Private Sub m_oComboBox_Click()
'/* raise event with selected path

Dim sPath   As String

On Error Resume Next

    '/* safety valve
    If m_oComboBox.SelectedItem Is Nothing Then
        m_oComboBox.ComboItems.Clear
        ListDrives m_eListMem
        m_oComboBox.ComboItems.Item(1).Selected = True
        m_oComboBox.Container.Refresh
        Exit Sub
    End If
    Select Case m_oComboBox.SelectedItem.Text
    Case "Browse.."
        BrowseForFolder
    Case "My Documents"
        sPath = Get_Folder(CSIDL_MYDOCUMENTS) + Chr$(92)
    Case "My Music"
        sPath = Get_Folder(CSIDL_MYMUSIC)
    Case "My Pictures"
        sPath = Get_Folder(CSIDL_MYPICTURES)
    Case "My Video"
        sPath = Get_Folder(CSIDL_MYVIDEO)
    Case "Profile"
        sPath = Get_Folder(CSIDL_PROFILE)
    Case "Program Files"
        sPath = Get_Folder(CSIDL_PROGRAM_FILES)
    Case Else
        If InStr(1, m_oComboBox.SelectedItem.Text, Chr$(58)) > 0 Then
            sPath = left$(m_oComboBox.SelectedItem.Text, 2) + Chr$(92)
        Else
            sPath = m_ExtDrive + m_oComboBox.SelectedItem.Text
        End If
    End Select
    
    m_oComboBox.Container.Refresh
    RaiseEvent eFilePath(sPath)

On Error GoTo 0

End Sub

Private Sub m_oComboBox_GotFocus()
'/* resize entry

    If m_bAutoSelect Then
        m_oComboBox.SelStart = 0
        m_oComboBox.SelLength = Len(m_oComboBox.Text)
    End If

End Sub

Private Sub m_oComboBox_KeyPress(KeyAscii As Integer)
'/* auto complete

Dim found       As Boolean
Dim sPart       As String
Dim sTemp       As String
Dim bFound      As Boolean
Dim sComp       As String
Dim cbItem      As ComboItem

On Error GoTo Handler

    If Not (m_bAutoComplete Or m_bRestrictItems) Then Exit Sub
    Select Case KeyAscii
        '/* backspace - erase
    Case vbKeyBack
        If m_oComboBox.SelStart <= 1 Then
            m_oComboBox.Text = ""
            KeyAscii = 0
            Exit Sub
        End If
        If Not m_oComboBox.SelLength = 0 Then
            sPart = left$(m_oComboBox.Text, m_oComboBox.SelStart - 1)
        End If
        '/* enter
    Case vbKeyReturn, vbKeyTab
        m_oComboBox.Refresh
        Exit Sub
        '/* chars
    Case 32 Or (KeyAscii > 127 And KeyAscii < 161)
        Exit Sub
    Case Else
        sPart = left$(m_oComboBox.Text, m_oComboBox.SelStart) + Chr$(KeyAscii)
    End Select
    
    '/* loop though list and
    '/* get best match
    If m_bAutoComplete Or m_bRestrictItems Then
        For Each cbItem In m_oComboBox.ComboItems
            sTemp = left$(cbItem.Text, Len(sPart))
            If LCase$(sTemp) Like LCase$(sPart) Then
                If m_bAutoComplete Then
                    With m_oComboBox
                        Set .SelectedItem = cbItem
                        .SelStart = Len(sPart)
                        .SelLength = Len(.Text) - .SelStart
                    End With
                    KeyAscii = 0
                End If
                bFound = True
                Exit For
            End If
        Next cbItem
    End If
    If Not bFound Then
        If m_bRestrictItems Then
            KeyAscii = 0
        End If
    End If
    blnAuto = False

Handler:
On Error GoTo 0

End Sub

Private Sub SizeComboList()
'/* set list size to contents

Dim cbItem          As ComboItem
Dim sText           As String
Dim lHdc            As Long
Dim lLength         As Long
Dim tPnt            As POINTAPI

    lHdc = GetDC(m_oComboBox.Container.hwnd)

    For Each cbItem In m_oComboBox.ComboItems
        sText = cbItem.Text
        GetTextExtentPoint32 lHdc, sText, Len(sText) + 2, tPnt
        If tPnt.X > lLength Then
            lLength = tPnt.X
        End If
    Next cbItem

    lLength = lLength + 64
    SendMessage m_oComboBox.hwnd, CB_SETDROPPEDWIDTH, lLength, ByVal 0

End Sub

Public Sub SetReference(oCombo As ImageCombo)
'/* attatch the class

    Set m_oComboBox = oCombo
    
End Sub

Private Sub m_oComboBox_Dropdown()

Static bOnce    As Boolean

    '/* subclass the parent (to get the list handle only)
    If Not bOnce Then
        sc_Subclass m_oComboBox.hwnd
        sc_AddMsg m_oComboBox.hwnd, &H134, MSG_AFTER
        bOnce = True
    End If

    '/* auto size contents
    If m_bAutoSize Then
        SizeComboList
    End If
    
End Sub

Private Sub Class_Terminate()

On Error GoTo Handler

    '/* subclass end
     sc_Terminate

Handler:

End Sub


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>SubClass<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Private Sub zAddMsg(ByVal uMsg As Long, _
                    ByVal nTable As Long)

Dim nCount      As Long
Dim nBase       As Long
Dim i           As Long

    nBase = z_ScMem                                                         'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                                 'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                             'If ALL_MESSAGES are being added to the table...
        nCount = ALL_MESSAGES                                               'Set the table entry count to ALL_MESSAGES
    Else
        nCount = zData(0)                                                   'Get the current table entry count
        If nCount >= MSG_ENTRIES Then                                       'Check for message table overflow
            zError "zAddMsg", "Message table overflow. Either increase the value" & _
            "of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
            GoTo Bail
        End If

        For i = 1 To nCount                                                 'Loop through the table entries
            If zData(i) = 0 Then                                            'If the element is free...
                zData(i) = uMsg                                             'Use this element
                GoTo Bail
            ElseIf zData(i) = uMsg Then                                     'If the message is already in the table...
                GoTo Bail
            End If
        Next i                                                              'Next message table entry

        nCount = i                                                          'On drop through: i = nCount + 1, the new table entry count
        zData(nCount) = uMsg                                                'Store the message in the appended table entry
    End If

    zData(0) = nCount                                                       'Store the new table entry count
Bail:
    z_ScMem = nBase                                                         'Restore the value of z_ScMem

End Sub

Private Function zAddressOf(ByVal oCallback As Object, _
                            ByVal nOrdinal As Long) As Long

Dim bSub        As Byte
Dim bVal        As Byte
Dim nAddr       As Long
Dim i           As Long
Dim j           As Long


    RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                       'Get the address of the callback object's instance
    If Not zProbe(nAddr + &H1C, i, bSub) Then                               'Probe for a Class method
        If Not zProbe(nAddr + &H6F8, i, bSub) Then                          'Probe for a Form method
            If Not zProbe(nAddr + &H7A4, i, bSub) Then                      'Probe for a UserControl method
                Exit Function
            End If
        End If
    End If

    i = i + 4                                                               'Bump to the next entry
    j = i + 1024                                                            'Set a reasonable limit, scan 256 vTable entries
    Do While i < j
        RtlMoveMemory VarPtr(nAddr), i, 4                                   'Get the address stored in this vTable entry

        If IsBadCodePtr(nAddr) Then                                         'Is the entry an invalid code address?
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4         'Return the specified vTable entry address
            Exit Do                                                         'Bad method signature, quit loop
        End If

        RtlMoveMemory VarPtr(bVal), nAddr, 1                                'Get the byte pointed to by the vTable entry
        If Not bVal = bSub Then                                             'If the byte doesn't match the expected value...
            RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4         'Return the specified vTable entry address
            Exit Do                                                         'Bad method signature, quit loop
        End If

        i = i + 4                                                           'Next vTable entry
    Loop

End Function

Private Property Get zData(ByVal nIndex As Long) As Long

    RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4

End Property

Private Property Let zData(ByVal nIndex As Long, _
                           ByVal nValue As Long)

    RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4

End Property

Private Sub zDelMsg(ByVal uMsg As Long, _
                    ByVal nTable As Long)

Dim nCount      As Long
Dim nBase       As Long
Dim i           As Long

    nBase = z_ScMem                                             'Remember z_ScMem so that we can restore its value on exit
    z_ScMem = zData(nTable)                                     'Map zData() to the specified table

    If uMsg = ALL_MESSAGES Then                                 'If ALL_MESSAGES are being deleted from the table...
        zData(0) = 0                                            'Zero the table entry count
    Else
        nCount = zData(0)                                       'Get the table entry count
        For i = 1 To nCount                                     'Loop through the table entries
            If zData(i) = uMsg Then                             'If the message is found...
                zData(i) = 0                                    'Null the msg value -- also frees the element for re-use
                GoTo Bail
            End If
        Next i                                                  'Next message table entry
        zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If

Bail:
    z_ScMem = nBase                                             'Restore the value of z_ScMem

End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, _
                   ByVal sMsg As String)

    App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
    MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

Private Function zFnAddr(ByVal sDLL As String, _
                         ByVal sProc As String) As Long

    zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    'Get the specified procedure address
    Debug.Assert zFnAddr

End Function

Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long

    If z_Funk Is Nothing Then                           'Ensure that subclassing has been started
        zError "zMap_hWnd", "Subclassing hasn't been started"
    Else
        On Error GoTo Catch                             'Catch unsubclassed window handles
        z_ScMem = z_Funk("h" & lng_hWnd)                'Get the thunk address
        zMap_hWnd = z_ScMem
    End If

Exit Function                                           'Exit returning the thunk address

Catch:
    zError "zMap_hWnd", "Window handle isn't subclassed"

End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, _
                        ByRef nMethod As Long, _
                        ByRef bSub As Byte) As Boolean

Dim bVal        As Byte
Dim nAddr       As Long
Dim nLimit      As Long
Dim nEntry      As Long


    nAddr = nStart                                          'Start address
    nLimit = nAddr + 32                                     'Probe eight entries
    Do While nAddr < nLimit                                 'While we've not reached our probe depth
        RtlMoveMemory VarPtr(nEntry), nAddr, 4              'Get the vTable entry
        If Not nEntry = 0 Then                              'If not an implemented interface
            RtlMoveMemory VarPtr(bVal), nEntry, 1           'Get the value pointed at by the vTable entry
            If bVal = &H33 Or bVal = &HE9 Then              'Check for a native or pcode method signature
                nMethod = nAddr
                                                            'Store the vTable entry
                bSub = bVal                                 'Store the found method signature
                zProbe = True                               'Indicate success
                Exit Function                               'Return
            End If
        End If
        nAddr = nAddr + 4                                   'Next vTable entry
    Loop

End Function



Private Sub sc_AddMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then       'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                     'If the message is to be added to the before original WndProc table...
            zAddMsg uMsg, IDX_BTABLE                    'Add the message to the before table
        End If
        If When And MSG_AFTER Then                      'If message is to be added to the after original WndProc table...
            zAddMsg uMsg, IDX_ATABLE                    'Add the message to the after table
        End If
    End If

End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, _
                                    ByVal uMsg As Long, _
                                    ByVal wParam As Long, _
                                    ByVal lParam As Long) As Long

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then
        'Ensure that the thunk hasn't already released its memory
        sc_CallOrigWndProc = CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam)
    End If

End Function

Private Sub sc_DelMsg(ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then           'Ensure that the thunk hasn't already released its memory
        If When And MSG_BEFORE Then                         'If the message is to be deleted from the before original WndProc table...
            zDelMsg uMsg, IDX_BTABLE                        'Delete the message from the before table
        End If
        If When And MSG_AFTER Then                          'If the message is to be deleted from the after original WndProc table...
            zDelMsg uMsg, IDX_ATABLE                        'Delete the message from the after table
        End If
    End If

End Sub

Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then           'Ensure that the thunk hasn't already released its memory
        sc_lParamUser = zData(IDX_PARM_USER)                'Get the lParamUser callback parameter
    End If

End Property

Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, _
                                   ByVal NewValue As Long)

    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then           'Ensure that the thunk hasn't already released its memory
        zData(IDX_PARM_USER) = NewValue                     'Set the lParamUser callback parameter
    End If

End Property

'-SelfSub code------------------------------------------------------------------------------------
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                             Optional ByVal lParamUser As Long = 0, _
                             Optional ByVal nOrdinal As Long = 1, _
                             Optional ByVal oCallback As Object = Nothing, _
                             Optional ByVal bIdeSafety As Boolean = True) As Boolean

Const CODE_LEN     As Long = 260                                'Thunk length in bytes
Const MEM_LEN      As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))
Const PAGE_RWX     As Long = &H40&                              'Allocate executable memory
Const MEM_COMMIT   As Long = &H1000&                            'Commit allocated memory
Const MEM_RELEASE  As Long = &H8000&                            'Release allocated memory flag
Const IDX_EBMODE   As Long = 3                                  'Thunk data index of the EbMode function address
Const IDX_CWP      As Long = 4                                  'Thunk data index of the CallWindowProc function address
Const IDX_SWL      As Long = 5                                  'Thunk data index of the SetWindowsLong function address
Const IDX_FREE     As Long = 6                                  'Thunk data index of the VirtualFree function address
Const IDX_BADPTR   As Long = 7                                  'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER    As Long = 8                                  'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK As Long = 10                                 'Thunk data index of the callback method address
Const IDX_EBX      As Long = 16                                 'Thunk code patch index of the thunk data
Const SUB_NAME     As String = "sc_Subclass"                    'This routine's name
Dim nAddr          As Long
Dim nID            As Long
Dim nMyID          As Long


    If IsWindow(lng_hWnd) = 0 Then                              'Ensure the window handle is valid
        zError SUB_NAME, "Invalid window handle"
    Else
        nMyID = GetCurrentProcessId                             'Get this process's ID
        GetWindowThreadProcessId lng_hWnd, nID                  'Get the process ID associated with the window handle
        If Not nID = nMyID Then                                 'Ensure that the window handle doesn't belong to another process
            zError SUB_NAME, "Window handle belongs to another process"
            Exit Function
        End If

        If oCallback Is Nothing Then                            'If the user hasn't specified the callback owner
            Set oCallback = Me                                  'Then it is me
        End If

        nAddr = zAddressOf(oCallback, nOrdinal)                 'Get the address of the specified ordinal method
        If nAddr = 0 Then                                       'Ensure that we've found the ordinal method
            zError SUB_NAME, "Callback method not found"
            Exit Function
        End If

        If z_Funk Is Nothing Then                               'If this is the first time through, do the one-time initialization
            Set z_Funk = New Collection                         'Create the hWnd/thunk-address collection
            z_Sc(14) = &HD231C031
            z_Sc(15) = &HBBE58960
            z_Sc(17) = &H4339F631
            z_Sc(18) = &H4A21750C
            z_Sc(19) = &HE82C7B8B
            z_Sc(20) = &H74&
            z_Sc(21) = &H75147539
            z_Sc(22) = &H21E80F
            z_Sc(23) = &HD2310000
            z_Sc(24) = &HE8307B8B
            z_Sc(25) = &H60&
            z_Sc(26) = &H10C261
            z_Sc(27) = &H830C53FF
            z_Sc(28) = &HD77401F8
            z_Sc(29) = &H2874C085
            z_Sc(30) = &H2E8&
            z_Sc(31) = &HFFE9EB00
            z_Sc(32) = &H75FF3075
            z_Sc(33) = &H2875FF2C
            z_Sc(34) = &HFF2475FF
            z_Sc(35) = &H3FF2473
            z_Sc(36) = &H891053FF
            z_Sc(37) = &HBFF1C45
            z_Sc(38) = &H73396775
            z_Sc(39) = &H58627404
            z_Sc(40) = &H6A2473FF
            z_Sc(41) = &H873FFFC
            z_Sc(42) = &H891453FF
            z_Sc(43) = &H7589285D
            z_Sc(44) = &H3045C72C
            z_Sc(45) = &H8000&
            z_Sc(46) = &H8920458B
            z_Sc(47) = &H4589145D
            z_Sc(48) = &HC4836124
            z_Sc(49) = &H1862FF04
            z_Sc(50) = &H35E30F8B
            z_Sc(51) = &HA78C985
            z_Sc(52) = &H8B04C783
            z_Sc(53) = &HAFF22845
            z_Sc(54) = &H73FF2775
            z_Sc(55) = &H1C53FF28
            z_Sc(56) = &H438D1F75
            z_Sc(57) = &H144D8D34
            z_Sc(58) = &H1C458D50
            z_Sc(59) = &HFF3075FF
            z_Sc(60) = &H75FF2C75
            z_Sc(61) = &H873FF28
            z_Sc(62) = &HFF525150
            z_Sc(63) = &H53FF2073
            z_Sc(64) = &HC328&

            z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")    'Store CallWindowProc function address in the thunk data
            z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")     'Store the SetWindowLong function address in the thunk data
            z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")     'Store the VirtualFree function address in the thunk data
            z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")  'Store the IsBadCodePtr function address in the thunk data
        End If

        z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)    'Allocate executable memory

        If z_ScMem <> 0 Then                                        'Ensure the allocation succeeded
            On Error GoTo CatchDoubleSub                            'Catch double subclassing
            z_Funk.Add z_ScMem, "h" & lng_hWnd                      'Add the hWnd/thunk-address to the collection
            On Error GoTo 0

            If bIdeSafety Then                                      'If the user wants IDE protection
                z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")        'Store the EbMode function address in the thunk data
            End If

            z_Sc(IDX_EBX) = z_ScMem                                 'Patch the thunk data address
            z_Sc(IDX_HWND) = lng_hWnd                               'Store the window handle in the thunk data
            z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                   'Store the address of the before table in the thunk data
            z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4) 'Store the address of the after table in the thunk data
            z_Sc(IDX_OWNER) = ObjPtr(oCallback)                     'Store the callback owner's object address in the thunk data
            z_Sc(IDX_CALLBACK) = nAddr                              'Store the callback address in the thunk data
            z_Sc(IDX_PARM_USER) = lParamUser                        'Store the lParamUser callback parameter in the thunk data

            nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
            If nAddr = 0 Then                                       'Ensure the new WndProc was set correctly
                zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
                GoTo ReleaseMemory
            End If

            z_Sc(IDX_WNDPROC) = nAddr                               'Store the original WndProc address in the thunk data
            RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN        'Copy the thunk code/data to the allocated memory
            sc_Subclass = True                                      'Indicate success
        Else
            zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
        End If

        Exit Function                                               'Exit sc_Subclass

CatchDoubleSub:
        zError SUB_NAME, "Window handle is already subclassed"

ReleaseMemory:
        VirtualFree z_ScMem, 0, MEM_RELEASE                         'sc_Subclass has failed after memory allocation, so release the memory
    End If

End Function

Private Sub sc_Terminate()

Dim i       As Long

    If Not (z_Funk Is Nothing) Then                                 'Ensure that subclassing has been started
        With z_Funk
            For i = .Count To 1 Step -1                             'Loop through the collection of window handles in reverse order
                z_ScMem = .Item(i)                                  'Get the thunk address
                If IsBadCodePtr(z_ScMem) = 0 Then                   'Ensure that the thunk hasn't already released its memory
                    sc_UnSubclass zData(IDX_HWND)                   'UnSubclass
                End If
            Next i                                                  'Next member of the collection
        End With
        Set z_Funk = Nothing                                        'Destroy the hWnd/thunk-address collection
    End If

End Sub

Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)

    If z_Funk Is Nothing Then                                       'Ensure that subclassing has been started
        zError "sc_UnSubclass", "Window handle isn't subclassed"
    Else
        If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then               'Ensure that the thunk hasn't already released its memory
            zData(IDX_SHUTDOWN) = -1                                'Set the shutdown indicator
            zDelMsg ALL_MESSAGES, IDX_BTABLE                        'Delete all before messages
            zDelMsg ALL_MESSAGES, IDX_ATABLE                        'Delete all after messages
        End If
        z_Funk.Remove "h" & lng_hWnd                                'Remove the specified window handle from the collection
    End If

End Sub

'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)

Static bOnce    As Boolean
Static sIC      As Long
Dim sPath       As String

    If Not m_bDirectoryList Then Exit Sub

    Select Case uMsg
    '/* menu
    Case &H134
        '/* get the menu child handle from lparam
        If lng_hWnd = m_oComboBox.hwnd Then
            m_lMenuHandle = lParam
        End If
        '/* subclass the list
        If Not m_lMenuHandle = 0 Then
            If Not bOnce Then
                sc_Subclass m_lMenuHandle
                sc_AddMsg m_lMenuHandle, &H202, MSG_BEFORE
                sc_AddMsg m_lMenuHandle, &H203, MSG_BEFORE
                sc_DelMsg m_oComboBox.hwnd, &H134, MSG_AFTER
                bOnce = True
            End If
        End If

    '/* l-button up
    Case &H202
        On Error Resume Next
        If m_oComboBox.SelectedItem Is Nothing Then Exit Sub
        If InStr(1, m_oComboBox.SelectedItem.Text, Chr$(58)) > 0 Then
            sIC = sIC + 1
            If sIC < 3 Then
                bHandled = True
            Else
                sIC = 0
            End If
        End If
        On Error GoTo 0

    '/* added a l-button dbl-click evt
    Case &H203
        '/* test handle
        If lng_hWnd = m_lMenuHandle Then
            '/* is it a drive?
            On Error Resume Next
            If m_oComboBox.SelectedItem Is Nothing Then Exit Sub
            If InStr(1, m_oComboBox.SelectedItem.Text, Chr$(58)) > 0 Then
                bHandled = True
                '/* is it listed?
                If Not m_bExtended Then
                    '/* enumerate folders
                    sPath = left$(m_oComboBox.SelectedItem.Text, 2) + Chr$(92)
                    If Directory_Exists(sPath) Then
                        Directory_Return sPath
                        m_bExtended = True
                    End If
                Else
                    '/* retract list
                    Directory_Remove
                    m_bExtended = False
                End If
            End If
            On Error GoTo 0
        End If
    
    End Select
    
End Sub
