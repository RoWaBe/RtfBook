'RtfBook:     Free eBook reader. The used *.rbk format bases on RTF text files and JPEG/PNG picture files in a ZIP archive.
'             I personally had a special need for this format (and some tools around this) for scientific and author's use,
'             Plus, I used this project to get to know FreeBasic with it's numerous features. A phantastic language I must say!
'             So don't be astonished about some "zoological" ways with crazy pointers and so on. I really had to test those highly effective
'             features, even if they might look freaky sometimes ;-) My conclusion: With FreeBasic you can write phantastic fast and small
'             programs, The language has all I ever wanted to have; In FreeBasic You can write like in C without the disadvantages of C.
'Compiler:    FreeBasic for Windows 1.05
'             Aditionally we use static (!) LibZip library in FB directory: zip.bi, libzip.a and libz.a
'             see http://libzip.org/documentation/libzip.html
'Hint:        Don't edit Menus.rc using the the resource editor ResEd - it removes the attribute OWNERDRAW in menus,, for example:
'             MENUITEM "",IDC_ZOOMSMALLER.OWNERDRAW               will become          MENUITEM "",IDC_ZOOMSMALLER
'             The script RtfBook.rc you however _can_ edit using this very good resource editor.
'Developer:   Roland Walter 2018
'Last change: 10.08.2018, 17:30h
'************************************************************************************************************************
#define UNICODE

#include Once "windows.bi"
#include Once "win\shellapi.bi"
#include Once "win\windef.bi"
#include Once "win\commctrl.bi"
#include Once "win\commdlg.bi"
#include Once "win\richedit.bi"
#include Once "win\gdiplus.bi"
#include once "win\GdiPlus.bi"
#include once "win\GdiPlusInit.bi"
#include once "win\shlwapi.bi"
#include Once "zip.bi"          'ZIP archives library. We use static LibZip library and need: zip.bi, libzip.a and libz.a
#include Once "RtfBook.h"
'
Using GDIPLUS
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Newer RichEdit versions:
'#define EM_INSERTIMAGE       (WM_USER + 314), EM_INSERTTABLE  (WM_USER + 232)  ...
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
#define TXT_APP_NAME                    "RTF book viewer"  'Application name, used for windows/MessageBox caption bar, Registry entries
#define TXT_RTF_BOOK                    "RTF book: "       'Text for main window caption bar
#define TXT_RTFBOOK_EXTENSION           ".rbk"             'RTF book file extension to register for RTF book viewer
'
'String constatnts for the registry:
#define TXT_REGKEY_PATH                 "Software\"
#define TXT_REGKEY_WINDOWS              "Windows"
#define TXT_REGKEY_LASTARCHIVE          "LastArchiveFile"
#define TXT_REGKEY_LAST_CHAPTERFILE     "LastChapterFile"
#define TXT_REGKEY_LAST_CHAPTERPOSITION "LastChapterPosition"
#define TXT_REGKEY_STORE_SETTINGS       "StoreSettings"
#define TXT_REGKEY_NO_FILE_ASSOCIATION  "NoFileAssociation"
#define TXT_REGKEY_CRASH_FLAG           "IgnoreFile"
'
#define TXT_RBK_INDEXFILE               "#INDEX.TXT"                'Reserved filename in *.rbk eBooks
#define TXT_RBK_METAFILE                "#META.TXT"                 'Reserved filename in *.rbk eBooks
'
#define TXT_LINKID_BAGGAGE              "[BAGGAGE]"                 'Internal code for RTF link field (see EN_LINK): Extract the baggage (attachement) file currently selected 
'
#define TXT_UNICODE_SYSTEMFONT          "Lucida Sans Unicode"       'Unicode system font (since NT 3.1 and Windows 98) with symbol characters similiar to Arial
'
#define TXT_BOM_UTF8                    Chr(&HEF,&HBB,&HBF)         'Magic number for UTF8 text files
'#define TXT_BOM_UTF16LE2                Chr(&HFF,&HFE,&H5B,&H00)   'Rare occuring 4-bytes Magic number for UTF-16, Little Endian (from "Little End In")
'#define TXT_BOM_UTF16LE                 Chr(&HFF,&HFE)             'Magic number for UTF-16, Little Endian, same as before, but 2-bytes BOM mark
'#define TXT_BOM_UTF16BE2                Chr(&HFE,&HFF,&HBF,&H00)   'Magic number for UTF-16, Big Endian (from "Big End In"), incompatible to the Windows API
'#define TXT_BOM_UTF16BE                 Chr(&HFE,&HFF)             'Magic number for UTF-16, Big Endian, same as before, but 2-bytes BOM mark
'                                                                    UTF-16 Little Endian is the internal format of Windows 2000 und later.
'
#define COLOR_TOCTEXT    &H00000000            'TOC listbox: Text color is black (RGB value, owner drawn)
#define COLOR_TOCGRAY    &H00A0CFFF            'TOC listbox: Text color is gray, used if TOC hasn't focus (RGB value, owner drawn)
#define COLOR_TOCBK      &H00C0EFFF            'TOC listbox: Background color is paper white (RGB value, owner drawn)
#define COLOR_DARKGRAY   &H00808080            'Search dialog. Used if searched for colored text (RGB value)
#define TOC_PRE2_TEXT        &H0030            'TOC 2nd prefix character "0": Text. Ignored, invisible, reserved for metainfo files
#define TOC_PRE2_METAINFO    &H0031            'TOC 2nd prefix character "1": Metainfo file "#INFO.TXT"
#define TOC_PRE2_RTF         &H0032            'TOC 2nd prefix character "2": RTF text file
#define TOC_PRE2_PICTURE     &H0033            'TOC 2nd prefix character "3": Picture file, JPEG or PNG
#define TOC_PRE2_ATTACHEMENT &H0034            'TOC 2nd prefix character "4": Binary attachement file, any unknown format
'
#define QUICKSEARCH_TIMEOUT_MS 1000            'Quick search timeout in milliseconds. In this time the user can add characters to the search term. After the term is reset to "".
'
#define EXP_CHAR_RTF         &H25AA            'Prefix character in "Export" listbox for RTF text
#define EXP_CHAR_PIC         &H25AB            'Prefix character in "Export" listbox for picture file
'
#define ZOOM_MINIMAL         1                 'Zoom level in Richedit text window
#define ZOOM_NORMAL          20                'Zoom level in Richedit text window
#define ZOOM_MAXIMAL         80                'Zoom level in Richedit text window. Value 80: Just to have any end of zoom.
'
#define MINIMAL_TOCWNDWIDTH 2                  'minimal with in pixels of the TableOfContent child window
#define WIDTH_VERT_LINE     5                  'with in pixels of the "drag line" window, which is between the TableOfContent window and the text window
'
#define INVALID_ZIP_INDEX    -1                'For LibZip: index of file in ZIP archive is invalis (the indexes ar 0 based). File zip.bi doesn't have such a definition.
'
#define IDM_POPUP 1000
'
Dim Shared swAppFile As wString*MAX_PATH                 'Current (real) path and filename of application
Dim Shared swIniFile As wString*MAX_PATH                 'Current (real) path and filename of application's ini fle
Dim Shared swAppPath As wString*MAX_PATH                 'Current (real) path and filename of application with ending backslash
Dim Shared szArchivePassword As ZString*512              'Password for encrypted files. Changed (or not) in dialog function PasswordDlgProc()
Dim Shared fDontAskPasswordForThisBook As Integer        'Option "Password ... Dont ask again for this book"
Dim Shared sMetaData As String                           'Content of current #meta.txt file if existing or empty string. 
Dim Shared swAppName As wString*MAX_PATH                 'Application's name
Dim Shared hwndMain As HWND                              'Main window
Dim Shared hMenuMain As HMENU                            'Main window's menu
Dim Shared hLibRichEdit As HMODULE
Dim Shared hwndTOC As HWND                               'Listbox with the Table of Content tree
Dim Shared hBrushTocListBk As HBRUSH                     'For the ownerdrawn TOC listbox, text and icon background  hBrushTocListBk As HBRUSH
Dim Shared hWndVertLine As HWND                          'Vertical line between TOC window and text window
Dim Shared hwndShadowTOC_Unsorted As HWND                'Invisible listbox for the raw file list, (raw Table of Content)
Dim Shared hwndShadowTOC_Sorted As HWND                  'Invisible listbox for the raw file list, (raw Table of Content)
Dim Shared hwndShadowTOC As HWND                         'Invisible listbox for the raw file list, (raw Table of Content) receives either hwndShadowTOC_Sorted or hwndShadowTOC_Unsorted
Dim Shared hwndText As HWND                              'Richedit window used, if a text is to disply   (hidden if currently a picture is to display)
Dim Shared hwndPicture  As HWND                          'Standard child window used, if a picture is to display (hidden if currently a text is to display)
Dim Shared hDlgHelp  As HWND                             'Modeless Help dialog window
Dim Shared hDlgSearch As HWND                            'Modeless Search dialog window
Dim Shared hProgInst As HINSTANCE                        'Program instance
Dim Shared pfncOrigTxtWndProc As UInteger                'Function pointer to the RichEdit default window's proc (we subclass this control) Any Ptr???
Dim Shared pfncOrigTocWndProc As UInteger                'Function pointer to the TOC default window's proc (we subclass this control) Any Ptr???
Dim Shared pCurZipArchive As Any Ptr                     'LibZip's ZIP pointer to the ZIP archive currently open or NULL if no ZIP book is open.
Dim Shared hTocFont As HFONT                             'Font used for main window's TOC listbox as well as for DialogBox TOC listboxes
Dim Shared dwHeightOfTocFont As Dword                    'Y size of the TOC font
'
Type UserSettings
  fStoreSettings As Long                         'FALSE: Don't store user settings TRUE: Store   wLanguageID As Ushort                          'User language ID as defined by LoByte(GetUserDefaultUILanguage()) (examples: &H00=Neutral (english), &H07=German, &H03=Catalan and so on )
  fLanguageAutomatic As UShort                   'User language Ianguage will be choosed automaticly by system language
  wLanguageID As Long                            'User language ID as defined by LoByte(GetUserDefaultUILanguage()) (examples: &H00=Neutral (english), &H07=German, &H03=Catalan and so on )
  WinState As Long
  WinX As Long
  WinY As Long
  WinWidth As Long
  WinHeight As Long
  TocWinWidth As Long                            'Current width of the TableOfContent window
  TextZoom As Long                               'Current text zoom in the RichEdit control. Must be in the range 1...50, where 20 stands for "normal Zoom" (Ratio 1:20)
  fShowPassword As Long                          'Password input dialog
  HelpWinState As Long
  HelpWinX As Long
  HelpWinY As Long
  HelpWinWidth As Long
  HelpWinHeight As Long
  dwCurHelpChapter As Dword                       '0-based ID of current help chapter (to restore if Help opened multiple time)
  dwHelpScrollPos As Dword                        'Current scroll position in Help window (to restore if Help opened multiple time)
  dwHelpTextPos As Dword                          'Current text position in Help window (to restore if Help opened multiple time)
End Type
Dim Shared UserSettings As UserSettings
'
Type PrintInfos                                   'The printing thread parameters
  hwndMain As HWND                                'Handle of the main application window
  hInst As HINSTANCE                              'Instance handle of the application
  hwndText As HWND                                'Handle of the RichEdit control, which's text is to print
  lLeftMargin As Long                             'Page margin in Millimeters
  lRightMargin As Long                            'Page margin in Millimeters
  lTopMargin As Long                              'Page margin in Millimeters
  lBottomMargin As Long                           'Page margin in Millimeters
End Type
'
'Type CHARFORMAT2_CORR Field=1                             'Unicode version. Incorrect in richedit.bi  (CHARFORMAT: 60 bytes, CHARFORMAT2:)
'  cbSize As Dword
'  dwMask As Dword
'  dwEffects As Dword
'  yHeight As Long
'  yOffset As Long                ' > 0 for superscript, < 0 for subscript
'  crTextColor As Dword
'  bCharSet As Byte
'  bPitchAndFamily As Byte
'  szFaceName As ZString * LF_FACESIZE
'  wFiller As WORD
'  wWeight As WORD                ' Font weight (LOGFONT value)
'  sSpacing As Integer              ' Amount to space between letters
'  crBackColor As Dword           ' Background color
'  lcid As Dword                  ' Locale ID
'  dwReserved As Dword            ' Reserved. Must be 0
'  sStyle As Integer              ' Style handle
'  wKerning As Word               ' Twip size above which to kern char pair
'  bUnderlineType As Byte         ' Underline type
'  bAnimation As Byte             ' Animated text like marching ants
'  bRevAuthor As Byte             ' Revision author index
'  bReserved1  As Byte   ' <- added..
'End Type

Type CHARFORMAT2_CORR Field = 4   'Unicode version. Incorrect in richedit.bi  (CHARFORMAT: 60 bytes, CHARFORMAT2: 115/120 bytes)
	cbSize As UINT
	dwMask As Dword
	dwEffects As Dword
	yHeight As Long
	yOffset As Long
	crTextColor As COLORREF
	bCharSet As UByte
	bPitchAndFamily As UByte
	szFaceName As Wstring * 32
	wWeight As Word
	sSpacing As Short
	crBackColor As COLORREF
	lcid As LCID
  dwReserved As Dword
	sStyle As Short
	wKerning As Word
	bUnderlineType As UByte
	bAnimation As UByte
	bRevAuthor As UByte
'	bUnderlineColor As UByte
'  bReserved1  As Byte             ' <- added..
End Type


'
Dim Shared UserSettings_swzCurArchiveName As Wstring*MAX_PATH       'File currently opened, must be "" if no file is open.
Dim Shared UserSettings_swzLastArchiveName As Wstring*MAX_PATH      'Path+FileName of the last file
Dim Shared UserSettings_lLastChapterTitle As Long                   'ID of TOC title of the last chapter (0=first is default)
Dim Shared UserSettings_dwLastChapterPos As DWORD                   'Last Position in the last chapter file
Dim Shared UserSettings_szwFindText As Wstring*80                   'Current Find string
Dim Shared UserSettings_fNoFilenameAssociation As Integer           'Extension .rbk shouldn't be associated with eFolksBook
Dim Shared UserSettings_fIgnoreFileAtStart As Integer                       'True after program start, False at program end. For the case, that auto-loaded files make crahes
'
Declare Function WinMain(ByVal hInstance As HINSTANCE,ByVal hPrevInstance As HINSTANCE,ByVal lpCmdLine As LPSTR,ByVal iCmdShow As Integer) As Integer
Declare Function MainWndProc(ByVal hwnd As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Sub UpdateMainMenu(ByVal hWnd As HWND)
Declare Sub ResizeChildWindows(ByVal hWndParent As HWND,ByVal lTocWinWidth As Long)
Declare Function OnDrawItems(ByVal hWnd As HWND,ByVal lpdis As DRAWITEMSTRUCT Ptr) As LRESULT
Declare Function OpenBook(ByVal hWnd As HWND,ByVal sBookFile As String) As LRESULT
Declare Function IndexFile_get_name(ByVal pCurZipArchive As Any Ptr,ByVal fInit As Integer,ByVal idxZipIndexFile As Integer) As Const wString Ptr
Declare Function LoadBookChapter(ByVal iChapter As Integer) As LRESULT
Declare Function TextWndProc(ByVal hWnd As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function TocWndProc(ByVal hWnd As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function ExtractCurFile(ByVal hWndParent As HWND) As LRESULT
Declare Function ExportDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function ExportRtf(ByVal hwndParent As HWND,ByVal hExportList As HWND,Byref szwFilename As WSTRING) As LRESULT
Declare Function ExportKeyNote(ByVal hwndParent As HWND,ByVal hExportList As HWND,Byref szwFilename As Wstring) As LRESULT
Declare Function SettingsDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function PasswordDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function SearchDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function HelpDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function WStringToZString(ByVal sInputString As ZString Ptr,ByVal iOutputCodepage As Integer) As String
Declare Function ConvertStringEncoding(ByVal sInputString As ZString Ptr,ByVal iInputCodepage As Integer,ByVal iOutputCodepage As Integer) As String
'Declare Function SetRegistryBinData(ByVal hLocation As HKEY,ByVal sSubKeys As String,ByVal sValueName As String,ByVal pData As Byte Ptr,ByVal dwLenData As Dword) As LRESULT
'Declare Function GetRegistryBinData(ByVal hLocation As HKEY,ByVal sSubKeys As String,ByVal sValueName As String,ByVal pData As Byte Ptr,ByVal dwLenData As Dword) As LRESULT
Declare Function PictureToRtf(ByVal hInBuffer As HGLOBAL) As HGLOBAL
Declare Function MetaInfoToRtf(ByVal pInBuffer As Any Ptr) As Any Ptr
Declare Sub ShowSystemErrorMessage(ByVal hWndParent As HWND,ByVal sErrorText As String)
Declare Sub RegisterFileExtension()
Declare Sub UnRegisterFileExtension()
Declare Function PrintRichText(ByVal tPrintInfos As PrintInfos Ptr) As LRESULT
Declare Function WaitDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
Declare Function StreamInRtfFile(ByVal hFile As Dword,ByVal pBuffer As Any Ptr,ByVal dwBytesToRead As Dword,Byref dwBytesDone As Dword) As LRESULT
Declare Function StreamOutRtfFile(ByVal hFile As Dword,ByVal pbBuffer As Any Ptr,ByVal dwBytesToWrite As Dword,Byref dwBytesDone As Dword) As LRESULT
'. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 
'OLE stuff necessary to display pictures in a RichEdit control:
Type RichCom_OLEObject                                           'OLE stuff
  pIntf As Dword Ptr
  Refcount As Dword
End Type
Dim Shared pObj_RichCom As Dword Ptr                             'OLE stuff
Dim Shared pObj_RichComObject As RichCom_OLEObject               'OLE stuff
Dim Shared nObj_RichComObjectcnt As Long                         'OLE stuff
'
Declare Sub RichCom_SetComInterface(ByVal hWndEdit As HWND)                                                                                             'OLE stuff
Declare Function RichCom_Object_QueryInterface(pObject As Dword,REFIID As Dword,ppvObj As Dword) As Dword                                               'OLE stuff
Declare Function RichCom_Object_AddRef(ByVal pObject As Dword Ptr) As Dword                                                                             'OLE stuff
Declare Function RichCom_Object_Release(ByVal pObject As Dword Ptr) As Dword                                                                            'OLE stuff
Declare Function RichCom_Object_GetInPlaceContext(ByVal pObject As Dword Ptr,lplpFrame As Dword,lplpDoc As Dword,lpFrameInfo As Dword) As Dword         'OLE stuff
Declare Function RichCom_Object_ShowContainerUI(ByVal pObject As Dword Ptr,fShow As Long) As Dword                                                      'OLE stuff
Declare Function RichCom_Object_QueryInsertObject(ByVal pObject As Dword Ptr,lpclsid As Dword,ByVal lpstg As Dword Ptr,cp As Long) As Dword             'OLE stuff
Declare Function RichCom_Object_DeleteObject(ByVal pObject As Dword Ptr,lpoleobj As Dword) As Dword                                                     'OLE stuff
Declare Function RichCom_Object_QueryAcceptData(ByVal pObject As Dword Ptr,lpdataobj As Dword,lpcfFormat As Dword,reco As Dword,fReally As Long,hMetaPict As Dword) As Dword  'OLE stuff
Declare Function RichCom_Object_ContextSensitiveHelp(ByVal pObject As Dword Ptr,fEnterMode As Long) As Dword                                            'OLE stuff
Declare Function RichCom_Object_GetClipboardData(ByVal pObject As Dword Ptr,lpchrg As Dword,reco As Dword,lplpdataobj As Dword) As Dword                'OLE stuff
Declare Function RichCom_Object_GetDragDropEffect(ByVal pObject As Dword Ptr,fDrag As Long,grfKeyState As Dword,pdwEffect As Dword) As Dword            'OLE stuff
Declare Function RichCom_Object_GetContextMenu(ByVal pObject As Dword Ptr,seltype As Word,lpoleobj As Dword,lpchrg As Dword,lphmenu As Dword) As Dword  'OLE stuff
Declare Function RichCom_Object_GetNewStorage(ByVal pObject As Dword Ptr,lplpstg As Dword) As Dword                                                     'OLE stuff
Declare Sub TestTime(ByVal iMode As Integer) 'Time measurement for debugging. 0 starts, 1 ends and shows result
'. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . 
End WinMain(GetModuleHandle(NULL),NULL,command$,SW_NORMAL)
'************************************************************************************************************************
Function WinMain(ByVal hInstance As HINSTANCE,ByVal hPrevInstance As HINSTANCE,ByVal lpCmdLine As LPSTR,ByVal iCmdShow As Integer) As Integer
  Static MainWndClass As WNDCLASSEX,Msg As MSG,swMainWndClass As wString*32,hAccTable As HACCEL,i As Integer,rct As RECT,swBuffer As wString*256
  Static hRegKey As HKEY,tTextMetric As TEXTMETRIC
  '
  hProgInst      = hInstance
  swAppName      = TXT_APP_NAME                                    'For window titles and similiar
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'RichEdit 4.1 is available in WindowsXP SP1 and greater. It has a new dll name MSFTEDIT.DLL and a new classname "RICHEDIT50W".
  'msftedit.dll is available on XP SP1 only and it is not redistributable on down platforms. (Rhett Gong [MS] Microsoft Online Partner Support)
  hLibRichEdit=LoadLibrary("Msftedit")   'For the "RichEdit50W" class (much faster than riched20.dll and some of the features we really need)
  If hLibRichEdit=NULL Then
    'ShowErrorMessage(hWnd,IDTXT_RICHDLL,"")
    GoTo WinMainDone                         'End of program
  End If
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  InitCommonControls()
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  UserSettings_swzCurArchiveName=Trim(*lpCmdLine,Chr(34))            'Double-quotes trimming of filename given by command line (if given)
  '
  GetModuleFileName(hProgInst,swAppFile,SizeOf(swAppFile))           'Get current path and filename of application (all uppercase unfortunately, never understood the reason...)
  GetLongPathName(swAppFile,swAppFile,SizeOf(swAppFile))             'Get mixed case path
  swAppPath=Left$(swAppFile,InStrRev(swAppFile,"\"))                 'Result is fo rexample "c:\Projects\Test\" (with ending backslash)
  swIniFile=Left$(swAppFile,InStr(UCase$(swAppFile),".EXE"))+"ini"   'Example: "C:\TESTDIR\TEST.INI" (if program filename "TEST.EXE")
  'MessageBox(hwndMain,"Cur file: "+UserSettings_swzCurArchiveName+Chr$(13)+"App path: "+swAppPath+Chr$(13)+"ini file: "+swIniFile,"Debug",0)
  '
  GetTextMetrics(GetDC(NULL),@tTextMetric)                           'Get font metrics of the desktop font and...
  dwHeightOfTocFont=tTextMetric.tmHeight+3                           '...calculate based on this the font for the Table of Contents
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'Get the program settings stored in the Registry (if existing):
  UserSettings_fIgnoreFileAtStart=True
  If RegOpenKeyEx(HKEY_CURRENT_USER,TXT_REGKEY_PATH+TXT_APP_NAME,0,KEY_ALL_ACCESS,@hRegKey)=ERROR_SUCCESS Then   'KEY_READ isn't enough her
    i=SizeOf(UserSettings_fIgnoreFileAtStart)
    RegQueryValueEx(hRegKey,TXT_REGKEY_CRASH_FLAG,0,NULL,Cast(Byte Ptr,@UserSettings_fIgnoreFileAtStart),@i)     'Did last session end with crash? (than flag is TRUE)
    '
    i=True
    RegSetValueEx(hRegKey,TXT_REGKEY_CRASH_FLAG,0,REG_DWORD,Cast(Const Byte Ptr,@i),SizeOf(i))                   'Crash flag in Registry - reset at normal program end.
    '
    i=SizeOf(UserSettings_fNoFilenameAssociation)
    RegQueryValueEx(hRegKey,TXT_REGKEY_NO_FILE_ASSOCIATION,0,NULL,Cast(Byte Ptr,@UserSettings_fNoFilenameAssociation),@i)
    i=SizeOf(Long)
    RegQueryValueEx(hRegKey,TXT_REGKEY_STORE_SETTINGS,0,NULL,Cast(Byte Ptr,@UserSettings.fStoreSettings),@i)
    i=SizeOf(UserSettings_dwLastChapterPos)
    RegQueryValueEx(hRegKey,TXT_REGKEY_LAST_CHAPTERPOSITION,0,NULL,Cast(Byte Ptr,@UserSettings_dwLastChapterPos),@i)
    i=SizeOf(Long)
    RegQueryValueEx(hRegKey,TXT_REGKEY_LAST_CHAPTERFILE,0,NULL,Cast(Byte Ptr,@UserSettings_lLastChapterTitle),@i)
    If UserSettings_fIgnoreFileAtStart=True Then   'Last file caused a crash. Don't load automaticly (this whould give an deadly endless loop for the user).
      UserSettings_swzLastArchiveName=""
    Else
      i=SizeOf(UserSettings_swzLastArchiveName)
      RegQueryValueEx(hRegKey,TXT_REGKEY_LASTARCHIVE,0,NULL,Cast(Byte Ptr,@UserSettings_swzLastArchiveName),@i)
    End If
    i=SizeOf(UserSettings)
    RegQueryValueEx(hRegKey,TXT_REGKEY_WINDOWS,0,NULL,Cast(Byte Ptr,@UserSettings),@i)
    RegCloseKey(hRegKey)
    '
    If UserSettings.WinX<1 Then  '
      UserSettings.WinX=0
    ElseIf UserSettings.WinX>GetSystemMetrics(SM_CXSCREEN)-20 Then 
      UserSettings.WinX=GetSystemMetrics(SM_CXSCREEN)-20
    End If
    If UserSettings.WinY<0 Then
      UserSettings.WinY=0
    ElseIf UserSettings.WinY>GetSystemMetrics(SM_CYSCREEN)-20 Then
      UserSettings.WinY=GetSystemMetrics(SM_CYSCREEN)-20
    End If
    If UserSettings.WinWidth<100 Then
      UserSettings.WinWidth=100
    ElseIf UserSettings.WinWidth>GetSystemMetrics(SM_CXSCREEN) Then
      UserSettings.WinWidth=GetSystemMetrics(SM_CXSCREEN)
    End If
    If UserSettings.WinHeight>GetSystemMetrics(SM_CYSCREEN) Then
      UserSettings.WinHeight=GetSystemMetrics(SM_CYSCREEN)
    ElseIf UserSettings.WinHeight<100 Then
      UserSettings.WinHeight=100
    End If
    If UserSettings.TocWinWidth<2 Then
      UserSettings.TocWinWidth=UserSettings.WinWidth\6
    ElseIf UserSettings.TocWinWidth>UserSettings.WinWidth-20 Then
      UserSettings.TocWinWidth=UserSettings.WinWidth\6
    End If
    If UserSettings.WinState<>SW_SHOWMAXIMIZED Then UserSettings.WinState=SW_SHOWNORMAL        'State of the main window: Only SW_SHOWMAXIMIZED or SW_SHOWNORMAL is allowed
    If UserSettings.TextZoom<1 Or UserSettings.TextZoom>50 Then UserSettings.TextZoom=20        'text zoom in the RichEdit control. Must be in the range 1...50, where 20 stands for "normal Zoom" (Ratio 1:20)
    If UserSettings.fLanguageAutomatic<>False Then UserSettings.wLanguageID=LoByte(GetUserDefaultUILanguage())  'User language Ianguage will be choosed automaticly by system language)
  Else       'Settings not got from Registry
    UserSettings.fStoreSettings=False
    UserSettings.fLanguageAutomatic=True                                'User language Ianguage will be choosed automaticly by system language
    UserSettings.WinX=(GetSystemMetrics(SM_CXSCREEN)\2)-20
    UserSettings.WinY=0
    UserSettings.WinWidth=(GetSystemMetrics(SM_CXSCREEN)\2)+20
    UserSettings.WinHeight=GetSystemMetrics(SM_CYSCREEN)
    UserSettings.TocWinWidth=250
    UserSettings.WinState=SW_SHOWNORMAL                          'State of the main window: Only SW_SHOWMAXIMIZED or SW_SHOWNORMAL is allowed
    UserSettings.TextZoom=20                                     'text zoom in the RichEdit control. Must be in the range 1...50, where 20 stands for "normal Zoom" (Ratio 1:20)
    UserSettings.wLanguageID=LoByte(GetUserDefaultUILanguage())  'User language Ianguage will be choosed by hand (not automaticly by system language)
    UserSettings.dwCurHelpChapter=0
    UserSettings.dwHelpTextPos=0
    UserSettings.dwHelpScrollPos=0
    UserSettings.HelpWinX=GetSystemMetrics(SM_CXSCREEN)\2
    UserSettings.HelpWinY=0
    UserSettings.HelpWinWidth=GetSystemMetrics(SM_CXSCREEN)\2
    UserSettings.HelpWinHeight=GetSystemMetrics(SM_CYSCREEN)
  End If
  '
  If UserSettings_swzCurArchiveName="" Then UserSettings_swzCurArchiveName=UserSettings_swzLastArchiveName 'Empty command line: Use last file
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'Lucida Sans Unicode
  swMainWndClass              = "RtfBookMainWndClass"
  MainWndClass.cbSize         = SizeOf(WNDCLASSEX)
  MainWndClass.style          = CS_HREDRAW Or CS_VREDRAW
  MainWndClass.lpfnWndProc    = @MainWndProc
  MainWndClass.cbClsExtra     = 0
  MainWndClass.cbWndExtra     = 0
  MainWndClass.hInstance      = hProgInst
  MainWndClass.hIcon          = LoadIcon(hInstance,ByVal MAKEINTRESOURCE(IDRI_MAINAPP_ICO))
  MainWndClass.hCursor        = LoadCursor(NULL,ByVal IDC_ARROW)
  MainWndClass.hbrBackground  = GetSysColorBrush(CTLCOLOR_DLG)      'GetSysColorBrush(COLOR_ACTIVEBORDER)  'GetStockObject(BLACK_BRUSH)  COLOR_ACTIVEBORDER,COLOR_WINDOW,COLOR_MENU      
  If UserSettings.wLanguageID=&H07 Then  '&H00=Neutral (english), &H07=German, &H03=Catalan and so on, Value of LoByte(GetUserDefaultUILanguage())
    MainWndClass.lpszMenuName = MAKEINTRESOURCE(IDM_MAIN_GERMAN) 'IDM_MAIN_GERMAN,IDM_MAIN
  Else
    MainWndClass.lpszMenuName = MAKEINTRESOURCE(IDM_MAIN) 'IDM_MAIN_GERMAN,IDM_MAIN
  End If
  MainWndClass.lpszClassName  = @swMainWndClass
  MainWndClass.hIconSm        = LoadIcon(hProgInst,ByVal MAKEINTRESOURCE(IDRI_MAINAPP_ICO))
  RegisterClassEx(@MainWndClass)
  '
  hWndMain=CreateWindow(swMainWndClass,swAppName,WS_OVERLAPPEDWINDOW,UserSettings.WinX,UserSettings.WinY,UserSettings.WinWidth,UserSettings.WinHeight,HWND_DESKTOP,0,hProgInst,ByVal NULL)
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  ShowWindow(hwndMain,iCmdShow)
  UpdateWindow(hWndMain)
  '
  hAccTable=LoadAccelerators(hInstance,ByVal MAKEINTRESOURCE(1))                       'Load the accelerator table (we have only one)
  '
  While GetMessage(@Msg,NULL,0,0)
    If IsDialogMessage(hDlgSearch,@Msg)=0 Then                    'Check wether the message is for the (modeless) Search dialog window
      If IsDialogMessage(hDlgHelp,@Msg)=0 Then                    'Check wether the message is for the (modeless) Help dialog window
        If TranslateAccelerator(hWndMain,hAccTable,@Msg)=0 Then
          TranslateMessage(@Msg)
          DispatchMessage(@Msg)
        End If
      End If
    End If
  Wend
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'Free the used objects:
  UnregisterClass(swMainWndClass,hInstance)
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  If UserSettings.fStoreSettings=True Then
    If RegCreateKeyEx(HKEY_CURRENT_USER,TXT_REGKEY_PATH+TXT_APP_NAME,0,"",0,KEY_WRITE,ByVal NULL,@hRegKey,ByVal NULL) = ERROR_SUCCESS Then
      UserSettings_fIgnoreFileAtStart=False                                                                                       'False: ession didn't end with crash
      RegSetValueEx(hRegKey,TXT_REGKEY_CRASH_FLAG,0,REG_DWORD,Cast(Const Byte Ptr,@UserSettings_fIgnoreFileAtStart),SizeOf(LONG)) 'This is for next pogram start only
      RegSetValueEx(hRegKey,TXT_REGKEY_STORE_SETTINGS,0,REG_DWORD,Cast(Const Byte Ptr,@UserSettings.fStoreSettings),SizeOf(LONG))
      RegSetValueEx(hRegKey,TXT_REGKEY_NO_FILE_ASSOCIATION,0,REG_DWORD,Cast(Const Byte Ptr,@UserSettings_fNoFilenameAssociation),SizeOf(UserSettings_fNoFilenameAssociation))
      RegSetValueEx(hRegKey,TXT_REGKEY_LASTARCHIVE,0,REG_SZ,Cast(Const Byte Ptr,@UserSettings_swzLastArchiveName),Len(UserSettings_swzLastArchiveName)*2) 'Unicode! Len() returns charcters, not bytes!
      RegSetValueEx(hRegKey,TXT_REGKEY_LAST_CHAPTERFILE,0,REG_DWORD,Cast(Const Byte Ptr,@UserSettings_lLastChapterTitle),SizeOf(UserSettings_lLastChapterTitle))
      RegSetValueEx(hRegKey,TXT_REGKEY_LAST_CHAPTERPOSITION,0,REG_DWORD,Cast(Const Byte Ptr,@UserSettings_dwLastChapterPos),SizeOf(UserSettings_dwLastChapterPos))
      RegSetValueEx(hRegKey,TXT_REGKEY_WINDOWS,0,REG_BINARY,Cast(Const Byte Ptr,@UserSettings),SizeOf(UserSettings))
      RegCloseKey(hRegKey)
      '
      RegisterFileExtension()   'Writes only if not just registered
    End If
  Else
    swBuffer=TXT_REGKEY_PATH+TXT_APP_NAME
    If RegDeleteKey(HKEY_CURRENT_USER,swBuffer) <> ERROR_SUCCESS Then
      If SHDeleteKey(hRegKey,swBuffer) <> ERROR_SUCCESS Then RegDeleteKey(HKEY_CURRENT_USER,swBuffer)
    End If
    '
    UnRegisterFileExtension()
  End If
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  WinMainDone:
  Function=Msg.wParam
End Function
'************************************************************************************************************************
Function MainWndProc(ByVal hWndMainWindow As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  Static fResizingTocWindow As Long,i As Integer,j As Integer,k As Integer,pt As Point,rct As RECT,scri As SCROLLINFO,hOldMenu As HMENU
  Static hBmp As HBITMAP,swBuffer As wString*1024,tPrintInfos As PrintInfos,tCharRange As CHARRANGE,tCharFormat As CHARFORMAT2_CORR
  Static wplc As WINDOWPLACEMENT,pEnLink As ENLINK Ptr,pNmHdr As NMHDR Ptr,tTextRange As TEXTRANGE

  Static tOfn As OPENFILENAME,swInitialFileName As Wstring*MAX_PATH
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Select Case(uMsg)
  Case WM_CREATE
    fResizingTocWindow=FALSE   'Flag: We are currently not resizing the TOC list window
    '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
    hMenuMain=GetMenu(hWndMainWindow) 'Global variable
    '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
    'Create the Content-Listbox    
    hWndToc=CreateWindow("LISTBOX","",LBS_OWNERDRAWFIXED Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT Or LBS_NOTIFY Or LBS_DISABLENOSCROLL Or WS_VISIBLE Or WS_VSCROLL Or WS_HSCROLL Or WS_BORDER Or WS_CHILD,0,0,0,0,hWndMainWindow,Cast(HMENU,IDC_TOCWND),hProgInst,NULL)
    SendMessage(hWndToc,LB_SETHORIZONTALEXTENT,1024,0)     'Horizontal scrolling range (estimated only)
   'TOC list must have a Unicode font. TOC list item height depends on caption bar height, TOC text depends on caption bar button height
    hTocFont=CreateFont(dwHeightOfTocFont,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,PROOF_QUALITY,VARIABLE_PITCH Or FF_SWISS,"Arial")  'Used in TOC window for example 
    'hTocFont=CreateFont(dwHeightOfTocFont,0,0,0,FW_NORMAL,0,0,0,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,PROOF_QUALITY,VARIABLE_PITCH Or FF_SWISS,TXT_UNICODE_SYSTEMFONT)  'Used in TOC window for example 
    SendMessage(hWndToc,WM_SETFONT,Cast(WPARAM,hTocFont),False)   'Select a Unicode font
    hwndShadowTOC_Unsorted=CreateWindow("LISTBOX","",LBS_HASSTRINGS,0,0,0,0,NULL,NULL,hProgInst,NULL)           'Invisible listbox for the raw file list
    hwndShadowTOC_Sorted=CreateWindow("LISTBOX","",LBS_HASSTRINGS Or LBS_SORT,0,0,0,0,NULL,NULL,hProgInst,NULL) 'Invisible sorting listbox for the raw file list
    hBrushTocListBk=CreateSolidBrush(COLOR_TOCBK)                                                        'For the ownerdrawn TOC listbox, text and icon background
    pfncOrigTocWndProc=SetWindowLong(hWndToc,GWL_WNDPROC,Cast(UINT,@TocWndProc))                  'Subclass the TOC listbox window, so that we can examine any message it gets
    '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
    'Now create the small vertical line, which is always between the TOC window and the text window and is used to drag the window sizes:
    hWndVertLine=CreateWindow("STATIC",ByVal NULL,SS_NOPREFIX Or WS_VISIBLE Or WS_CHILD,NULL,NULL,NULL,NULL,hWndMainWindow,0,hProgInst,ByVal NULL)  'SS_GRAYRECT/SS_BLACKRECT/SS_SIMPLE/SS_NOPREFIX/SS_SUNKEN ===> SS_NOPREFIX has window background color!
    '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
    'Now create the RichEdit Text window:
    hWndText=CreateWindowEx(WS_EX_STATICEDGE,MSFTEDIT_CLASS,ByVal NULL,WS_CHILD Or WS_VISIBLE Or WS_VSCROLL Or ES_READONLY Or ES_NOHIDESEL Or ES_MULTILINE Or ES_AUTOVSCROLL,NULL,NULL,NULL,NULL,hWndMainWindow,Cast(HMENU,IDC_TXTWND),hProgInst,ByVal NULL) 
    'GetWindowLongPtr(hWndText,GWLP_WNDPROC) 'Get a pointer to the RichEdit's default Proc, so that we can call it with CallWindowProc()
    pfncOrigTxtWndProc=SetWindowLong(hWndText,GWL_WNDPROC,Cast(UINT,@TextWndProc))                  'Subclass the RichEdit window, so that we can examine any message it gets
    SendMessage(hWndText,EM_SETTEXTMODE,TM_RICHTEXT Or TM_SINGLELEVELUNDO Or TM_SINGLECODEPAGE,0)   '
    SendMessage(hWndText,EM_SETTYPOGRAPHYOPTIONS,TO_ADVANCEDTYPOGRAPHY,1)                           'Necessary to support "justify" paragraph formatting
    ZeroMemory(@tCharFormat,SizeOf(tCharFormat))
    tCharFormat.cbSize=SizeOf(tCharFormat)
    tCharFormat.dwMask=CFM_SIZE Or CFM_CHARSET Or CFM_FACE Or CFM_BOLD
    tCharFormat.yHeight=200   '1/1440 of an inch, or 1/20 of a printer's point
    tCharFormat.bCharSet=ANSI_CHARSET
    tCharFormat.bPitchAndFamily=VARIABLE_PITCH Or FF_SWISS
    tCharFormat.szFaceName=WStr("Arial")
    i=SendMessage(hWndText,EM_SETCHARFORMAT,SCF_ALL,Cast(LPARAM,@tCharFormat))     'Set default formatting to support most simple RTF file such as {\rtf1 Blabla}
    'EM_SETPARAFORMAT ???
    SendMessage(hWndText,EM_EXLIMITTEXT,0,&H3FFFFFF)                                                '&H1FFFFFF=32MB, &H3FFFFFF=64MB  (default limit: 64K characters).  
    SendMessage(hWndText,EM_SETUNDOLIMIT,0,0)                                                       'Disable the Undo feature (speeds up loading of big files)
    'SendMessage(hWndText,EM_SETEDITSTYLE,0,SES_HYPERLINKTOOLTIPS)                                  'Activate tooltips for links (not with RichEdit 4.1)
    SendMessage(hWndText,EM_SETEVENTMASK,0,ENM_LINK)                                                'We need events for Link clicks
    SendMessage(hWndText,EM_SETMARGINS,EC_LEFTMARGIN Or EC_RIGHTMARGIN,MAKELPARAM(10,10))             'Set a small right+left margin
    'SendMessage(hWndText,EM_SETMARGINS,EC_USEFONTINFO,0)             'Set a small right+left margin
    RichCom_SetComInterface(hWndText)                                                               'Init the OLE interface (otherwise no picture support)
    '
    RegisterClipboardFormat(CF_RTF)  'CF_RTF, CF_RTFNOOBJS, and CF_RETEXTOBJ
    '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
    SetFocus(hWndText)
    '
    If UserSettings_swzCurArchiveName<>"" Then
      If UserSettings_swzCurArchiveName<>UserSettings_swzLastArchiveName Then UserSettings_lLastChapterTitle=0  'Reset "Last chapter" value
      PostMessage(hWndMainWindow,WM_COMMAND,IDCV_LOADFILE,0)   'Virtual control: Load file by indirect call (allows further initialization before)
    End If
    '
    Return 0
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDC_OPENFILE         '
      swBuffer=UserSettings_swzCurArchiveName
      tOfn.lStructSize     = SizeOf(OPENFILENAME)
      tOfn.hwndOwner       = hWndMainWindow
      tOfn.lpstrFilter     = Cast(LPCWSTR,StrPtr(WStr(!"RTF Book (*.rbk)\0*.rbk\0ZIP files (*.zip)\0*.zip\0All files (*.*)\0*.*\0\0")))
      Select Case UCase(Mid(swBuffer,Len(swBuffer)-2,8))
      Case "RBK",""
        tOfn.nFilterIndex = 1
      Case "ZIP"
        tOfn.nFilterIndex = 2
      Case Else
        tOfn.nFilterIndex = 3
      End Select
      tOfn.lpstrFile       = Cast(LPTSTR,@swBuffer)
      tOfn.lpstrInitialDir = VarPtr(swBuffer)
      tOfn.nMaxFile        = MAX_PATH
      tOfn.Flags           = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY Or OFN_EXPLORER Or OFN_ENABLESIZING
      'Ofn.FlagsEx        = OFN_EX_NOPLACESBAR   'Win2000: Don't show icons for favorits, desktop and so on
      If GetOpenFileName(@tOfn)<>0 Then
        UserSettings_swzCurArchiveName=swBuffer
        OpenBook(hWndMainWindow,UserSettings_swzCurArchiveName)
      End If
    Case IDCV_LOADFILE          'Virtual control, Target for PostMessage()
      OpenBook(hWndMainWindow,UserSettings_swzCurArchiveName)
    Case IDC_SAVECHAPTER
      ExtractCurFile(hWndMainWindow)
    Case IDC_EXPORT      '
      DialogBox(hProgInst,MAKEINTRESOURCE(IDD_EXPORT),hWndMainWindow,@ExportDlgProc)
    Case IDC_PRINT            '
      'MessageBox(hWndMainWindow,"Print: Still to do...","Debug",0)
      'Set the parameters for the printing thread:
      tPrintInfos.hwndMain      = hWndMainWindow 'Handle of the main application window
      tPrintInfos.hInst         = hProgInst      'Instance handle of the application
      tPrintInfos.hwndText      = hWndText       'Handle of the RichEdit control, which's text is to print
      tPrintInfos.lLeftMargin   = 20             'Page margin in Millimeters
      tPrintInfos.lRightMargin  = 20             'Page margin in Millimeters
      tPrintInfos.lTopMargin    = 16             'Page margin in Millimeters
      tPrintInfos.lBottomMargin = 16             'Page margin in Millimeters
      CreateThread(NULL,NULL,Cast(LPTHREAD_START_ROUTINE,@PrintRichText),@tPrintInfos,NULL,@i) 'Printing we do in an own thread
    Case IDC_LANG_AUTO            '
      UserSettings.fLanguageAutomatic=True
      UserSettings.wLanguageID=LoByte(GetUserDefaultUILanguage())   'Get current user's system language, &H00=Neutral (english), &H07=German, &H03=Catalan and so on 
      hOldMenu=GetMenu(hWndMainWindow)
      Select Case UserSettings.wLanguageID
      Case &H07 : hMenuMain=LoadMenu(hProgInst,MAKEINTRESOURCE(IDM_MAIN_GERMAN))    'German
      Case Else : hMenuMain=LoadMenu(hProgInst,MAKEINTRESOURCE(IDM_MAIN))           '&H00=Neutral (english)
      End Select
      hMenuMain=LoadMenu(hProgInst,MAKEINTRESOURCE(IDM_MAIN_GERMAN))
      DestroyMenu(hOldMenu)
      UpdateMainMenu(hWndMainWindow)
    Case IDC_LANG_GERMAN
      UserSettings.fLanguageAutomatic=False
      UserSettings.wLanguageID=&H07
      hOldMenu=GetMenu(hWndMainWindow)
      hMenuMain=LoadMenu(hProgInst,MAKEINTRESOURCE(IDM_MAIN_GERMAN))
      SetMenu(hWndMainWindow,hMenuMain)
      DestroyMenu(hOldMenu)
      UpdateMainMenu(hWndMainWindow)
    Case IDC_LANG_ENGLISH
      UserSettings.fLanguageAutomatic=False
      UserSettings.wLanguageID=&H00
      hOldMenu=GetMenu(hWndMainWindow)
      hMenuMain=LoadMenu(hProgInst,MAKEINTRESOURCE(IDM_MAIN))
      SetMenu(hWndMainWindow,hMenuMain)
      DestroyMenu(hOldMenu)
      UpdateMainMenu(hWndMainWindow)
    Case IDC_SETTINGS_SAVE
      UserSettings.fStoreSettings=True
    Case IDC_SETTINGS_NOSAVE
      UserSettings.fStoreSettings=False
    Case IDC_EXIT             '
      'MessageBox(hWndMainWindow,"Really Exit? etc.","",0)
      PostQuitMessage(0)
    Case IDC_COPY             '
      SendMessage(hWndText,WM_COPY,0,0)
    Case IDC_SELECT_ALL
      tCharRange.cpMin=0
      tCharRange.cpMax=-1
      SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tCharRange))
    Case IDC_SEARCH           '
      hDlgSearch=CreateDialog(hProgInst,MAKEINTRESOURCE(IDD_SEARCHDLG),hWndMainWindow,@SearchDlgProc)                       'Open non-modal "Search" window
    Case IDC_ZOOMSMALLER       '
      UserSettings.TextZoom-=1
      If UserSettings.TextZoom<=ZOOM_MINIMAL Then
        UserSettings.TextZoom=ZOOM_MINIMAL
      Else
        SendMessage(hWndText,EM_SETZOOM,UserSettings.TextZoom,ZOOM_NORMAL)
      End If
      UpdateMainMenu(hWndMainWindow)
    Case IDC_QUICKSEARCH
      Select Case UserSettings.wLanguageID
      Case &H07
        swBuffer="Die Schnellsuche arbeitet ohne Eingabefenster, und zwar entweder im Inhaltsverzeichnis oder innerhalb des aktuellen Texts,"+_
                 " abhängig davon, welches Fenster gerade aktiv ist. Die Schnellsuche startet, sobald man Zeichen in die Tastatur eingibt. "+_
                 "Längere Suchbegriffe bildet man durch Zeitabstände kleiner als einer Sekunde zwischen den Tastendrücken. Mit "+WChr(&H0022)+_
                 "Schnellsuche fortsetzen"+WChr(&H0022)+" (Kurztaste F3) wird nach dem letzten Suchbegriff gesucht."+WChr(13)+_
                 "Beide Fenster arbeiten mit je einem eigenen Suchbegriff."+WChr(13)+_
                 "Die Schnellsuche beachtet stets die Groß-/Kleinschreibung."
        MessageBox(hWndMainWindow,swBuffer,"Hinweis zur Schnellsuche",MB_ICONINFORMATION)
      Case Else
        swBuffer="Quick-search works without an input window, either in the table of contents or within the current text, depending on which "+_
                 "window is currently active. Quick-search starts whenever characters are entered into the keyboard. Longer search terms are "+_
                 "formed by time intervals of less than one second between the keystrokes. Press "+WChr(&H0022)+"Continue quick search"+WChr(&H0022)+_
                 " (shortcut F3) to search for the last term."+WChr(13)+"Both windows work with their own search term."+WChr(13)+_
                 "Quick-search is always case-sensitive."
        MessageBox(hWndMainWindow,swBuffer,"Hints for quick search",MB_ICONINFORMATION)
      End Select
    Case IDC_QUICKSEARCH_AGAIN                  'Continue quick search
      If GetFocus()=hWndTOC Then
        PostMessage(hWndToc,WM_CHAR,NULL,0)     'More see subclassing function of TOC listbox TocWndProc()
      Else                                      'Any other window than TOC has the focus  ---> otherwise: ElseIf GetFocus()=hWndText Then
        PostMessage(hWndText,WM_CHAR,NULL,0)    'More see subclassing function of RichEdit text window TextWndProc()
      End If
    Case IDC_ZOOMNORMAL       '
      UserSettings.TextZoom=ZOOM_NORMAL
      SendMessage(hWndText,EM_SETZOOM,UserSettings.TextZoom,ZOOM_NORMAL)
      UpdateMainMenu(hWndMainWindow)
    Case IDC_ZOOMBIGGER       '
      UserSettings.TextZoom+=1
      If UserSettings.TextZoom>=ZOOM_MAXIMAL Then     'Just to have any end of zoom
        UserSettings.TextZoom=ZOOM_MAXIMAL
      Else
        SendMessage(hWndText,EM_SETZOOM,UserSettings.TextZoom,ZOOM_NORMAL)
      End If
      UpdateMainMenu(hWndMainWindow)
    Case IDC_CHAPTERUP        '
      i=SendMessage(hWndToc,LB_GETCURSEL,0,0)
      If i>0 Then LoadBookChapter(SendMessage(hWndToc,LB_SETCURSEL,i-1,0))  'ZERO based index!
      UpdateMainMenu(hWndMainWindow)
    Case IDC_CHAPTERDOWN      '
      i=SendMessage(hWndToc,LB_GETCURSEL,0,0)
      If i<SendMessage(hWndToc,LB_GETCOUNT,0,0)-1 Then   'ZERO based index!
        LoadBookChapter(SendMessage(hWndToc,LB_SETCURSEL,i+1,0))
      End If
      UpdateMainMenu(hWndMainWindow)
    Case IDC_SWITCH_WIN
      If GetFocus()=hWndTOC Then
        SetFocus(hWndText)
      ElseIf GetFocus()=hWndText Then
        SetFocus(hWndTOC)
      End If
    Case IDC_BIGGER_TOCWIN    '
      ResizeChildWindows(hWndMainWindow,UserSettings.TocWinWidth+10)
    Case IDC_SMALLER_TOCWIN   '
      ResizeChildWindows(hWndMainWindow,UserSettings.TocWinWidth-10)
    Case IDC_HELPMAIN         '
      hDlgHelp=CreateDialog(hProgInst,MAKEINTRESOURCE(IDD_HELPDLG),hWndMainWindow,@HelpDlgProc)                       'Open non-modal "Help" window
   ' Case IDC_TXTWND            ' 
    Case IDC_TOCWND            'Table of Contents listbox
      If HiWord(wParam)=LBN_SELCHANGE Then LoadBookChapter(SendMessage(hWndToc,LB_GETCURSEL,0,0))
    End Select
    Return 0
  Case WM_CTLCOLORLISTBOX
   If lParam=hwndTOC Then
     SetTextColor(Cast(HDC,WPARAM),COLOR_TOCTEXT)                  'Set text color black
     SetBkColor(Cast(HDC,WPARAM),COLOR_TOCBK)                      'Set text background color
     Return Cast(LRESULT,hBrushTocListBk)
   Else
     Return 0
   End If
  Case WM_NOTIFY
    pNmHdr=Cast(NMHDR Ptr,lParam)
    If pNmHdr->idFrom=IDC_TXTWND And pNmHdr->code=EN_LINK Then                     'From hyperlink in Richedit text window.
      pEnLink=Cast(ENLINK Ptr,lParam)                                              'In this case the NMHDR type is enhanced to an ENLINK type
      If pEnLink->Msg=WM_LBUTTONDOWN Then
        tTextRange.chrg=pEnLink->chrg
        tTextRange.lpstrText=@swBuffer
        SendMessage(hWndText,EM_GETTEXTRANGE,0,Cast(LPARAM,@tTextRange))
        If swBuffer=TXT_LINKID_BAGGAGE Then                                        '"[BAGGAGE]" baggage/attachement file
          ExtractCurFile(hWndMainWindow)                                           'All information we get from TOC entry currently selected.
        'Else
          'Reserved for later ideas!
        End If
      End If
    End If
    Function=0
  Case WM_MOUSEMOVE 'This message we get only, when the mouse is over the "free slot" between TOC and text window. Everything else in the main window is completely covered.
    SetCursor(LoadCursor(NULL,ByVal IDC_SIZEWE))  'Display a left-right cursor
    Function=0
  Case WM_LBUTTONDOWN                                                              'Resize the TOC list window (and with it the text list window): Switch to "resizing mode"
    fResizingTocWindow=TRUE                                                        'Flag: We are currently resizing the TOC list window
    InvalidateRect(hWndMainWindow,ByVal NULL,TRUE)                                 'Erase entire client background aof main window. This prevents ugly redrawing artefacts while dragging the vertival line.
    Function=0
  Case WM_LBUTTONUP                                                                'End resizing the TOC list window (and with it the text list window): quit the "resizing mode"
    If fResizingTocWindow=TRUE Then
      fResizingTocWindow=FALSE                                                     'Flag: We are currently not resizing the TOC list window
      InvalidateRect(hWndMainWindow,ByVal NULL,TRUE)                               'Refresh (causes refresh of all child windows too)
      Function=0
    End If
  Case WM_SETCURSOR
    If fResizingTocWindow=TRUE Then                                                'Flag: We are currently resizing the TOC list window
      GetCursorPos(@pt)                                                            'Get current mouse position in screen coordinates
      ScreenToClient(hWndMainWindow,@pt)                                           'Convert to client coordinates
      If UserSettings.TocWinWidth<>pt.x Then
        ResizeChildWindows(hWndMainWindow,pt.x)
        Function=0
      Else
        Return DefWindowProc(hWndMainWindow,uMsg,wParam,lParam)
      End If
    Else
      Return DefWindowProc(hWndMainWindow,uMsg,wParam,lParam)
    End If
  Case WM_INITMENUPOPUP '
    If HiWord(LParam)=False Then 'From main Window menu
      Select Case LoWord(LParam)                                                          '0-based position in main menu
      Case 0                                                                              '0: Popup menu "File"
        If pCurZipArchive=NULL Then                                                       'LibZip's ZIP pointer to the ZIP archive: NULL: No ZIP book is open.
          EnableMenuItem(Cast(HMENU,wParam),IDC_EXPORT,MF_BYCOMMAND Or MF_GRAYED)
        Else
          EnableMenuItem(Cast(HMENU,wParam),IDC_EXPORT,MF_BYCOMMAND Or MF_ENABLED)
        End If
      Case 6                                                                               '6: Popup menu "File"--->"Settings"
        If UserSettings.fStoreSettings=False Then
          CheckMenuRadioItem(Cast(HMENU,wParam),IDC_SETTINGS_SAVE,IDC_SETTINGS_NOSAVE,IDC_SETTINGS_NOSAVE,MF_BYCOMMAND)
        Else
          CheckMenuRadioItem(Cast(HMENU,wParam),IDC_SETTINGS_SAVE,IDC_SETTINGS_NOSAVE,IDC_SETTINGS_SAVE,MF_BYCOMMAND)
        End If
        If UserSettings.fLanguageAutomatic<>False Then                                     'Grrrrr, "If UserSettings.fLanguageAutomatic=True" doesn't work propper.
          CheckMenuRadioItem(Cast(HMENU,wParam),IDC_LANG_AUTO,IDC_LANG_LAST,IDC_LANG_AUTO,MF_BYCOMMAND)
        Else
          If UserSettings.wLanguageID=0 Then
            CheckMenuRadioItem(Cast(HMENU,wParam),IDC_LANG_AUTO,IDC_LANG_LAST,IDC_LANG_ENGLISH,MF_BYCOMMAND)
          Else
            CheckMenuRadioItem(Cast(HMENU,wParam),IDC_LANG_AUTO,IDC_LANG_LAST,IDC_LANG_GERMAN,MF_BYCOMMAND)
          End If
        End If
      End Select
    End If
    Return False
  Case WM_MEASUREITEM  'Ownerdraw listbox, menu and so on measurement information '- - - - - - - - - - - - - - - - - - - - - - -
    Dim lpMeasureItem As MEASUREITEMSTRUCT Ptr=Cast(MEASUREITEMSTRUCT Ptr,lParam)
    If wParam=0 And lpMeasureItem->CtlType=ODT_MENU Then
      lpMeasureItem->itemWidth  = GetSystemMetrics(SM_CYMENU)\2
      lpMeasureItem->itemHeight = GetSystemMetrics(SM_CYMENU)
    ElseIf wParam=IDC_TOCWND And lpMeasureItem->CtlType=ODT_LISTBOX Then
      lpMeasureItem->itemHeight=dwHeightOfTocFont
    End If
    Return True
  Case WM_DRAWITEM 'Draw Ownerdrawn listboxes  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Return OnDrawItems(hWndMainWindow,ByVal Cast(DRAWITEMSTRUCT Ptr,lParam))
  Case WM_SIZE                                                         'LoWord(lParam)=width, hiWord(lParam)=height
    ResizeChildWindows(hWndMainWindow,UserSettings.TocWinWidth)
    Return 0
  Case WM_SETFOCUS
    SetFocus hWndText
    Return 0
  Case WM_DESTROY
    If hWndMainWindow=hWndMain Then  'Main app window closed (hWndMain is the global variable, which holds the handle local in hWndMainWindow)
      DeleteObject(hTocFont)
      'WinHelp hWnd,ByCopy szPBHelpFile,HELP_QUIT,0  'Close possibly opened helpfile
      PostQuitMessage 0
      Return 0
    Else                     'Child window closed
      Return DefWindowProc(hWndMainWindow,uMsg,wParam,lParam)
    End If
  Case WM_CLOSE
    'MessageBox(hWnd,"Do you really want?","Debug",MB_ICONQUESTION)
    '
    wplc.length=SizeOf(WINDOWPLACEMENT) 
    GetWindowPlacement(hWndMainWindow,@wplc)
    If wplc.showCmd<>SW_SHOWMAXIMIZED Then wplc.showCmd=SW_SHOWNORMAL
    UserSettings.WinState=wplc.showCmd                                                         'State of the main window: Only SW_SHOWMAXIMIZED or SW_SHOWNORMAL is allowed
    '
    GetWindowRect(hWndMainWindow,@rct)                                                         'Get Window rectangle to store the values in ini file
    UserSettings.WinX=rct.left
    UserSettings.WinY=rct.top
    UserSettings.WinWidth=rct.right-rct.left
    UserSettings.WinHeight=rct.bottom-rct.top
    '
    UserSettings_swzLastArchiveName=UserSettings_swzCurArchiveName
    UserSettings_lLastChapterTitle=SendMessage(hwndTOC,LB_GETCURSEL,0,0)
    '
    GetClientRect(hWndText,@rct)                                                               'Get Window rectangle to store the values in ini file
    pt.x=0                                                                                     'Left of the RichEdit's client area
    pt.y=rct.bottom-rct.top-32                                                                 'Bottom (-32 is a tested-out value, which seems to be perfect)
    UserSettings_dwLastChapterPos=SendMessage(hWndText,EM_CHARFROMPOS,0,Cast(lParam,@pt))      'Get position of the character in the most down left corner in the window
    '
    SetWindowLong(hWndText,GWL_WNDPROC,pfncOrigTxtWndProc)                                     'Set original RichEdit window procedure (we subclassed the control)
    SetWindowLong(hWndToc,GWL_WNDPROC,pfncOrigTocWndProc)                                      'Set original TOC listbox window procedure (we subclassed the control)
    DestroyWindow(hWndMainWindow)
    Function=DefWindowProc(hWndMainWindow,uMsg,wParam,lParam)                                  'Close the window
  Case Else
    Return DefWindowProc(hWndMainWindow,uMsg,wParam,lParam)
  End Select
End Function
'************************************************************************************************************************
Sub UpdateMainMenu(ByVal hWnd As HWND)
  'Updates the main menu's OwnerDraw elements depending on the current settings
  'hWnd: Handle of the main window, hMenu: Handle of the main window's manu bar.
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static hMenu As HMENU,hMenuWindow As HMENU,hMenuSettings As HMENU,i As Integer,j As Integer,fRedraw As Integer
  hMenu=GetMenu(hWnd)
  hMenuWindow=GetSubMenu(hMenu,2)      'Can change, if the user loads another maoin menu by changing the language
  hMenuSettings=GetSubMenu(GetSubMenu(hMenu,0),5)    'Can change, if the user loads another maoin menu by changing the language
  fRedraw=False
  '
  If UserSettings.TextZoom<=ZOOM_MINIMAL Then
    If EnableMenuItem(hMenu,IDC_ZOOMSMALLER,MF_BYCOMMAND Or MF_GRAYED)<>MF_GRAYED Then fRedraw=True
    If EnableMenuItem(hMenu,IDC_ZOOMBIGGER,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    EnableMenuItem(hMenuWindow,IDC_ZOOMSMALLER,MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem(hMenuWindow,IDC_ZOOMBIGGER,MF_BYCOMMAND Or MF_ENABLED)
  ElseIf UserSettings.TextZoom>=ZOOM_MAXIMAL Then
    If EnableMenuItem(hMenu,IDC_ZOOMSMALLER,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    If EnableMenuItem(hMenu,IDC_ZOOMBIGGER,MF_BYCOMMAND Or MF_GRAYED)<>MF_GRAYED Then fRedraw=True
    EnableMenuItem(hMenuWindow,IDC_ZOOMSMALLER,MF_BYCOMMAND Or MF_ENABLED)
    EnableMenuItem(hMenuWindow,IDC_ZOOMBIGGER,MF_BYCOMMAND Or MF_GRAYED)
  Else
    If EnableMenuItem(hMenu,IDC_ZOOMSMALLER,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    If EnableMenuItem(hMenu,IDC_ZOOMBIGGER,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    EnableMenuItem(hMenuWindow,IDC_ZOOMSMALLER,MF_BYCOMMAND Or MF_ENABLED)
    EnableMenuItem(hMenuWindow,IDC_ZOOMBIGGER,MF_BYCOMMAND Or MF_ENABLED)
  End If
  '
  i=SendMessage(hWndToc,LB_GETCURSEL,0,0)
  j=SendMessage(hWndToc,LB_GETCOUNT,0,0)-1  'We work with ZERO based index
  If i<=0 Then
    If EnableMenuItem(hMenu,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_GRAYED)<>MF_GRAYED Then fRedraw=True
    If EnableMenuItem(hMenu,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    EnableMenuItem(hMenuWindow,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem(hMenuWindow,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_ENABLED)
  ElseIf i>=j Then
    If EnableMenuItem(hMenu,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_GRAYED)<>MF_GRAYED Then fRedraw=True
    If EnableMenuItem(hMenu,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    EnableMenuItem(hMenuWindow,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_GRAYED)
    EnableMenuItem(hMenuWindow,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_ENABLED)
  Else
    If EnableMenuItem(hMenu,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    If EnableMenuItem(hMenu,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_ENABLED)<>MF_ENABLED Then fRedraw=True
    EnableMenuItem(hMenuWindow,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_ENABLED)
    EnableMenuItem(hMenuWindow,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_ENABLED)
  End If
  '
  'Done dynnamicly after WM_INITMENUPOPUP:
  '  If UserSettings.fLanguageAutomatic=True Then 'User language Ianguage will be choosed automaticly by system language
  '    CheckMenuRadioItem(hMenuSettings,IDC_LANG_AUTO,IDC_LANG_LAST,IDC_LANG_AUTO,MF_BYCOMMAND)
  '  ElseIf UserSettings.wLanguageID=&H07 Then                      '&H07=German
  '    CheckMenuRadioItem(hMenuSettings,IDC_LANG_GERMAN,IDC_LANG_LAST,IDC_LANG_GERMAN,MF_BYCOMMAND)
  '  Else                                                           '&H00 Neutral (english)
  '    CheckMenuRadioItem(hMenuSettings,IDC_LANG_ENGLISH,IDC_LANG_LAST,IDC_LANG_ENGLISH,MF_BYCOMMAND)
  '  End If
  '
  'Done dynnamicly after WM_INITMENUPOPUP:
  '  If UserSettings.fStoreSettings=True Then
  '    CheckMenuRadioItem(hMenuSettings,IDC_SETTINGS_SAVE,IDC_SETTINGS_NOSAVE,IDC_SETTINGS_SAVE,MF_BYCOMMAND)
  '  Else
  '    CheckMenuRadioItem(hMenuSettings,IDC_SETTINGS_SAVE,IDC_SETTINGS_NOSAVE,IDC_SETTINGS_NOSAVE,MF_BYCOMMAND)
  '  End If
  '
  If fRedraw=True Then DrawMenuBar(hWnd)
End Sub
'************************************************************************************************************************
Sub ResizeChildWindows(ByVal hWndParent As HWND,ByVal lTocWinWidth As Long)
  'Changes width(s) of TOC window and text window depending on each other. Hides TOC window if one-file-archive loaded.
  'Used global variables: hWndToc, hWndVertLine, hWndText, hMenuMain, Type UserSettings
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static rct As RECT
  GetClientRect(hWndParent,@rct)                                                       'Get the client width and height
  If SendMessage(hWndToc,LB_GETCOUNT,0,0)=1 Then                                       'Single chapter mode: Don't show the TOC window
    EnableWindow(hWndToc,FALSE)
    EnableWindow(hWndVertLine,FALSE)
    MoveWindow(hWndText,0,0,rct.right,rct.bottom,True)
    SetFocus(hWndText)
    '
    EnableMenuItem(hMenuMain,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_GRAYED)                  'Disable ChapterUp menu button (has no function for the moment)
    EnableMenuItem(hMenuMain,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_GRAYED)                'Disable ChapterDown menu button (has no function for the moment)
  Else                                                                                 'Multi chapter mode: Show the TOC window with dragable width
    If lTocWinWidth<MINIMAL_TOCWNDWIDTH Then lTocWinWidth=MINIMAL_TOCWNDWIDTH          'Prevent (technically) unusual size  [was: pt.x=Max(pt.x,MINIMAL_TOCWNDWIDTH)]
    UserSettings.TocWinWidth=lTocWinWidth                                              'Store for resize of entire window and saving into ini file
    EnableWindow(hWndToc,TRUE)
    EnableWindow(hWndVertLine,TRUE)
    MoveWindow(hWndToc,0,1,UserSettings.TocWinWidth,rct.bottom-1,TRUE)
    MoveWindow(hWndVertLine,UserSettings.TocWinWidth+1,1,WIDTH_VERT_LINE,rct.bottom-1,TRUE)
    MoveWindow(hWndText,UserSettings.TocWinWidth+WIDTH_VERT_LINE+1,1,rct.right-lTocWinWidth-WIDTH_VERT_LINE,rct.bottom,TRUE)  'Never forget: Let 4 pixels space between TOC and text window. The free "slot" is important for resizing function.
    '
    EnableMenuItem(hMenuMain,IDC_CHAPTERUP,MF_BYCOMMAND Or MF_ENABLED)                 'Enable ChapterUp menu button if disabled
    EnableMenuItem(hMenuMain,IDC_CHAPTERDOWN,MF_BYCOMMAND Or MF_ENABLED)               'Enable ChapterUp menu button if disabled
  End If
  DrawMenuBar(hWndParent)
End Sub
'************************************************************************************************************************
Function OnDrawItems(ByVal hWnd As HWND,ByVal lpdis As DRAWITEMSTRUCT Ptr) As LRESULT
  'Ownerdraw actions in main window after WM_DRAWITEM
   '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static hIconZoomSmaller As HICON,hIconZoomSmaller2 As HICON,hIconZoomNormal As HICON,hIconZoomNormal2 As HICON,hIconZoomBigger As HICON,hIconZoomBigger2 As HICON
  Static hIconDown As HICON,hIconDown2 As HICON,hIconUp As HICON,hIconUp2 As HICON,swTitle As wString*MAX_PATH,hBrushMenuBkHighLighted As HBRUSH,i As Long
  Static hIconDocument As HICON,hIconPicture As HICON,hIconInfo As HICON,hIconAttachement As HICON,iIndentionLevel As Integer,hOriginalBrush As HBRUSH,hDC As HDC
  Static iItemHeight As Integer,hBrushMenuBk As HBRUSH,iRetVal As Integer,hCurIcon As HICON,hCurIcon2 As HICON,fDrawMenuIcon As Integer,iTextLen As Integer
  Static tTextMetric As TEXTMETRIC,hMenuFont As HFONT,hOldFont As HFONT,swBuffer As Wstring*255,tRect As RECT,iMargin As Integer
  Static fFlatMenuUsed As Integer,hOldBrush As HBRUSH,hPen As HPEN,hOldPen As HPEN,tPoint(3) As Point
  Static iTopPos As Integer,iVertCenter As Integer,iHorzCenter As Integer,iBottomPos As Integer,iRightPos As Integer

  'Dim lpdis As DRAWITEMSTRUCT Ptr=Cast(DRAWITEMSTRUCT Ptr,lParam)
  
  If hIconDocument=NULL Then     'First function call: Make intializations
    SystemParametersInfo(SPI_GETFLATMENU,NULL,@fFlatMenuUsed,NULL)             'Are WindowsXP flat menus used?
    '
    hIconDocument=LoadIcon(hProgInst,Cast(LPCTSTR,IDRI_DOCUMENT_ICO))          'For the ownerdrawn TOC listbox
    hIconPicture=LoadIcon(hProgInst,Cast(LPCTSTR,IDRI_PICTURE_ICO))            'For the ownerdrawn TOC listbox
    hIconInfo=LoadIcon(hProgInst,Cast(LPCTSTR,IDRI_INFO_ICO))                  'For the ownerdrawn TOC listbox
    hIconAttachement=LoadIcon(hProgInst,Cast(LPCTSTR,IDRI_ATTACH_ICO))         'For the ownerdrawn TOC listbox
    '
    GetTextMetrics(lpdis->hDC,@tTextMetric)                                    'Get data of the default font, which might be an ANSI font without extended Unicode characters
    hPen=CreatePen(PS_SOLID,Max(1,tTextMetric.tmHeight\10),RGB(0,0,0))
  End If
  '
  If lpdis->itemID = &HFFFFFFFF& Then
    iRetVal=False
    Goto DoneOnDrawItems
  End If
  '
  iItemHeight           = lpdis->rcItem.bottom-lpdis->rcItem.top
  iMargin               = Max(1,iItemHeight\7)
  iTopPos               = lpdis->rcItem.top+1
  iVertCenter           = iTopPos+tTextMetric.tmHeight\2
  iHorzCenter           = lpdis->rcItem.left+tTextMetric.tmHeight\2
  iBottomPos            = iTopPos+tTextMetric.tmHeight
  iRightPos             = lpdis->rcItem.left+tTextMetric.tmHeight
  '
  If lpdis->CtlID=IDC_TOCWND Then      'TOC listbox
    hCurIcon  = hIconDocument
    hCurIcon2 = hIconDocument
    SendMessage(lpdis->hwndItem,LB_GETTEXT,lpdis->itemID,Cast(LPARAM,@swTitle))  'Item text: very first character is a 1-digit number giving the eindention level
    iIndentionLevel=ValUInt(Str(Left(swTitle,1))) * iItemHeight                  '1st char: Indention level: One indention unit is one icon width (tested, looks fine)
    '
    Select Case *Cast(UShort Ptr,StrPtr(swTitle)+2)                              'Second digit in file title prefix: file type
    'Case TOC_PRE2_TEXT    : xxxxx                                               'Invisible! Text. Ignored, invisible, reserved for metainfo files
    Case TOC_PRE2_METAINFO : hCurIcon=hIconInfo
    Case TOC_PRE2_RTF      : hCurIcon=hIconDocument
    Case TOC_PRE2_PICTURE  : hCurIcon=hIconPicture
    Case Else              : hCurIcon=hIconAttachement                           'TOC_PRE2_ATTACHEMENT
    End Select
    '
    DrawIconEx(lpdis->hDC,iIndentionLevel+lpdis->rcItem.left,(lpdis->rcItem.top+(lpdis->rcItem.bottom-lpdis->rcItem.top-iItemHeight)\2)+2,hCurIcon,iItemHeight,iItemHeight,NULL,hBrushTocListBk,DI_NORMAL)
    If (lpdis->itemState And ODS_SELECTED) Then
      SetBkColor(lpdis->hDC,COLOR_TOCTEXT)
      If GetFocus()=lpdis->hwndItem Then
        SetTextColor(lpdis->hDC,&H00FFFFFF)
      Else
        SetBkColor(lpdis->hDC,COLOR_TOCGRAY)
      End If
    Else
      SetTextColor(lpdis->hDC,COLOR_TOCTEXT)
      SetBkColor(lpdis->hDC,COLOR_TOCBK)
    End If
    TextOut(lpdis->hDC,iIndentionLevel+iItemHeight,lpdis->rcItem.top+2,@swTitle+2,Len(swTitle)-2) 'Draw text without first two characters (which are internal info)
    iRetVal=True
  ElseIf lpdis->CtlType=ODT_MENU Then    'Must be an ownerdraw menu item of the main menu
    SystemParametersInfo(SPI_GETFLATMENU,NULL,@fFlatMenuUsed,NULL)             'Are WindowsXP flat menus used?
    If (lpdis->itemState And ODS_SELECTED) Then
      If lpdis->hwndItem=hMenuMain Then
        SetTextColor(lpdis->hDC,GetSysColor(COLOR_MENUBAR))
      Else
        SetTextColor(lpdis->hDC,GetSysColor(COLOR_MENU))
      End iF
      SetBkColor(lpdis->hDC,GetSysColor(COLOR_HIGHLIGHT))
      hOldBrush=SelectObject(lpdis->hDC,GetStockObject(LTGRAY_BRUSH))
    Else
      If fFlatMenuUsed Then
        If lpdis->hwndItem=hMenuMain Then
          SetBkColor(lpdis->hDC,GetSysColor(COLOR_MENUBAR))
        Else
          SetBkColor(lpdis->hDC,GetSysColor(COLOR_MENU))
        End If
      Else
        SetBkColor(lpdis->hDC,GetSysColor(COLOR_MENU))
      End If
      SetTextColor(lpdis->hDC,GetSysColor(COLOR_MENUTEXT))
      hOldBrush=SelectObject(lpdis->hDC,GetStockObject(WHITE_BRUSH))
    End If
    '
    hOldPen=SelectObject(lpdis->hDC,hPen)
    '
    'For ownerdraw menu entries unfortunately we can't set/retrieve the text just set in resource. So we must do it here:
    Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07                                          'German: Default language in resources
      Select Case lpdis->itemID
      Case IDC_ZOOMSMALLER: swBuffer=wChr(&H25CB)+"Text &kleiner zoomen"+wChr(9)+"F10"
      Case IDC_ZOOMNORMAL:  swBuffer=wChr(&H25C9)+"Text &normal zoomen"+wChr(9)+"F11"
      Case IDC_ZOOMBIGGER:  swBuffer=wChr(&H25CF)+"Text &größer zoomen"+wChr(9)+"F12"
      Case IDC_CHAPTERDOWN: swBuffer=wChr(&H25BC)+"Kapitel &runter "+wChr(9)+"Strg+BildAb"
      Case IDC_CHAPTERUP:   swBuffer=wChr(&H25B2)+"Kapitel &hoch "+wChr(9)+"Strg+BildAuf"
      End Select
    Case Else                                          'Neutral (english)
      Select Case lpdis->itemID
      Case IDC_ZOOMSMALLER: swBuffer=wChr(&H25CB)+"Text zoom &smaller"+wChr(9)+"F12"
      Case IDC_ZOOMNORMAL:  swBuffer=wChr(&H25C9)+"Text zoom &normal"+wChr(9)+"F11"
      Case IDC_ZOOMBIGGER:  swBuffer=wChr(&H25CF)+"Text zoom &bigger "+wChr(9)+"F13"
      Case IDC_CHAPTERDOWN: swBuffer=wChr(&H25BC)+"Chapter &up"+wChr(9)+"Ctrl+PageUp"
      Case IDC_CHAPTERUP:   swBuffer=wChr(&H25B2)+"Chapter &down"+wChr(9)+"Ctrl+PageDown"
      End Select
    End Select
    If lpdis->hwndItem=hMenuMain Then     'In main menu we schow only the "icon" character of the menu entry
      iTextLen=1
    Else
      iTextLen=Len(swBuffer)
    End If
    '
    Select Case lpdis->itemID
    Case IDC_ZOOMSMALLER                                                            'Draw a circle with a "Minus"
      Ellipse( lpdis->hDC,lpdis->rcItem.left,iTopPos,iRightPos,iBottomPos)          'circle
      MoveToEx(lpdis->hDC,lpdis->rcItem.left+iMargin,iVertCenter,NULL)              'horicontal stroke
      LineTo(  lpdis->hDC,iRightPos-iMargin-2,iVertCenter)                          'horicontal stroke
    Case IDC_ZOOMNORMAL                                                             'Draw a circle with a small circle in the center
      Ellipse( lpdis->hDC,lpdis->rcItem.left,iTopPos,iRightPos,iBottomPos)          'circle
      Ellipse( lpdis->hDC,iHorzCenter-1,iVertCenter-1,iHorzCenter+2,iVertCenter+2)  'circle
    Case IDC_ZOOMBIGGER                                                             'Draw a circle with a "Plus"
      Ellipse( lpdis->hDC,lpdis->rcItem.left,iTopPos,iRightPos,iBottomPos)          'circle
      MoveToEx(lpdis->hDC,lpdis->rcItem.left+iMargin,iVertCenter,NULL)              'horicontal stroke
      LineTo(  lpdis->hDC,iRightPos-iMargin-2,iVertCenter)                          'horicontal stroke
      MoveToEx(lpdis->hDC,iHorzCenter,iTopPos+iMargin+1,NULL)                       'vertical stroke
      LineTo(  lpdis->hDC,iHorzCenter,iBottomPos-iMargin-2)                         'vertical stroke
    Case IDC_CHAPTERDOWN
      tPoint(0).x=lpdis->rcItem.left    : tPoint(0).y=iTopPos
      tPoint(1).x=iRightPos             : tPoint(1).y=iTopPos
      tPoint(2).x=iHorzCenter           : tPoint(2).y=iBottomPos
      Polygon(lpdis->hDC,@tPoint(0),3)
    Case IDC_CHAPTERUP
      tPoint(1).x=iRightPos             : tPoint(1).y=iBottomPos
      tPoint(0).x=iHorzCenter           : tPoint(2).y=iBottomPos
      tPoint(2).x=lpdis->rcItem.left    : tPoint(0).y=iTopPos
      Polygon(lpdis->hDC,@tPoint(0),3)
    End Select
    '
    If iTextLen>1 Then                                                              'Output the other letters
      tRect.left=lpdis->rcItem.left+iItemHeight
      tRect.top=iTopPos
      tRect.right=lpdis->rcItem.right
      tRect.bottom=lpdis->rcItem.bottom
      DrawText(lpdis->hDC,@swBuffer+1,iTextLen-1,@tRect,DT_LEFT Or DT_TOP Or DT_SINGLELINE Or DT_TABSTOP Or DT_EXPANDTABS Or &H900)      'DT_TABSTOP Tab width: &H0100,&H0200,&H0300...&HFF00
    End If
    '
    SelectObject(lpdis->hDC,hOldPen)
    SelectObject(lpdis->hDC,hOldBrush)
    '
    iRetVal=True
  Else                    'Control type
    iRetVal=False
  End If
  '......................................................................................................................
  DoneOnDrawItems:
  Function=iRetVal
End Function
'************************************************************************************************************************
Function OpenBook(ByVal hWnd As HWND,ByVal sBookFile As String) As LRESULT
  'Opens the ZIP file of a book, fills the TOC listbox and sets the current chapters into the text window
  'see http://www.nih.at/libzip/
  'see http://libzip.org/documentation/libzip.html
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'Used global variables:
  'pCurZipArchive As Any Ptr                     'LibZip's ZIP pointer to the ZIP archive currently open or NULL if no ZIP book is open.
  'hWndToc As HWND                               'Listbox with the Table of Content tree
  'hwndShadowTOC As HWND                         'Invisible listbox for the raw file list, (raw Table of Content)
  'hwndText As HWND                              'Richedit window used, if a text is to disply   (hidden if currently a picture is to display)
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static iRetVal As Integer,i As Integer,j As Integer,k As Integer,iNumArchiveFiles As Integer,sBuffer As String,hTempList As HWND
  Static pszCurFileNameUTF8 As Const zString Ptr,pswCurFileName As Const Wstring Ptr,swCurFileNameWide As wString*MAX_PATH,idxZipIndexFile As Integer
  Static szCurFileNameUTF8 As String*MAX_PATH,swCurFileName As Wstring*MAX_PATH,swFileType As Wstring*4,swIndention As Wstring*4
  Static sFileExtension As String,hFile As HFILE,swzChapterTitle As Wstring*MAX_PATH,szIndexText As zString*4096,swBuffer As Wstring*MAX_PATH
  Static iPublicZipFileNo As Integer,iZipFileIndex As Integer
  '
  iRetVal=0                                             'Initial function return value: Success
  '
  SendMessage(hWndText,WM_SETTEXT,0,0)                  'Empty text window
  SendMessage(hWndToc,LB_RESETCONTENT,0,0)              'Empty TOC listbox
  '                                                     'Empty TOC shadow listbox: Later in this function
  '
  If pCurZipArchive Then                                'There is a ZIP file open currently (even if same filename because content could be changed)
    'zip_error_clear(pCurZipArchive)
    'zip_discard(pCurZipArchive)                        'We have nothing to write into the ZIP file, therefore zip_close(pCurZipArchive) isn't necessary
    zip_close(pCurZipArchive)
    pCurZipArchive=NULL
  End If
  '
  swCurFileName=sBookFile
  i=WideCharToMultiByte(CP_UTF8,0,@swCurFileName,Len(swCurFileName),Cast(LPSTR,@szCurFileNameUTF8),SizeOf(szCurFileNameUTF8),NULL,NULL)
  *Cast(UByte Ptr,@szCurFileNameUTF8[i])=0    'NULL termination. Was necessary in some cases.

'  If InStr(sBookFile," ") Then
'    If Left(szCurFileNameUTF8,1)<>Chr(&H22) Then szCurFileNameUTF8=Chr(&H22)+szCurFileNameUTF8+Chr(&H22) ' Double-Quotes
'  End If
  '
  pCurZipArchive=zip_open(szCurFileNameUTF8,ZIP_RDONLY,NULL)
  If pCurZipArchive=0 Then                                       'forget 3rd function parameter. The returned error value is false for several situations.
    'If Peek(ULong,@i)<>&H04034B50 Then ....                     'MagicNumber of ZIP files:  "PK\x03\x04" = ASCII &H50 &H4B &H03 &H04 = DWORD &H04034B50
    iRetVal=-1                                                   '-1: Couldn't open ZIP file
    Goto DoneOpenBook
  End If
  '
  iNumArchiveFiles=zip_get_num_entries(pCurZipArchive,NULL) 
  If (iNumArchiveFiles<1) Then
    iRetVal=-2                                                   '-1: No archived files in ZIP file found
    Goto DoneOpenBook
  End If
  '
  'First look, whethere is an #index.txt in the eBook:
  iZipFileIndex=zip_name_locate(pCurZipArchive,@TXT_RBK_INDEXFILE,ZIP_FL_ENC_GUESS Or ZIP_FL_NOCASE)  'Get ZIP index of possible existing "#index.txt" (no UTF8-encoding needed)
  If iZipFileIndex=-1 Then                                                                            'No #index.txt: Sort files by ordinal number
    idxZipIndexFile=INVALID_ZIP_INDEX                                                                 'INVALID_ZIP_INDEX=-1
    hwndShadowTOC=hwndShadowTOC_Sorted
  Else
    idxZipIndexFile=iZipFileIndex
    hwndShadowTOC=hwndShadowTOC_Unsorted                                                              '#index.txt: Leave files in order as given in the index file
  End If
  hwndShadowTOC=hwndShadowTOC_Unsorted                                                                '#index.txt: Copy filenames into a listbox, which sorts alphabetical
  '
  SendMessage(hwndShadowTOC,LB_RESETCONTENT,0,0)                                                      'Empty the (currently used) TOC shadow listbox
  '
  'Now retrieve all files found in the opened ZIP archive.
  'First we copy the file names into an invisible shadow listbox. From the shadow list we *after* copy only true content files with their true chapter titles
  'into the TOC listbox. The TOC listbox item data values store the listbox ID of corresponding entry in the shadow listbox. The shadow listbox we still need
  'to get ZIP file pointers and the file types.
  '
  If idxZipIndexFile=INVALID_ZIP_INDEX Then                                                           'No #index.txt: Copy filenames+ZipIndes into the (sorting) shadow listbox
    i=0
    Do
      pszCurFileNameUTF8=zip_get_name(pCurZipArchive,i,ZIP_FL_ENC_GUESS)                               'Get filename in UTF-8 format
      If pszCurFileNameUTF8=NULL Then Exit Do
      MultiByteToWideChar(CP_UTF8,0,pszCurFileNameUTF8,-1,swCurFileNameWide,SizeOf(swCurFileNameWide)) 'Convert UTF-8 filename into 16-bit Unicode
      j=SendMessage(hwndShadowTOC,LB_ADDSTRING,0,Cast(LPARAM,@swCurFileNameWide))                      'Copy the wide-Unicode filename into the (sorting) shadow listbox
      SendMessage(hwndShadowTOC,LB_SETITEMDATA,j,i)                                                    'Store original ZIP index into item data. We need it later to load the file fom ZIP.
      i+=1
    Loop
  Else                                                                                   '#index.txt exists: Copy filenames in original order from there
    j=True
    Do
      pswCurFileName=IndexFile_get_name(pCurZipArchive,j,idxZipIndexFile)                'Get filename from #index.txt in zip file in wString format
      j=False                                                                            'j is once only True to init IndexFile_get_name()
      If pswCurFileName=NULL Then Exit Do
      szCurFileNameUTF8=Left(*pswCurFileName,InStr(*pswCurFileName,Any "- ")-1)          'Get the true filename as used in ZIP archive, filename "0001" etc. needs no UTF-8 encoding
      iZipFileIndex=zip_name_locate(pCurZipArchive,Cast(Const ZString Ptr,@szCurFileNameUTF8),ZIP_FL_ENC_GUESS)  'Find ZIP file index of this filename
      If iZipFileIndex<>-1 Then
        i=SendMessage(hwndShadowTOC,LB_ADDSTRING,0,Cast(LPARAM,pswCurFileName))          'Copy the filename into the (not sorting) shadow listbox
        SendMessage(hwndShadowTOC,LB_SETITEMDATA,i,iZipFileIndex)                        'Store real ZIP index of file into item data. We need it later to load the file fom ZIP.
      End If
    Loop
  End If
  '
  sMetaData=""                                                                    'Global variable. Stores the conent of #meta.txt if existing or must an be empty string.
  For i=0 To SendMessage(hwndShadowTOC,LB_GETCOUNT,0,0)-1
    SendMessage(hwndShadowTOC,LB_GETTEXT,i,Cast(LPARAM,@swCurFileNameWide))       'Get raw filename from the shadow listbox. Example: "0001-1 The introduction.rtf"
    If *Cast(UShort Ptr,StrPtr(swCurFileNameWide))=&H0023 Then                    'Filename starts with "#" character: #meta.txt, #index.txt and so on. Invisible files, indirectly to handle
      swBuffer=UCase(swCurFileNameWide)
      If swBuffer=TXT_RBK_METAFILE Then
        Select Case UserSettings.wLanguageID                                      'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07 :  swzChapterTitle="01[Über...]"                                'German               very left: "01" means indention level 0, file icon 1
        Case Else :  swzChapterTitle="01[About...]"                               'Neutral (english)    very left: "01" means indention level 0, file icon 1
        End Select
        j=SendMessage(hwndTOC,LB_ADDSTRING,0,Cast(LPARAM,@swzChapterTitle))       'Copy the chapter title derived from filename into the TOC listbox
        SendMessage(hwndTOC,LB_SETITEMDATA,j,i)                                   'Store ShadowList index fitting to this TOC entry
      'ElseIf swBuffer=TXT_RBK_INDEXFILE Then                                     'Just procesed, simply ignore here. Contains the real filenames/chapter names, if exiting.
      'Else                                                                       'Ignore all unknown file names starting with # and don't show them in TOC
      End If
    Else                                                                          'Filename doesn't start with a "#" character
      sFileExtension=UCase(Mid(swCurFileNameWide,InStrRev(swCurFileNameWide,".")+1))
      If sFileExtension="TXT" Then
        swFileType=WChr(TOC_PRE2_TEXT)                                                    'Text files are ignored, reserved for meta info files for each file
      ElseIf sFileExtension="RTF" Then
        swFileType=WChr(TOC_PRE2_RTF)                                                     'icon 1 (=text)   (icon 1 is the meta file)
      ElseIf sFileExtension="PNG" Or sFileExtension="JPG" Or sFileExtension="JPEG" Then
        swFileType=WChr(TOC_PRE2_PICTURE)                                                 'icon 2 (=picture)
      Else
        swFileType=WChr(TOC_PRE2_ATTACHEMENT)                                             'icon 3 (=attachement)
      End If
      '
      If sFileExtension<>"TXT" Then                                                       'Text files are hidden
        If ValInt(swCurFileNameWide)>0 Or *Cast(UShort Ptr,@swCurFileNameWide)=&H0030 Then  'File number exists (and possibly hierarchy level info). --> Chr(&H0030)="0"
          j=InStr(swCurFileNameWide,"-")
          If j=0 Then                                                             'No hierarchy level info exists
            swIndention="0"                                                       'Set main level
            j=InStr(swCurFileNameWide," ")+1                                      'Start of chapter name
          Else
            swIndention=Mid(swCurFileNameWide,j+1,1)                              'Get indention deep digit as a string of 1 character such as "1" or "2" or "3"...
            If ValInt(swIndention)=0 Then                                         'Value 0 or possible a "-" deeper in the filename
              swIndention="0"
              j=InStr(swCurFileNameWide," ")+1                                    'Start of chapter name
            Else
              j=InStr(j+1,swCurFileNameWide," ")+1                                'Chapter title starts after space character beheind level info
            End If
          End If
        Else                                                                      'No file number exists: Use raw filename as chapter title
          swIndention="0"
          j=1                                                                     'Start of chapter name
        End If  'Check existance of filenumber
        '
        'Finally remove file extension from chapter title string:
        k=InStrRev(swCurFileNameWide,".",Len(swCurFileNameWide))                  'Find begin of file extension such as ".txt", ".rtf", ".jpg"
        If k=0 Then                                                               'No file extension
          k=Len(swCurFileNameWide)                                                'End of chapter title string
        Else
          k-=1                                                                    'End of chapter title string       
        End If
        '
        If j<2 And (*Cast(UShort Ptr,@swCurFileNameWide)>&H002F And *Cast(UShort Ptr,@swCurFileNameWide)<&H003A) Then 'No title exists (continuation of last file ("j": Position of "-")
          swzChapterTitle=Str(Val(swIndention)+1)+swFileType+""                  'Indent one more and show character "..." instead of a chapter name
        Else
          swzChapterTitle=swIndention+swFileType+Mid(swCurFileNameWide,j,k-j+1)   'Indention plus filename without prefix and without suffix
        End If
        j=SendMessage(hwndTOC,LB_ADDSTRING,0,Cast(LPARAM,@swzChapterTitle))       'Copy the chapter title derived from filename into the TOC listbox
        SendMessage(hwndTOC,LB_SETITEMDATA,j,i)                                   'Store ShadowList index fitting to this TOC entry
      End If    'No TXT file
    End If    'Check for reserved filenames starting with # character
  Next i
  SendMessage(hwndTOC,LB_SETCURSEL,UserSettings_lLastChapterTitle,0)                '0 (first entry) is default, last chapter otherwise
  '....................................................................................................................................
  DoneOpenBook:
  If iRetVal<0 Then        'Error
    zip_close(pCurZipArchive)
    pCurZipArchive=NULL
  End If
  '
  If iRetVal=-1 Then
    Select Case UserSettings.wLanguageID                                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : MessageBox(hWnd,"Fehler beim Öffnen der Datei"+Chr(13)+sBookFile,swAppName,MB_ICONSTOP)                         '&H07: German
    Case Else : MessageBox(hWnd,"Error opening file"+Chr(13)+sBookFile,swAppName,MB_ICONSTOP)                                   '&H00: Neutral (english)
    End Select
  ElseIf iRetVal=-2 Then
    Select Case UserSettings.wLanguageID                                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : MessageBox(hWnd,"Die Datei ist leer!"+Chr(13,13)+sBookFile,swAppName,MB_ICONSTOP)                               '&H07: German
    Case Else : MessageBox(hWnd,"The file is empty!"+Chr(13,13)+sBookFile,swAppName,MB_ICONSTOP)                                '&H00: Neutral (english)
    End Select
  ElseIf iRetVal=-3 Then
    Select Case UserSettings.wLanguageID                                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : MessageBox(hWnd,"Konnte Inhaltsverzeichnis der Datei nicht laden."+Chr(13,13)+sBookFile,swAppName,MB_ICONSTOP)  '&H07: German
    Case Else : MessageBox(hWnd,"Could not load table of contents."+Chr(13,13)+sBookFile,swAppName,MB_ICONSTOP)                 '&H00: Neutral (english)
    End Select
  Else
    SetWindowText(hWnd,TXT_RTF_BOOK+Mid(UserSettings_swzCurArchiveName,InStrRev(UserSettings_swzCurArchiveName,"\")+1))
    ResizeChildWindows(hWnd,UserSettings.TocWinWidth)                     'Shows or hides TOC depending on number of chapters
    zip_set_default_password(pCurZipArchive,Cast(LPCSTR,StrPtr(!"\0")))   'Reset password for case of encrypted files
    szArchivePassword=""                                                  'Reset password for case of encrypted files
    fDontAskPasswordForThisBook=False                                     'Reset this option
    LoadBookChapter(SendMessage(hwndTOC,LB_GETCURSEL,0,0))
  End If
  '
  Return iRetVal
End Function
'************************************************************************************************************************
Function IndexFile_get_name(ByVal pCurZipArchive As Any Ptr,ByVal fInit As Integer,ByVal idxZipIndexFile As Integer) As Const wString Ptr
  'Returns a NULL terminated wString containing a filename (text line) listed in #index.txt. pCurZipArchive must be the ponter of
  'an open ZIP archive and idxZipIndexFile zhe ZIP file index of #index.txt in this archive.
  'If fInit=TRUE:  Loads initially #index.txt given by ZIP file pointer idxZipIndexFile from ZIP file given by ZIP archive pointer pCurZipArchive,
  '                The return value is a WString pointer, pointing to the begin of the first filename in index file.
  'If fInit=FALSE: Returns a WString pointer, pointing to the begin of the next filename in index file.
  '                If no more file names exists, the function retuns 0.
  'The function returns 0 if error or the address of the string containg the current filename otherwise.
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static pswRetVal As Const wString Ptr,swCurFilename As wString*512,ZipStat As zip_stat_,pCurZipFile As Any Ptr,pusIndexFileDataEnd As Any Ptr
  Static pIndexFileRawData As Any Ptr,pusIndexFileData As Ushort Ptr,pusCurFilename As Ushort Ptr,iBufferLen As Integer,pusCurChar As UShort Ptr
  '
  pswRetVal=0
  If pCurZipArchive=NULL Then Goto Done_IndexFile_get_name                                                             'No ZIP file open
  If idxZipIndexFile=INVALID_ZIP_INDEX Then Goto Done_IndexFile_get_name                                               'No index file 
  '....................................................................................................................................
  If fInit=True Then                                                                                                   'Load and init #index.txt
    pusCurFilename=NULL                                                                                                'Invalidate start address of first line 
    If pusIndexFileData<>NULL Then GlobalFree(pusIndexFileData) : pusIndexFileData=NULL                                'Free old memory
    '
    'Now get size and password info for the index file and open it in ZIP archive:
    If zip_stat_index(pCurZipArchive,idxZipIndexFile,ZIP_FL_ENC_GUESS,@ZipStat)=-1 Then Goto Done_IndexFile_get_name   'Load error info for Index file
    If ZipStat.encryption_method<>ZIP_EM_NONE And szArchivePassword="" And fDontAskPasswordForThisBook=False Then
      If DialogBox(hProgInst,MAKEINTRESOURCE(IDD_PASSWORDDLG),hWndMain,@PasswordDlgProc)=1 Then                        'returns 1 if user pushed "OK" or FALSE otherwise
        zip_set_default_password(pCurZipArchive,@szArchivePassword)                                                    'szArchivePassword is an 8 bit string!
      End If
    End If
    pCurZipFile=zip_fopen_index(pCurZipArchive,idxZipIndexFile,0) 
    If pCurZipFile=NULL Then Goto Done_IndexFile_get_name                                                              'Error open #index.txt in ZIP archive
    '
    'Now unzip file #index.txt into memory and convert into 16 bit wString format if necessary:
    '
    pIndexFileRawData=GlobalAlloc(GMEM_FIXED,ZipStat.size)
    zip_fread(pCurZipFile,pIndexFileRawData,ZipStat.size)
    If *Cast(UByte Ptr,pIndexFileRawData)=&HEF And *Cast(UByte Ptr,pIndexFileRawData+1)=&HBB And *Cast(UByte Ptr,pIndexFileRawData+2)=&HBF Then   'UTF-8 ByteOrderMark "ï»¿"
      iBufferLen=MultiByteToWideChar(CP_UTF8,0,pIndexFileRawData+3,ZipStat.size-3,NULL,NULL)*2                                  'Get buffer size in bytes (!) for output buffer
      pusIndexFileData=GlobalAlloc(GMEM_FIXED,iBufferLen+1)
      iBufferLen=MultiByteToWideChar(CP_UTF8,0,pIndexFileRawData+3,ZipStat.size-3,pusIndexFileData,iBufferLen)*2                'Convert UTF-8 encoded text into 16-bit Unicode
      *Cast(UShort Ptr,pusIndexFileData+iBufferLen\2)=&H0000                                                                    'NULL termination of string
    Else                                                                                                                        'Assume ANSI text (we intentional don't support more then ANSI and UTF8)
      iBufferLen=MultiByteToWideChar(CP_ACP,0,pIndexFileRawData,ZipStat.size,NULL,NULL)*2                                       'Get buffer size in bytes (!) for output buffer
      pusIndexFileData=GlobalAlloc(GMEM_FIXED,iBufferLen+1)
      iBufferLen=MultiByteToWideChar(CP_ACP,0,pIndexFileRawData,ZipStat.size,pusIndexFileData,iBufferLen)*2                     'Convert ANSI text into 16-bit Unicode
      *Cast(UShort Ptr,pusIndexFileData+iBufferLen\2)=&H0000                                                                    'NULL termination of string
    End If
    If pIndexFileRawData<>NULL Then GlobalFree(pIndexFileRawData) : pIndexFileRawData=NULL
    '
    'Now #index.txt is in wSting format in a global memory block. We now replace all CR+RTs by NULL characters. This we use as a NULL termination
    'for the strings returned. In later calls of this function we simply search for NULL characters while counting up until the wished index is reached.
    'then we return the start address. That's all.
    '
    pusIndexFileDataEnd=pusIndexFileData+(iBufferLen\2)-1
    For pusCurChar=pusIndexFileData To pusIndexFileDataEnd
      Select Case *pusCurChar
       Case &H000A,&H000D :  *pusCurChar=&H0000
      End Select
    Next pusCurChar
    pusCurFilename=pusIndexFileData        'Start address of first line is start of #index.txt. Also keep this value for next function call
    pswRetVal=pusCurFilename               'Return start of current filename/text line
    Goto Done_IndexFile_get_name           'Nothing more to do
  End If
  '....................................................................................................................................
  'fInit>0: Get start of line from memory block containing #index.txt in wString format:
  '
  If (pusCurFilename<pusIndexFileData) Or (pusIndexFileData=NULL) Then
    pswRetVal=NULL
    Goto Done_IndexFile_get_name
  End If
  '
  For pusCurChar=pusCurFilename To pusIndexFileDataEnd                       'Find end of line
    If *pusCurChar=&H0000 Then Exit For
  Next pusCurChar
  If pusCurChar>=pusIndexFileDataEnd Then                                    'End of #index.txt, no more file names/lines
    pswRetVal=NULL
  Else                                                                       'Find start of next line
    Do
      pusCurChar+=1
      If pusCurChar>pusIndexFileDataEnd Then
        pswRetVal=NULL                                                       'No more filenames
        Exit Do
      End If
      If *pusCurChar<>&H0000 Then
        pusCurFilename=pusCurChar                                            'Store address for next function call
        pswRetVal=pusCurFilename                                             'Return start of current filename/text line
        Exit Do
      End If
    Loop
  End If
  '....................................................................................................................................
  Done_IndexFile_get_name:
  Function=pswRetVal
End Function
'************************************************************************************************************************
Function LoadBookChapter(ByVal iCurChapter As Integer) As LRESULT
  'Loads the chapter file given in iCurChapter from ZIP book file currently open. 
  '
  'see http://www.nih.at/libzip/
  'see http://libzip.org/documentation/libzip.html
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'Used global variables:
  'pCurZipArchive As Any Ptr                     'LibZip's ZIP pointer to the ZIP archive currently open or NULL if no ZIP book is open.
  'hwndMain
  'hWndToc
  'hWndText
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static iRetVal As Integer,i As Integer,swBuffer As Wstring*256,swBuffer2 As Wstring*256,szFileExtension As ZString*4,hCurFileData As HGLOBAL
  Static swFileName As wString*MAX_PATH,pCurZipFile As Any Ptr,pCurFileData As Any Ptr,sBuffer As String,sBuffer2 As String
  Static iZipArchiveIndex As Integer,SetTxtEx As SETTEXTEX,ZipStat As zip_stat_,ParaFmt As PARAFORMAT2,pt As Point
  Static fVeryFirstChapterLoaded As Integer 'Must (!) be static. Used to restore last position for very first book loaded after program start.
  '
  iRetVal=0                                               'Initial function return value: Success
  '
  If iCurChapter=LB_ERR Then Goto DoneLoadBookChapter     'Selection error in TOC listbox. Do nothing, return "silent" without error vallue.
  '
  If pCurZipArchive=NULL Then
    iRetVal=-1                                            'No ZIP archive opened
    Goto DoneLoadBookChapter
  End If
  '
  SendMessage(hwndTOC,LB_SETCURSEL,Cast(WPARAM,iCurChapter),NULL)           'Select the TOC listbox entry for the current chapter (if not just done)
  iZipArchiveIndex=SendMessage(hWndShadowToc,LB_GETITEMDATA,SendMessage(hWndToc,LB_GETITEMDATA,iCurChapter,0),0)  'The shadow TOC stores the true ZIP file index in user data
  '
  '    ZipStat.valid                 'which fields have valid values
  '    ZipStat.Name                  'name of the file  (Const zstring Ptr )
  '    ZipStat.index                 'index within archive
  '    ZipStat.size                  'size of file (uncompressed)
  '    ZipStat.comp_size             'size of file (compressed)
  '    ZipStat.mtime                 'modification time
  '    ZipStat.crc                   'crc of file data
  '    ZipStat.comp_method           'compression method used
  '    ZipStat.encryption_method     'encryption method used
  '    ZipStat.flags                 'reserved for future use
  '
  If zip_stat_index(pCurZipArchive,iZipArchiveIndex,ZIP_FL_ENC_GUESS,@ZipStat)=-1 Then
    Select Case UserSettings.wLanguageID                                                   'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Konnte die Dateiinformationen zum Kapitel nicht ermitteln."      'German: Default language in resources
    Case Else : swBuffer="Could not obtain file information of chapter."                   'Neutral (english)
    End Select
    SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,@swBuffer))
    iRetVal=-2
    Goto DoneLoadBookChapter
  End If
  '
  If ZipStat.encryption_method<>ZIP_EM_NONE And szArchivePassword="" And fDontAskPasswordForThisBook=False Then
    If DialogBox(hProgInst,MAKEINTRESOURCE(IDD_PASSWORDDLG),hWndMain,@PasswordDlgProc)=1 Then 'returns 1 if the user leaves the dialog with "OK" or FALSE otherwise
      zip_set_default_password(pCurZipArchive,@szArchivePassword) 'szArchivePassword is an 8 bit string!
    End If
  End If
  '
  pCurZipFile=zip_fopen_index(pCurZipArchive,iZipArchiveIndex,0) 
  If (pCurZipFile=NULL) Then
    szArchivePassword=""                                      'Erase possibly wrong password, so thet next time Password dialog appears again
    Select Case UserSettings.wLanguageID                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer2="Konnte Kapitel nicht öffnen."      'German: Default language in resources
    Case Else : swBuffer2="Could not open chapter."           'Neutral (english)
    End Select
    swBuffer=swBuffer2+Chr(13,10)+"("+Chr(34)+*Cast(ZString Ptr,zip_strerror(pCurZipArchive))+Chr(34)+")"
    SetTxtEx.flags=ST_DEFAULT
    SetTxtEx.codepage=1200                  '1200: Unicode code page, CP_ACP: system code page
    SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,@swBuffer))
    iRetVal=-3
    Goto DoneLoadBookChapter
  End If    
  '
  If (ZipStat.size>3) And (ZipStat.valid Or ZIP_STAT_SIZE) And (ZipStat.valid Or ZIP_STAT_NAME) Then
    SendMessage(hWndShadowToc,LB_GETTEXT,SendMessage(hWndToc,LB_GETITEMDATA,iCurChapter,0),Cast(lParam,@swFileName))  'Get filename from Shadow TOC 
    szFileExtension=UCase(Mid(swFileName,InStrRev(swFileName,".")+1))
    If szFileExtension="RTF" Then
      pCurFileData=GlobalAlloc(GMEM_FIXED,ZipStat.size+1)
      Poke pCurFileData+ZipStat.size,0                    'NULL termination of string
      zip_fread(pCurZipFile,pCurFileData,ZipStat.size)
      'Hint: RTF is never Unicode because it encodes Unicode internally. For plain text we use CP_ACP because of the undocumented (?) functionality
      '      of the RichText control to automaticly convert plain text for CP_ACP from UTF8, if BOM exists.
      SetTxtEx.flags=ST_DEFAULT                           '
      SetTxtEx.codepage=CP_ACP                            '1200: Unicode code page, CP_ACP: system code page
      SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,pCurFileData))
      GlobalFree(pCurFileData)
    ElseIf szFileExtension="JPG" Or szFileExtension="JPEG" Or szFileExtension="PNG" Then
      hCurFileData=GlobalAlloc(GMEM_MOVEABLE,ZipStat.size+1)   'GMEM_MOVEABLE is a "must" for the function DecompressPicture()
      pCurFileData=GlobalLock(hCurFileData)
      Poke pCurFileData+ZipStat.size,0                         'NULL termination of string
      zip_fread(pCurZipFile,pCurFileData,ZipStat.size)
      GlobalUnlock(hCurFileData)
      hCurFileData=PictureToRtf(hCurFileData)                  'Decompress picture, get back RTF file in (another!) global memory
      pCurFileData=GlobalLock(hCurFileData)
      SetTxtEx.flags=ST_DEFAULT                       '
      SetTxtEx.codepage=1200                                   '1200: Unicode code page, CP_ACP: system code page
      SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,pCurFileData))
      GlobalUnlock(hCurFileData)
      GlobalFree(hCurFileData)                                 'We must free the memory buffer got from function DecompressPicture()
      'Because centering image using RTF codes doesn't work at least with RichEdit v4.1, we do it "by handwork":
      ParaFmt.cbSize=SizeOf(ParaFmt)
      ParaFmt.dwMask=PFM_ALIGNMENT Or PFM_SPACEBEFORE
      ParaFmt.wAlignment=PFA_CENTER                                      'Center the graphic
      ParaFmt.dySpaceBefore=100                                          '100 twips space above the graphic
      SendMessage(hWndText,EM_SETPARAFORMAT,0,Cast(lParam,@ParaFmt))
    ElseIf szFileExtension="TXT" And UCase(swFileName)=TXT_RBK_METAFILE Then
      pCurFileData=GlobalAlloc(GMEM_FIXED,ZipStat.size+1)
      Poke pCurFileData+ZipStat.size,0                    'NULL termination of string
      zip_fread(pCurZipFile,pCurFileData,ZipStat.size)
      pCurFileData=MetaInfoToRtf(pCurFileData)
      SetTxtEx.flags=ST_DEFAULT                           '
      SetTxtEx.codepage=1200                  '1200: Unicode code page, CP_ACP: system code page. RTF is not Unicode, because it encodes Unicode internally.
      SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,pCurFileData))
      GlobalFree(pCurFileData)
    Else                                                                               'Attachement
      'We generate an RTF file with a hyperlink on the fly. A juser click at the link generates than a WM_NOTIFY message with EN_LINK code.
      'Example 1: sBuffer="{\rtf1{\field{\*\fldinst{ HYPERLINK www.rowalt.de/rtfbook/ }}{\fldrslt{Klick me!}}}}"
      'Example 2: sBuffer="{\rtf1 www.rowalt.de/rtfbook/}"                                                                    <----- works with autolink flag only!
      sBuffer="{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\lang1031\f0\fs20 "
      Select Case UserSettings.wLanguageID                                                 'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : sBuffer=sBuffer+"Bin\u228?r-Attachement. Zum Extrahieren der Datei bitte den folgenden Link anklicken:"    'German: Default language in resources
      Case Else : sBuffer=sBuffer+"Binary Attachemnt. To extract the file to disc please click the following link:"         'Neutral (english)
      End Select
      If *Cast(UShort Ptr,@swFileName)>&H002F And *Cast(UShort Ptr,@swFileName)<&H003A Then 'First character of the filename is a number ("0"..."9")
        sBuffer2=Mid(swFileName,InStr(swFileName," ")+1)  'Remove prefix numbers up to terminating space character (we don't filename need to extract, it's only a cosmetical touch)
      Else
        sBuffer2=swFileName   'No number at begin of filename: Use filename without change
      End If
      sBuffer=sBuffer+"\par\par{\field{\*\fldinst{HYPERLINK "+TXT_LINKID_BAGGAGE+" }}{\fldrslt{\fs28\b "+sBuffer2+"\b0\fs20}}}"+"\par\par ("+Str(ZipStat.size\1024)+" KB)}"
      SetTxtEx.flags=ST_DEFAULT                       '
      SetTxtEx.codepage=CP_ACP                  '1200: Unicode code page, CP_ACP: system code page
      SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,StrPtr(sBuffer)))
    End If
  Else     'Buffer size 0 or <4    [  If (ZipStat.size>3) And (ZipStat.valid Or ZIP_STAT_SIZE) And (ZipStat.valid Or ZIP_STAT_NAME) Then ]
    Select Case UserSettings.wLanguageID                               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Konnte Text des Kapitels nicht laden."      'German: Default language in resources
    Case Else : swBuffer="Could not load the chapter's text."         'Neutral (english)
    End Select
    SetTxtEx.flags=ST_DEFAULT                       '
    SetTxtEx.codepage=1200                  '1200: Unicode code page, CP_ACP: system code page
    SendMessage(hWndText,EM_SETTEXTEX,Cast(wParam,@SetTxtEx),Cast(lParam,@swBuffer))
  End If   ' If (ZipStat.size>3) And (ZipStat.valid Or ZIP_STAT_SIZE) And (ZipStat.valid Or ZIP_STAT_NAME) Then
  zip_fclose(pCurZipFile)
  '
  If fVeryFirstChapterLoaded=False Then   'First book and first chapter opened after program start. Restore last text position now.
    fVeryFirstChapterLoaded=True
    If UserSettings_swzCurArchiveName=UserSettings_swzLastArchiveName Then
      UserSettings_swzLastArchiveName=""                                    'Not longer needed. Just to prevent errors. Next access when program quits.
      SendMessage(hWndText,EM_SETSEL,Cast(wParam,UserSettings_dwLastChapterPos),Cast(lParam,UserSettings_dwLastChapterPos)) 'Restore text position or set to begin
      UserSettings_dwLastChapterPos=0      'Just to prevent errors. Should be not necessary normally.
    End If
  Else
    SendMessage(hWndText,EM_SETSEL,0,0)                                      'Set to begin of text
    UpdateMainMenu(hwndMain)                                                   'For enabling/disabling up/down controls
  End If
  '
  SendMessage(hWndText,EM_SETZOOM,UserSettings.TextZoom,ZOOM_NORMAL)
  UpdateMainMenu(hwndMain)                                                   'For enabling/disabling up/down controls
  '....................................................................................................................................
  DoneLoadBookChapter:
  '
  Return iRetVal
End Function
'************************************************************************************************************************
Function TextWndProc(ByVal hWnd As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'Window subclass function for the RichEdit text window
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static pt As Point,dwTimeOfCurChar As Dword,dwTimeOfLastChar As Dword,swSearchString As String,i As Integer,swBuffer As Wstring*255
  Static tFindTextEx As FINDTEXTEXW,tCharRange As CHARRANGE,iFindPos As Integer
  '
  Select Case uMsg
  Case WM_RBUTTONUP 'A RichEdit control has no built-in cut/copy/paste context menu. We create such a context menu for handling of bookmarks
    'SetWindowText(hwndMain,"Debug: x="+Str$(LoWord(lParam))+", y="+Str$(HiWord(lParam)))
    pt.x=LoWord(lParam)                                                      'Copy mouse coordinates (Client position) into tPoint type
    pt.y=HiWord(lParam)                                                      'Copy mouse coordinates (Client position) into tPoint type
    ClientToScreen(hWnd,@pt)
    'The menu for the next line we take from main menu: Index 1 (zero based) is menu "Edit" 
    TrackPopupMenu(GetSubMenu(hMenuMain,1),TPM_LEFTALIGN,pt.x,pt.y,0,hWnd,ByVal 0)   'Put context menu where mouse is
    Return CallWindowProc(cast(WNDPROC,pfncOrigTxtWndProc),hWnd,uMsg,wParam,lParam) 'Pass the parameters to the original RichEdit window function
  Case WM_CHAR         'Used for QuickSearch
    If wParam>0 Then                 'Cannot occur in a normal way, but usefull for "Quick search again" using PostMessage(hwndTOC,WM_CHAR,0,0)
      dwTimeOfCurChar=GetTickCount()
      If (dwTimeOfLastChar=0) Or (dwTimeOfCurChar>dwTimeOfLastChar+QUICKSEARCH_TIMEOUT_MS) Then
        swSearchString=WChr(wParam)
      Else 
        swSearchString=swSearchString+WChr(wParam)
      End If
      dwTimeOfLastChar=dwTimeOfCurChar
    End If
    '
    swBuffer=swSearchString
    SendMessage(hWndText,EM_EXGETSEL,0,Cast(LPARAM,@tCharRange))
    For i=1 To 2
      tFindTextEx.chrg.cpMin=tCharRange.cpMax
      tFindTextEx.chrg.cpMax=-1
      tFindTextEx.lpstrText=Cast(LPCTSTR,@swBuffer)
      iFindPos=SendMessage(hWnd,EM_FINDTEXTEXW,FR_DOWN Or FR_MATCHCASE,Cast(LPARAM,@tFindTextEx))
      If iFindPos<>-1 Then
        SendMessage(hWnd,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))                 'Select found text
        Exit For
      End If
      tCharRange.cpMax=0   'Start again at begin of text
    Next i
    Return 0     
  Case Else
    Return CallWindowProc(cast(WNDPROC,pfncOrigTxtWndProc),hWnd,uMsg,wParam,lParam) 'Pass the parameters to the original RichEdit window function
  End Select
End Function
'***************************************************************************************************************************************
Function TocWndProc(ByVal hWnd As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'Window subclass function for the TableOfContents listbox window
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static swCurLbString As wString*MAX_PATH,sSearchString As String,iLbIndex As Integer,iLbFirstIndex As Integer,iLbLastIndex As Integer
  Static dwTimeOfCurChar As Dword,dwTimeOfLastChar As Dword,usCurChar As UShort,i As Integer,iCurSel As Integer
  '
  Select Case uMsg
  Case WM_CHAR       'Select first the listbox entry, starting with the typed character
    'Each entry in the TOC listbox starts with two (Unicode!) prefix characters specifying internal information:
    'Character 1: Indention level "0"..."9"
    'Character 2: File type "0"..."4" (see #define TOC_PRE2_...)
    'Therefore we can't simply use LB_GETTEXT to find the lisbox entry starting with the just now typed character.
    'To deal with different character sets, we do the job case sensitive (this is not a trivial problem!).
    'Characters typed with a pause not longer than a 1000 milliseconds (QUICKSEARCH_TIMEOUT_MS) are collected to one search term.
    'Using PostMessage(hwndTOC,WM_CHAR,0,0) ---> character value NULL, you can search again using the existing search term without timeout.
    '
    If wParam>0 Then                 'Cannot occur in a normal way, but usefull for "Quick search again" using PostMessage(hwndTOC,WM_CHAR,0,0)
      dwTimeOfCurChar=GetTickCount()
      If (dwTimeOfLastChar=0) Or (dwTimeOfCurChar>dwTimeOfLastChar+QUICKSEARCH_TIMEOUT_MS) Then
        sSearchString=WChr(wParam)
      Else 
        sSearchString=sSearchString+WChr(wParam)
      End If
      dwTimeOfLastChar=dwTimeOfCurChar
    End if
    '
    iLbFirstIndex = SendMessage(hWnd,LB_GETCURSEL,0,0)+1
    iLbLastIndex  = SendMessage(hWnd,LB_GETCOUNT,0,0)-1
    For i=1 To 2                'Two rounds, because we start at current list position and make one turn to begin, if necessary
      For iCurSel=iLbFirstIndex To iLbLastIndex
        SendMessage(hWnd,LB_GETTEXT,iCurSel,Cast(LPARAM,@swCurLbString))
        If InStr(swCurLbString,sSearchString) Then
          SendMessage(hWnd,LB_SETCURSEL,iCurSel,0)
          PostMessage(hWndMain,WM_COMMAND,MakeLong(IDC_TOCWND,LBN_SELCHANGE),0)    'Direct main program loop to load the chapter
          Exit For,For                                                             'Exit both loops
        End If
      Next iCurSel
      iLbFirstIndex=0   'Start at begin of list in the next turn, if we do the second turn
    Next i
    Return 0
  Case Else
    Return CallWindowProc(cast(WNDPROC,pfncOrigTocWndProc),hWnd,uMsg,wParam,lParam) 'Pass the parameters to the original RichEdit window function
  End Select
  'Return CallWindowProc(cast(WNDPROC,pfncOrigTocWndProc),hWnd,uMsg,wParam,lParam) 'Pass the parameters to the original RichEdit window function
End Function
'***************************************************************************************************************************************
Function ExtractCurFile(ByVal hWndParent As HWND) As LRESULT
  'Extracts currently in TOC selected file to disc, if the user doesn't cancel.
  'The function is very similiar to function LoadBookChapter(), but has another output ;-)
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static iRetVal As Integer,iCurChapter As Integer,iZipArchiveIndex As Integer,ZipStat As zip_stat_,swOutFileName As Wstring*MAX_PATH
  Static pCurZipFile As Any Ptr,pCurFileData As Any Ptr,tOfn As OPENFILENAME,swBuffer As Wstring*MAX_PATH,fSourceIsPicture As Integer
  Static hFile As HANDLE,dwBytesWritten As Dword,swOriginalFileExtension As WString*MAX_PATH,sBuffer As String,i As Integer
  Static pCurInByte As UByte Ptr,pszOut As zString Ptr,pOutBuffer As ZString Ptr,swDefExtension As Wstring*4
  '
  iRetVal=0                                               'Initial function return value: Success
  pCurZipFile=NULL
  pCurFileData=NULL
  hFile=INVALID_HANDLE_VALUE
  '
  iCurChapter=SendMessage(hWndToc,LB_GETCURSEL,0,0)
  '
  If pCurZipArchive=NULL Then
    iRetVal=-1
    Goto Done_ExtractCurFile
  End If
  '
  iZipArchiveIndex=SendMessage(hWndShadowToc,LB_GETITEMDATA,SendMessage(hWndToc,LB_GETITEMDATA,iCurChapter,0),0)  'The shadow TOC stores the true ZIP file index in user data
  If iZipArchiveIndex=LB_ERR Then
    iRetVal=-2
    Goto Done_ExtractCurFile
  End If
  '
  If zip_stat_index(pCurZipArchive,iZipArchiveIndex,ZIP_FL_ENC_GUESS,@ZipStat)=-1 Then
    iRetVal=-3
    Goto Done_ExtractCurFile
  End If
  '
  If ZipStat.encryption_method<>ZIP_EM_NONE And szArchivePassword="" And fDontAskPasswordForThisBook=False Then
    If DialogBox(hProgInst,MAKEINTRESOURCE(IDD_PASSWORDDLG),hWndParent,@PasswordDlgProc)=1 Then 'returns 1 if the user leaves the dialog with "OK" or FALSE otherwise
      zip_set_default_password(pCurZipArchive,@szArchivePassword) 'szArchivePassword is an 8 bit string!
    End If
  End If
  '
  pCurZipFile=zip_fopen_index(pCurZipArchive,iZipArchiveIndex,0) 
  If (pCurZipFile=NULL) Then
    szArchivePassword=""                                      'Erase possibly wrong password, so thet next time Password dialog appears again
    iRetVal=-4
    Goto Done_ExtractCurFile
  End If
  '
  If (ZipStat.size<4) Or (ZipStat.valid Or ZIP_STAT_SIZE)=0 Or (ZipStat.valid Or ZIP_STAT_NAME)=0 Then
    iRetVal=-5
    Goto Done_ExtractCurFile
  End If
  '
  MultiByteToWideChar(CP_UTF8,0,ZipStat.Name,-1,swOutFileName,SizeOf(swOutFileName)) 'Convert UTF-8 filename into 16-bit Unicode
  swOriginalFileExtension=UCase(Mid(swOutFileName,InStrRev(swOutFileName,".")+1))
  fSourceIsPicture=False
  If swOriginalFileExtension="PNG" Or swOriginalFileExtension="JPG" Or swOriginalFileExtension="JPEG" Then fSourceIsPicture=True
  '
  tOfn.lStructSize     = SizeOf(OPENFILENAME)
  tOfn.hwndOwner       = hWndParent
  If fSourceIsPicture=True Then
    Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"RTF-Datei, Bild eingebettet (*.rtf)\0*.rtf\0Bilddatei (*.png,*.jpg)\0*.png;*.jpg;*.jpeg\0Alle Dateien (*.*)\0*.*\0\0")))
    Case Else : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"RTF file, picture enbedded (*.rtf)\0*.rtf\0Picture file (*.png,*.jpg)\0*.png;*.jpg;*.jpeg\0All files (*.*)\0*.*\0\0")))
    End Select
    tOfn.nFilterIndex = 1   'RTF is default even for graphic files!
    swDefExtension="rtf"
  ElseIf swOriginalFileExtension="RTF" Then
    Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"RTF-Textdatei (*.rtf)\0*.rtf\0Alle Dateien (*.*)\0*.*\0\0")))
    Case Else : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"RTF text file (*.rtf)\0*.rtf\0All files (*.*)\0*.*\0\0")))
    End Select
    tOfn.nFilterIndex = 1
    swDefExtension="rtf"
  Else                 'Attachement files, text files (#meta.txt)
    Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"Alle Dateien (*.*)\0*.*\0\0")))
    Case Else : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"All files (*.*)\0*.*\0\0")))
    End Select
    tOfn.nFilterIndex = 3
    swDefExtension=""
  End If
  tOfn.lpstrDefExt     = @swDefExtension
  tOfn.lpstrFile       = @swOutFileName
  tOfn.lpstrInitialDir = @swOutFileName
  tOfn.nMaxFile        = MAX_PATH
  tOfn.Flags           = OFN_PATHMUSTEXIST Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_ENABLESIZING Or OFN_OVERWRITEPROMPT
  If GetSaveFileName(@tOfn)<>0 Then
    sBuffer=UCase(Mid(swOutFileName,InStrRev(swOutFileName,".")+1))     'Get file extension set by user
    If tOfn.nFilterIndex=1 And fSourceIsPicture=True Then
      If sBuffer="PNG" Or sBuffer="JPG" Or sBuffer="JPEG" Then swOutFileName=swOutFileName+".rtf"  'For RTF we use double-extension for RTF embedded picture files, for example "MyFile.png.rtf"
    End If
    pCurFileData=GlobalAlloc(GMEM_FIXED,ZipStat.size)
    zip_fread(pCurZipFile,pCurFileData,ZipStat.size)
    hFile=CreateFile(Cast(LPCWSTR,@swOutFileName),GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL)
    If hFile=INVALID_HANDLE_VALUE Then
      iRetVal=-6                'Creating error
      Goto Done_ExtractCurFile
    End If
    '
    If tOfn.nFilterIndex=1 And fSourceIsPicture=True Then  'tOfn.nFilterIndex=1: RTF
      If swOriginalFileExtension="PNG" Then 
        sBuffer="{\rtf1\ansi "+Chr(13)+"{\pict\wmetafile1\pngblip\picwgoal240\pichgoal240 "
      Else
        sBuffer="{\rtf1\ansi "+Chr(13)+"{\pict\wmetafile1\jpegblip\picwgoal240\pichgoal240 "
      End If
      If WriteFile(Cast(HANDLE,hFile),StrPtr(sBuffer),Len(sBuffer),@dwBytesWritten,NULL)=0 Then
        iRetVal=-6   'Writing error
        Goto Done_ExtractCurFile     'For this error the user error message is done below
      End If
      '
      pOutBuffer=GlobalAlloc(GMEM_FIXED,ZipStat.size*2)
      pszOut=pOutBuffer
      For pCurInByte=Cast(UBYTE Ptr,pCurFileData) To Cast(UBYTE Ptr,pCurFileData)+ZipStat.size-1
        *pszOut=Hex(*pCurInByte,2)
        pszOut+=2                                                                            'Increment by two bytes (each HEX byte needs space of 2 characters)
      Next pCurInByte
      If WriteFile(Cast(HANDLE,hFile),pOutBuffer,ZipStat.size*2,@dwBytesWritten,NULL)=0 Then
        iRetVal=-6   'Writing error
        Goto Done_ExtractCurFile     'For this error the user error message is done below
      End If
      GlobalFree(pOutBuffer)
      '
      sBuffer=" }}"
      If WriteFile(Cast(HANDLE,hFile),StrPtr(sBuffer),3,@dwBytesWritten,NULL)=0 Then
        iRetVal=-6   'Writing error
        Goto Done_ExtractCurFile     'For this error the user error message is done below
      End If
    Else            'Standard RTF files, plain picture files, plain text files and attachement files: Simply extract to disk.
      If WriteFile(Cast(HANDLE,hFile),pCurFileData,ZipStat.size,@dwBytesWritten,NULL)=0 Then
        iRetVal=-6   'Writing error
        Goto Done_ExtractCurFile     'For this error the user error message is done below
      End If
    End If
  End If
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  Done_ExtractCurFile:
  If iRetVal Then
    Select Case iRetVal
    Case -1
      Select Case UserSettings.wLanguageID                                                  'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : swBuffer="Kein Buch geöffnet!"                                            'German: Default language in resources
      Case Else : swBuffer="No book open."                                                  'Neutral (english)
      End Select
    Case -2
      Select Case UserSettings.wLanguageID                                                  'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : swBuffer="Ist kein Kapitel ausgewählt?"                                   'German: Default language in resources
      Case Else : swBuffer="Is there no chapter selected?"                                  'Neutral (english)
      End Select
    Case -3
      Select Case UserSettings.wLanguageID                                                   'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : swBuffer="Konnte die Dateiinformationen zum Kapitel nicht ermitteln."      'German: Default language in resources
      Case Else : swBuffer="Could not obtain file information of chapter."                   'Neutral (english)
      End Select
    Case -4
      Select Case UserSettings.wLanguageID                                                   'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : swBuffer="Konnte Kapitel nicht öffnen."                                    'German: Default language in resources
      Case Else : swBuffer="Could not open chapter."                                         'Neutral (english)
      End Select
    Case -5
      Select Case UserSettings.wLanguageID                                                   'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : swBuffer="Konnte Datei nicht laden."                                       'German: Default language in resources
      Case Else : swBuffer="Could not load file."                                            'Neutral (english)
      End Select
    Case Else  '-6
      Select Case UserSettings.wLanguageID                                                   'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : swBuffer="Fehler beim Schreiben der Datei."                                'German: Default language in resources
      Case Else : swBuffer="Error writing file."                                             'Neutral (english)
      End Select
    End Select
    MessageBox(hWndParent,swBuffer,TXT_APP_NAME,MB_ICONSTOP)
  End If
  '
  If pCurFileData Then GlobalFree(pCurFileData)
  If pCurZipFile<>NULL Then zip_fclose(pCurZipFile)
  If hFile<>INVALID_HANDLE_VALUE Then CloseHandle(hFile)
  Function=iRetVal
End Function
'***************************************************************************************************************************************
Function ExportDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'Window function for the modal "Export" Dialog box.
  'Idea for later: Use TXT_UNICODE_SYSTEMFONT for listbox, so that we can use better symbol characters
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static i As Long,j As Long,k As Long,szwBuffer As Wstring*MAX_PATH,szwBuffer2 As Wstring*MAX_PATH,szwFilename As Wstring*MAX_PATH
  Static hwndChapterListbox As HWND,hwndOK As HWND,hwndFileName As HWND,tOfn As OPENFILENAME
  '
  Select Case(uMsg)
  Case WM_INITDIALOG
    hwndChapterListbox=GetDlgItem(hDlg,IDCL_CHAPTER)
    hwndOK=GetDlgItem(hDlg,IDOK)
    hwndFileName=GetDlgItem(hDlg,IDCE_FILENAME)
    SendMessage(hwndChapterListbox,WM_SETFONT,Cast(WPARAM,hTocFont),False)   'Select a Unicode font (hTocFont is global)
    SendMessage(hwndChapterListbox,LB_SETHORIZONTALEXTENT,1024,0)            'Horizontal scrolling range (estimated only)
    '
    Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07                                         'German: Default language in resources
      SetDlgItemText(hDlg,IDCS_EXPRTHLP,"Es können mehrere Kapitel ausgewählt werden."+Chr(13,13)+"Maus mit Shift-Taste wählt eine Folge. Das ist auch mit einem "+_
                                        "Ziehen der Maus möglich."+Chr(13,13)+"Mit Maus und Strg-Taste können beliebig einzelne Einträge gewählt oder abgewählt werden.")
    Case Else                                         'Neutral (english)
      SetWindowText(hDlg,"Export as KeyNote file...")
      SetDlgItemText(hDlg,IDCS_FILENAME,"&File name:")
      SetDlgItemText(hDlg,IDCS_CHAPTERTEXT,"C&hapters to export (multi selection):")
      SetDlgItemText(hDlg,IDCANCEL,"&Cancel")
      SetDlgItemText(hDlg,IDCS_EXPRTHLP,"You can select several chapters."+Chr(13,13)+"Mouse with Shift key selects a sequence. This can also be done by dragging "+_
                                        "the mouse."+Chr(13,13)+"Use mouse and the Ctrl key to select or deselect individual entries.")
    End Select
    '
    szwBuffer=Left(UserSettings_swzCurArchiveName,InStrRev(UCase(UserSettings_swzCurArchiveName),"."))+"knt"
    SetWindowText(hwndFileName,szwBuffer)
    i=Len(szwBuffer)
    SendMessage(hwndFileName,EM_SETSEL,Cast(WPARAM,i),Cast(LPARAM,i))    'Set caret to end. This makes the current file extension visible without scrolling
    '
    For i=0 To SendMessage(hWndToc,LB_GETCOUNT,0,0)-1       'Create an exact copy of the TOC in the local multi-select listbox
      SendMessage(hWndToc,LB_GETTEXT,i,Cast(lParam,@szwBuffer))
      'File title starts (unvsable for user) with a 2-digits decimal number prefix. Digit 1: Indention, Digit 2: file type:
      'The following file types are defined:   0: Text (ignored!),   1: info (Meta Info),  2: RTF-Text,  3: Picture,  4: Attachement
      '
      j=(*Cast(UShort Ptr,StrPtr(szwBuffer))-&H30)*4          'Get number string as an integer (ASCII value mins &H30); 4 spaces per indention level
      szwBuffer2=Wstring(j,WStr(" "))                         'Emulate indention by indention ;-) (2 spaces per indention level)
      '
      j=False                                                                               'j=TRUE: Ignore this file for export
      Select Case *Cast(UShort Ptr,StrPtr(szwBuffer)+2)                                     'Second digit in file title prefix: file type
      Case TOC_PRE2_RTF      : szwBuffer2=szwBuffer2+WChr(EXP_CHAR_RTF)+Mid(szwBuffer,3)    'Code/"icon" for RTF text
      Case TOC_PRE2_PICTURE  : szwBuffer2=szwBuffer2+WChr(EXP_CHAR_PIC)+Mid(szwBuffer,3)    'Code/"icon" for Picture
      Case Else              : j=TRUE                                                       'Ignore all other file types
      End Select
      If j=FALSE Then
        j=SendMessage(hwndChapterListbox,LB_ADDSTRING,NULL,Cast(lParam,@szwBuffer2))
        SendMessage(hwndChapterListbox,LB_SETITEMDATA,j,SendMessage(hWndToc,LB_GETITEMDATA,i,NULL))   'Original ZIP index wwe need to load the file fom ZIP.
      End If
    Next i
    '
    PostMessage(hDlg,WM_COMMAND,MakeLong(IDCL_CHAPTER,LBN_SELCHANGE),0)                            'Results initial enabling/disabling of OK button
    SetFocus(GetDlgItem(hDlg,IDCB_FILENAME))
    '
    Return 0        'TRUE: Let the system set the focus
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDOK
      GetWindowText(hwndFileName,szwBuffer,SizeOf(szwBuffer))
      szwBuffer2=UCase(Mid(szwBuffer,InStrRev(szwBuffer,".")))
      Select Case szwBuffer2
'      Case ".RTF",".DOC"                                  '*.doc: Some people know this better
'        If ExportRtf(hDlg,hwndChapterListbox,szwBuffer)=0 Then EndDialog(hDlg,0)
      Case ".KNT"
        If ExportKeyNote(hDlg,hwndChapterListbox,szwBuffer)=0 Then EndDialog(hDlg,0)
      Case Else
        Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07                                         'German: Default language in resources
          MessageBox(hDlg,"Unbekannte Dateiendung "+WChr(&H22)+szwBuffer2+WChr(&H0022,&H000D)+"Kann Dateityp nicht bestimmen.",TXT_APP_NAME,MB_ICONSTOP)
        Case Else                                         'Neutral (english)
          MessageBox(hDlg,"Unknown file suffix "+WChr(&H22)+szwBuffer2+WChr(&H0022,&H000D)+"Can't determine file type.",TXT_APP_NAME,MB_ICONSTOP)
        End Select
      End Select
    Case IDCANCEL                  'ESC key                      
      EndDialog(hDlg,0)            'Modal dialog
    Case IDCB_FILENAME             'Select filename
      GetWindowText(hwndFileName,szwBuffer,SizeOf(szwBuffer))
      tOfn.lStructSize     = SizeOf(OPENFILENAME)
      tOfn.hwndOwner       = hDlg
      Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
'      Case &H07 : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"RTF-Textdatei (*.rtf)\0*.rtf\0Winword97-Datei (*.doc)\0*.doc\0KeyNote-Datei (*.knt)\0*.knt\0Alle Dateien (*.*)\0*.*\0\0")))
'      Case Else : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"RTF text file (*.rtf)\0*.rtf\0\0Winword97 file (*.doc)\0*.kntKeyNote file (*.knt)\0*.knt\0All files (*.*)\0*.*\0\0")))
      Case &H07 : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"KeyNote-Datei (*.knt)\0*.knt\0Alle Dateien (*.*)\0*.*\0\0")))
      Case Else : tOfn.lpstrFilter = Cast(LPCWSTR,StrPtr(WStr(!"KeyNote file (*.knt)\0*.knt\0All files (*.*)\0*.*\0\0")))
      End Select
      Select Case UCase(Mid(szwBuffer,Len(szwBuffer)-2,8))
      Case "RTF","DOC"
        tOfn.nFilterIndex = 1
      Case "KNT"
        tOfn.nFilterIndex = 2
      Case Else
        tOfn.nFilterIndex = 3
      End Select
      tOfn.lpstrFile       = @szwBuffer
      tOfn.lpstrInitialDir = @szwBuffer
      tOfn.nMaxFile        = MAX_PATH
      tOfn.Flags           = OFN_PATHMUSTEXIST Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_ENABLESIZING Or OFN_OVERWRITEPROMPT
      'Ofn.FlagsEx        = OFN_EX_NOPLACESBAR   'Win2000: Don't show icons for favorits, desktop and so on
      If GetSaveFileName(@tOfn)<>0 Then  'OFN_EXTENSIONDIFFERENT
        SetWindowText(hwndFileName,szwBuffer)
        i=Len(szwBuffer)
        SendMessage(hwndFileName,EM_SETSEL,Cast(WPARAM,i),Cast(LPARAM,i))    'Set caret to end. This makes the current file extension visible without scrolling
        If SendMessage(hwndChapterListbox,LB_GETSELCOUNT,NULL,NULL)>0 Then   'At least 1 chapter selected,
          EnableWindow(hwndOK,True)
        Else
          EnableWindow(hwndOK,False)
        End If
      End If
    Case IDCL_CHAPTER,IDCE_FILENAME              'local Table of Contents copy, edit control for filename
      If HiWord(wParam)=LBN_SELCHANGE Or HiWord(wParam)=EN_CHANGE Then  'A bit ugly, but works OK
        If SendMessage(hwndChapterListbox,LB_GETSELCOUNT,NULL,NULL)>0 And SendMessage(hwndFileName,WM_GETTEXTLENGTH,NULL,NULL)>4 Then   'At least 1 chapter selected, at leat 5 characters long filename
          EnableWindow(hwndOK,True)
        Else
          EnableWindow(hwndOK,False)
        End If
      End If
    End Select
    Return True
  Case Else
    Return 0
  End Select
End Function
'***************************************************************************************************************************************
Function ExportKeyNote(ByVal hwndParent As HWND,ByVal hExportList As HWND,Byref szwFilename As Wstring) As LRESULT
  'Exports all RTF and picture files selected in multiselection listbox control hExportList to the KeýNote file szwFilename.
  'This listbox must be an exact copy of the TOC/ShadowTOC listbox pair in the man window.
  'Returns 0 if OK or an error value otherwise.
  'An example KeyNote file is at the bottom of this function.
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static i As Integer,hFile As HGLOBAL,swBuffer As Wstring*MAX_PATH,dwFileSize As Dword,szBuffer As ZString*2048,szBuffer2 As ZString*2048
  Static swBuffer2 As Wstring*MAX_PATH,dwBytesWritten As Dword,iRetVal As Integer,fSkipThisChapter As Integer
  '
  iRetVal=0    'init function return value to "success"
  '
  hFile=CreateFile(@szwFilename,GENERIC_WRITE,0,NULL,CREATE_ALWAYS,FILE_ATTRIBUTE_NORMAL,NULL)
  If hFile=INVALID_HANDLE_VALUE Then
    iRetVal=1   'Creating error
    Goto Done_ExportKeyNote
  End If
  '
  'Watch! - We use 8 bit strings for the content of the KeyNote file.
  szBuffer2=Mid(szwFilename,InStrRev(szwFilename,"\")+1)
  '
  szBuffer="#!GFKNT 2.0"+Chr(13,10)+"#^011000000000000000000000"+Chr(13,10)+"%+"+Chr(13,10)+"NN="+szBuffer2+Chr(13,10)  'Filename als KeyNote tree title, Flags: Icons, RichEdit 3 features
  If WriteFile(Cast(HANDLE,hFile),@szBuffer,Len(szBuffer),@dwBytesWritten,NULL)=0 Then
    iRetVal=2   'Writing error
    Goto Done_ExportKeyNote     'For this error the user error message is done below
  End If
  
  Static iZipArchiveIndex As Integer,pCurZipFile As Any Ptr,pCurFileData As Any Ptr,ZipStat As zip_stat_,iLevel As Integer,iFileType As Integer
  Static hCurFileData As HGLOBAL

  For i=0 To SendMessage(hExportList,LB_GETCOUNT,NULL,NULL)-1
    If SendMessage(hExportList,LB_GETSEL,Cast(WPARAM,i),NULL) Then           'If current Listbox entry selected
      SendMessage(hExportList,LB_GETTEXT,i,Cast(lParam,@swBuffer))           '
      swBuffer2=LTrim(swBuffer)
      iLevel=(Len(swBuffer)-Len(swBuffer2))\2                                'We inserted 2 spaces per indention level in function ExportDlgProc()
      iFileType=0                                                            'Default is Picture file
      szBuffer2=Mid(swBuffer2,2)                                             'Convert to 8 bit without trailing file type digit
      If *Cast(UShort Ptr,StrPtr(swBuffer2))=EXP_CHAR_RTF Then iFileType=1   'The (now) first digit in file title prefix tells the file type
      szBuffer="%-"+Chr(13,10)+"LV="+Str(iLevel)+Chr(13,10)+"ND="+szBuffer2+Chr(13,10)+"NF=000000100100000000000000"+Chr(13,10)+"%:"+Chr(13,10)  'Watch: 8-bit string 'KeyNote Flags: Expanded, Word wrap
      If WriteFile(Cast(HANDLE,hFile),@szBuffer,Len(szBuffer),@dwBytesWritten,NULL)=0 Then
        iRetVal=2   'Writing error
        Exit For
      End If
      '
      iZipArchiveIndex=SendMessage(hWndShadowToc,LB_GETITEMDATA,SendMessage(hExportList,LB_GETITEMDATA,i,0),0)  'The shadow TOC stores the true ZIP file index in user data
      zip_stat_index(pCurZipArchive,iZipArchiveIndex,ZIP_FL_ENC_GUESS,@ZipStat)  'Return value -1: Error
      '
      If ZipStat.encryption_method<>ZIP_EM_NONE And szArchivePassword="" And fDontAskPasswordForThisBook=False Then
        If DialogBox(hProgInst,MAKEINTRESOURCE(IDD_PASSWORDDLG),hWndMain,@PasswordDlgProc)=1 Then 'returns 1 if the user leaves the dialog with "OK" or FALSE otherwise
          zip_set_default_password(pCurZipArchive,@szArchivePassword) 'szArchivePassword is an 8 bit string!
        End If
      End If
      '
      fSkipThisChapter=False
      '
      pCurZipFile=zip_fopen_index(pCurZipArchive,iZipArchiveIndex,0)             '    If (pCurZipFile=NULL) Then 'Error
      If (pCurZipFile=NULL) Then
        szArchivePassword=""                                      'Erase possibly wrong password, so thet next time Password dialog appears again
        If fDontAskPasswordForThisBook=False Then
          iRetVal=3   'Open Chapter error (possibly wrong password for encrypted chapter)
          Exit For
        Else
          fSkipThisChapter=TRUE
        End If
      End If    
      '
      If fSkipThisChapter=False Then
        If (ZipStat.size>3) And (ZipStat.valid Or ZIP_STAT_SIZE) And (ZipStat.valid Or ZIP_STAT_NAME) Then
          If iFileType=1 Then 'RTF file
            pCurFileData=GlobalAlloc(GMEM_FIXED,ZipStat.size+1)
            Poke pCurFileData+ZipStat.size,0                    'NULL termination of string
            zip_fread(pCurZipFile,pCurFileData,ZipStat.size)
            If WriteFile(Cast(HANDLE,hFile),pCurFileData,ZipStat.size,@dwBytesWritten,NULL)=0 Then
              GlobalFree(pCurFileData)
              zip_fclose(pCurZipFile)
              iRetVal=2   'Writing error
              Exit For
            End If
            GlobalFree(pCurFileData)
          Else     'iFileType=2: Picture file
            hCurFileData=GlobalAlloc(GMEM_MOVEABLE,ZipStat.size+1)   'GMEM_MOVEABLE is a "must" for the function DecompressPicture()
            pCurFileData=GlobalLock(hCurFileData)
            Poke pCurFileData+ZipStat.size,0                         'NULL termination of string
            zip_fread(pCurZipFile,pCurFileData,ZipStat.size)
            GlobalUnlock(hCurFileData)
            hCurFileData=PictureToRtf(hCurFileData)                  'Decompress picture, get back RTF file in (another!) global memory
            pCurFileData=GlobalLock(hCurFileData)
            If WriteFile(Cast(HANDLE,hFile),pCurFileData,GlobalSize(hCurFileData),@dwBytesWritten,NULL)=0 Then
              GlobalFree(pCurFileData)
              zip_fclose(pCurZipFile)
              iRetVal=2   'Writing error
              Exit For
            End If
            GlobalUnlock(hCurFileData)
            GlobalFree(hCurFileData)                                 'We must free the memory buffer got from function DecompressPicture()
          End If   'File type RTF text or picture
          zip_fclose(pCurZipFile)
        Else                      'Error in zip_stat_index()
          zip_fclose(pCurZipFile)
          iRetVal=4   'UnZip error
          Exit For
        End If
        '
        If WriteFile(hFile,Cast(LPCSTR,StrPtr(Chr(13,10))),2,@dwBytesWritten,NULL)=0 Then
          iRetVal=2   'Writing error
          Exit For
        End If
      End If     'fSkipThisChapter=FALSE
    End If    'If chapter selected
  Next i
  '
  If iRetVal>0 Then Goto Done_ExportKeyNote
  '
  If WriteFile(hFile,Cast(LPCSTR,StrPtr("%%")),2,@dwBytesWritten,NULL)=0 Then
    iRetVal=2   'Writing error
    Goto Done_ExportKeyNote     'For this error the user error message is done below
  End If
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .  
  Done_ExportKeyNote:
  '
  If hFile<>NULL Then
    CloseHandle(hFile)
    hFile=NULL   'Reset for next call (variable is static!)
  End If
  '
  Select Case iRetVal
  Case 0        'Creating error
    Select Case UserSettings.wLanguageID                                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Export fertig!"+Chr(13)+szwFilename                 'German: Default language in resources
    Case Else : swBuffer="Export done!"+Chr(13)+szwFilename                   'Neutral (english)
    End Select
    MessageBox(hwndParent,swBuffer,TXT_APP_NAME,MB_ICONINFORMATION)
  Case 1        'Creating error
    Select Case UserSettings.wLanguageID                                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Konnte Datei nicht anlegen:"+Chr(13)+szwFilename    'German: Default language in resources
    Case Else : swBuffer="Couldn't create file:"+Chr(13)+szwFilename          'Neutral (english)
    End Select
    ShowSystemErrorMessage(hwndParent,swBuffer)
  Case 2       'Writing error
    Select Case UserSettings.wLanguageID                                              'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Fehler beim Schreiben in die Datei"+Chr(13)+szwFilename    'German: Default language in resources
    Case Else : swBuffer="Error while writing to file"+Chr(13)+szwFilename           'Neutral (english)
    End Select
    ShowSystemErrorMessage(hwndParent,swBuffer)
  Case 3       '3: Open Chapter error (for example password wrong for encrypted chapters)
    Select Case UserSettings.wLanguageID                                                   'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Konnte Kapitel nicht öffnen:"+Chr(13)+szBuffer2+Chr(13,13)      'German: Default language in resources
    Case Else : swBuffer="Could not open chapter:"+Chr(13)+szBuffer2+Chr(13,13)           'Neutral (english)
    End Select
    swBuffer2=swBuffer+Chr(13,10)+"("+Chr(34)+*Cast(ZString Ptr,zip_strerror(pCurZipArchive))+Chr(34)+")"
    MessageBox(hwndParent,swBuffer2,TXT_APP_NAME,MB_ICONINFORMATION)
  Case Else    '4: UnZip error
    Select Case UserSettings.wLanguageID                                               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07 : swBuffer="Fehler beim Laden des Kapitels"+Chr(13)+Mid(swBuffer2,2)    'German: Default language in resources
    Case Else : swBuffer="Error while loading chapter"+Chr(13)+Mid(swBuffer2,2)       'Neutral (english)
    End Select
    ShowSystemErrorMessage(hwndParent,swBuffer)
  End Select
  '
  Return iRetVal
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'Below is an example keyNote *.knt file. "NN": Title of the Tab (central knot). "LV": Level. Let file header as it is.
  'The KeyNote file is a plain text file, one entry each line. RTF may contain more lines. Keynote format coninues after logical end of each RTF.
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .  
  '#!GFKNT 2.0
  '%+
  'NN=Collection of RTFs
  '%-
  'LV=0
  'ND=Eintrag 1 auf Level 0
  '%:
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\lang1031\f0\b\fs20 Hallo erstes Kapitel\b0\par}
  '%-
  'LV=0
  'ND=Eintrag 2 auf Level 0
  '%:
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\lang1031\f0\b\fs20 Hallo zweites Kapitel\b0\par}
  '%-
  'LV=1
  'ND=Eintrag 3 auf Level 1
  '%:
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\lang1031\f0\b\fs20 Hallo drittes Kapitel\b0\par}
  '%-
  'LV=1
  'ND=Eintrag 4 auf Level 1
  '%:
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\lang1031\f0\b\fs20 Hallo viertes Kapitel\b0\par}
  '%-
  'LV=1
  'ND=Eintrag 5 auf Level 0
  '%:
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\lang1031\f0\b\fs20 Hallo fünftes Kapitel\b0\par}
  '%%
End Function
'***************************************************************************************************************************************
Function SettingsDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'Standard dialog function for the modal "Settings" Dialog box (IDD_SETTINGS).
  'Used global vartiables: 
'IDC_GRP_LANGUAGE 1001
'IDC_LANGAUTO 1002
'IDC_LANGENGL 1003
'IDC_LANGERMAN 1004
'IDC_GRP_ZOOM 1005
'IDC_ZOOMAUTO 1006
'IDC_ZOOMNORMAL 1007
'IDC_ZOOMBIGGER 1008
'IDC_GRP_SETTINGS 1009
'IDC_SETTINGSVOLATILE 1010
'IDC_USEINIFILE 1011
'IDC_USEREGISTRY 1012
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static i As Long,j As Long,szwBuffer As Wstring*4096,lvItem As LVITEM,hwndCbbFirstDay As HWND,hwndCbbLastDay As HWND ',ulToday As Ulong
  '
  Select Case(uMsg)
  Case WM_INITDIALOG
    '.  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .
'    If fNoBackups<>0 Then CheckDlgButton(hDlg,IDC_NOBACKUPS,BST_CHECKED)
'    hwndCbbFirstDay = GetDlgItem(hDlg,IDC_FIRSTDAY)        'Combobox "First day"
'    hwndCbbLastDay  = GetDlgItem(hDlg,IDC_LASTDAY)         'Combobox "Last day"
    '
     '.  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .
    Return TRUE        'TRUE: Let the system set the focus
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDOK
      'If IsDlgButtonChecked(hDlg,IDC_NOBACKUPS)=BST_CHECKED Then fNoBackups=True Else fNoBackups=FALSE
      
      EndDialog(hDlg,0)            'Modal dialog
    Case IDCANCEL                  'ESC key                      
      EndDialog(hDlg,0)            'Modal dialog
    End Select
  Case Else
    Return 0
  End Select
End Function
'***************************************************************************************************************************************
Function SearchDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'DialogBox function for the Search window
  'ToDo: Search in metadata,
  '      RadioButton: Search in current book/Search in other books: Select home directory of other books.
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  #define SEARCH_UP_BOOK 0    'ID in "Search strategy" combobox
  #define SEARCH_UP_CHAPTER 1
  #define SEARCH_DOWN_CHAPTER 2
  #define SEARCH_DOWN_BOOK 3
  #define SEARCH_BOOK_BEGIN 4
  #define SEARCH_CHAPTER_BEGIN 5
  '
  Static tCurWindowRect As RECT,hWndStrategy As HWND,hwndSearchTerm As HWND,hwndOkButton As HWND,sSearchTerm As Wstring*256,iCurChapter As Integer
  Static dwCurSearchFlags As Dword,tFindTextEx As FINDTEXTEX,tCharRange As CHARRANGE,tCharRangeLast As CHARRANGE,tCharFormat As CHARFORMAT
  Static iSearchStrategy As Integer,fSearchDialogInitIsDone As Integer,fMatchCase As Integer,fWholeWord As Integer,fSearchFormattings As Integer
  Static fStopSearching As Integer,iFindPos As Integer
  '
  Select Case uMsg
  Case WM_INITDIALOG
    hWndStrategy=GetDlgItem(hDlg,IDC_SEARCH_STRATEGY)
    hwndSearchTerm=GetDlgItem(hDlg,IDC_SEARCHTERM)
    hwndOkButton=GetDlgItem(hDlg,IDOK)
    '
    Select Case UserSettings.wLanguageID               'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07                                         'German: Default language in resources
      'For the "Search strategy" combobox don't make changes without changing the ID definitions!
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Suche nach oben im ganzen Buch"))))      'ID: SEARCH_UP_BOOK
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Suche nach oben im Kapitel"))))          'ID: SEARCH_UP_CHAPTER
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Suche nach unten im Kapitel"))))         'ID: SEARCH_DOWN_CHAPTER
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Suche nach unten im ganzen Buch"))))     'ID: SEARCH_DOWN_BOOK
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Suche im ganzen Buch von Anfang an"))))  'ID: SEARCH_BOOK_BEGIN
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Suche im Kapitel von Anfang an"))))      'ID: SEARCH_CHAPTER_BEGIN
      SendMessage(hWndStrategy,CB_SETCURSEL,iSearchStrategy,NULL)
    Case Else                                         'Neutral (english)
      SetWindowText(hDlg,"Search...")
      SetDlgItemText(hDlg,IDCS_SEARCH_FOR,"Search &for:")
      SetDlgItemText(hDlg,IDCS_SEARCH_STRATEGY,"&Kind of search:")
      SetDlgItemText(hDlg,IDC_WHOLEWORD,"&Whole word only")
      SetDlgItemText(hDlg,IDC_MATCHCASE,"&Match case")
      SetDlgItemText(hDlg,IDC_COLOR_DARKGREY,"&Dark gray")
      SetDlgItemText(hDlg,IDOK,"&Search again")
      SetDlgItemText(hDlg,IDCANCEL,"&Cancel")
      '
      'For the "Search strategy" combobox don't make changes without changing the ID definitions!
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Search up to top of book"))))    'ID: SEARCH_UP_BOOK
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Search up in chapter"))))        'ID: SEARCH_UP_CHAPTER
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Search down in chapter"))))      'ID: SEARCH_DOWN_CHAPTER
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Search down in whole book"))))   'ID: SEARCH_DOWN_BOOK
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Search whole book from top"))))  'ID: SEARCH_BOOK_BEGIN
      SendMessage(hWndStrategy,CB_ADDSTRING,NULL,Cast(LPARAM,StrPtr(WStr("Search chapter from top"))))     'ID: SEARCH_CHAPTER_BEGIN
      SendMessage(hWndStrategy,CB_SETCURSEL,iSearchStrategy,NULL)
    End Select
    '
    If fSearchDialogInitIsDone=True Then            'Repeated call of Search window: Restore position&size
      MoveWindow(hDlg,tCurWindowRect.left,tCurWindowRect.top,tCurWindowRect.right-tCurWindowRect.left,tCurWindowRect.bottom-tCurWindowRect.top,True)
    Else
      fSearchDialogInitIsDone=True
      iSearchStrategy=SEARCH_DOWN_CHAPTER
    End If
    '
    SetWindowText(hwndSearchTerm,@sSearchTerm)
    SendDlgItemMessage(hDlg,IDC_MATCHCASE,BM_SETCHECK,fMatchCase,NULL)
    SendDlgItemMessage(hDlg,IDC_WHOLEWORD,BM_SETCHECK,fWholeWord,NULL)
    SendDlgItemMessage(hDlg,IDC_COLOR_DARKGREY,BM_SETCHECK,fSearchFormattings,NULL)
    '
    Function=1                                                                   '0: We set the focus, 1: Let the system set the focus
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDOK
      fStopSearching=False
      dwCurSearchFlags=0
      If SendDlgItemMessage(hDlg,IDC_MATCHCASE,BM_GETCHECK,NULL,NULL) Then dwCurSearchFlags=dwCurSearchFlags Or FR_MATCHCASE
      If SendDlgItemMessage(hDlg,IDC_WHOLEWORD,BM_GETCHECK,NULL,NULL) Then dwCurSearchFlags=dwCurSearchFlags Or FR_WHOLEWORD
      If SendDlgItemMessage(hDlg,IDC_COLOR_DARKGREY,BM_GETCHECK,NULL,NULL) Then fSearchFormattings=TRUE Else fSearchFormattings=FALSE
      GetWindowText(hwndSearchTerm,@sSearchTerm,SizeOf(sSearchTerm))
      ZeroMemory(@tFindTextEx,SizeOf(FINDTEXTEX))
      iCurChapter=SendMessage(hwndTOC,LB_GETCURSEL,NULL,NULL)
      iSearchStrategy=SendMessage(hWndStrategy,CB_GETCURSEL,NULL,NULL)
      '
      Select Case iSearchStrategy
      Case SEARCH_BOOK_BEGIN                'Search through chapters and texts upwards. Start at very first chapter at text position 0.
        fStopSearching=FALSE
        iCurChapter=0
        LoadBookChapter(iCurChapter)
        tFindTextEx.chrg.cpMin=0
        tFindTextEx.chrg.cpMax=-1
        tFindTextEx.lpstrText=@sSearchTerm
        dwCurSearchFlags=dwCurSearchFlags Or FR_DOWN
        iFindPos=SendMessage(hWndText,EM_FINDTEXTEXW,dwCurSearchFlags,Cast(LPARAM,@tFindTextEx))
        If iFindPos<>-1 Then                                                                        'If found
          If iCurChapter<SendMessage(hWndToc,LB_GETCOUNT,NULL,NULL)-1 Then                            'Not just in the highest chapter
            SendMessage(hWndStrategy,CB_SETCURSEL,SEARCH_DOWN_BOOK,NULL)                            'Start next search from current position, not from begin of book
            iCurChapter+=1
            tCharRangeLast.cpMin=0:tCharRangeLast.cpMax=0                                           'For cese of searching formattings
            LoadBookChapter(iCurChapter)                                                            'Hint: Results automaticly text position 0 in new chapter
            PostMessage(hDlg,WM_COMMAND,IDOK,NULL)                                                  'Continue with strategy SEARCH_UP_BOOK, not SEARCH_BOOK_BEGIN
          End If
        Else                                                                                        'Not found in current chapter
          SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))                 'Select found text
          SendMessage(hWndStrategy,CB_SETCURSEL,SEARCH_DOWN_BOOK,NULL)                              'Change search strategy to SEARCH_UP_BOOK for next "Search again" click
        End If
      Case SEARCH_DOWN_BOOK         'Search through chapters and texts downwards
        SendMessage(hWndText,EM_EXGETSEL,NULL,Cast(LPARAM,@tCharRange))                             'Returns  tCharRange.cpMin and tCharRange.cpMax
        tFindTextEx.chrg.cpMin=tCharRange.cpMax
        tFindTextEx.chrg.cpMax=-1
        tFindTextEx.lpstrText=@sSearchTerm
        dwCurSearchFlags=dwCurSearchFlags Or FR_DOWN
        iFindPos=SendMessage(hWndText,EM_FINDTEXTEXW,dwCurSearchFlags,Cast(LPARAM,@tFindTextEx))
        If iFindPos<>-1 Then                                                                        'If found
          SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))                 'Select found text
        Else                                                                                        'Found in current chapter
          If iCurChapter<SendMessage(hWndToc,LB_GETCOUNT,NULL,NULL)-1 Then                            'Not just in the highest chapter
            iCurChapter+=1
            tCharRangeLast.cpMin=0:tCharRangeLast.cpMax=0                                           'For cese of searching formattings
            LoadBookChapter(iCurChapter)                                                            'Hint: Results automaticly text position 0 in new chapter
            PostMessage(hDlg,WM_COMMAND,IDOK,NULL)                                                  'Continue with strategy SEARCH_UP_BOOK, not SEARCH_BOOK_BEGIN
          End If
        End If
      Case SEARCH_UP_BOOK       'Search through chapters and texts upwards. 
        SendMessage(hWndText,EM_EXGETSEL,NULL,Cast(LPARAM,@tCharRange))                             'Returns  tCharRange.cpMin and tCharRange.cpMax
        tFindTextEx.chrg.cpMin=tCharRange.cpMin
        tFindTextEx.chrg.cpMax=-1
        tFindTextEx.lpstrText=@sSearchTerm
        iFindPos=SendMessage(hWndText,EM_FINDTEXTEXW,dwCurSearchFlags,Cast(LPARAM,@tFindTextEx))
        If iFindPos<>-1 Then                                                                        'If found
          SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))                 'Select found text
        Else                                                                                        'Found in current chapter
          If iCurChapter>0 Then
            iCurChapter-=1
            tCharRangeLast.cpMin=0:tCharRangeLast.cpMax=0                                           'For cese of searching formattings
            LoadBookChapter(iCurChapter)                                                            'Hint: Results automaticly text position 0 in new chapter
            SendMessage(hWndText,EM_SETSEL,Cast(wParam,-1),Cast(lParam,-1))                         'Set to end (!) of text
            PostMessage(hDlg,WM_COMMAND,IDOK,NULL)                                                  'Continue with strategy SEARCH_UP_BOOK, not SEARCH_BOOK_BEGIN
          End If
        End If
      Case SEARCH_CHAPTER_BEGIN
        tFindTextEx.chrg.cpMin=0
        tFindTextEx.chrg.cpMax=-1
        tFindTextEx.lpstrText=@sSearchTerm
        dwCurSearchFlags=dwCurSearchFlags Or FR_DOWN
        iFindPos=SendMessage(hWndText,EM_FINDTEXTEXW,dwCurSearchFlags,Cast(LPARAM,@tFindTextEx))
        If iFindPos<>-1 Then                                                                        'If found
          SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))
        Else                                                                                        'Not found
          fStopSearching=True                                                                       'Disables formatting check below
        End If
        SendMessage(hWndStrategy,CB_SETCURSEL,SEARCH_DOWN_CHAPTER,NULL)                             'Next "Search again" starts from new current position, not from begin of chapter
      Case SEARCH_UP_CHAPTER
        SendMessage(hWndText,EM_EXGETSEL,NULL,Cast(LPARAM,@tCharRange))                             'Returns  tCharRange.cpMin and tCharRange.cpMax
        tFindTextEx.chrg.cpMin=tCharRange.cpMin
        tFindTextEx.chrg.cpMax=-1
        tFindTextEx.lpstrText=@sSearchTerm
        iFindPos=SendMessage(hWndText,EM_FINDTEXTEXW,dwCurSearchFlags,Cast(LPARAM,@tFindTextEx))
        If iFindPos<>-1 Then                                                                        'If found
          SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))
        Else
          fStopSearching=True                                                                       'Disables formatting check below
        End If
      Case Else  'SEARCH_DOWN_CHAPTER (2) is default
        SendMessage(hWndText,EM_EXGETSEL,NULL,Cast(LPARAM,@tCharRange))                             'Returns  tCharRange.cpMin and tCharRange.cpMax
        tFindTextEx.chrg.cpMin=tCharRange.cpMax
        tFindTextEx.chrg.cpMax=-1
        tFindTextEx.lpstrText=@sSearchTerm
        dwCurSearchFlags=dwCurSearchFlags Or FR_DOWN
        iFindPos=SendMessage(hWndText,EM_FINDTEXTEXW,dwCurSearchFlags,Cast(LPARAM,@tFindTextEx))
        If iFindPos<>-1 Then                                                                        'If found                                                                        'If found
          SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tFindTextEx.chrgText))
        Else
          fStopSearching=True                                                                       'Disables formatting check below
        End If
      End Select
      '
      If fSearchFormattings=True And iFindPos<>-1 Then                                              'Search for formattings enabled
        If (tCharRangeLast.cpMin<>0) And (tCharRangeLast.cpMax<>0) Then
          tCharFormat.cbSize=SizeOf(tCharFormat)
          tCharFormat.dwMask=CFM_COLOR
          SendMessage(hWndText,EM_GETCHARFORMAT,1,Cast(LPARAM,@tCharFormat))                        'Get character formatting
          If tCharFormat.crTextColor=COLOR_DARKGRAY Then                                            'Of course it whould be simple to add more formattings here, but why? ;-)
            tCharRangeLast=tFindTextEx.chrgText                                                     'Get current selection to jump back if last formatting match
          Else
            If fStopSearching=True Then
              SendMessage(hWndText,EM_EXSETSEL,NULL,Cast(LPARAM,@tCharRangeLast))                   'Restore last match with searched formattings   
            Else 
              PostMessage(hDlg,WM_COMMAND,IDOK,NULL)                                                'Not the searched color: Search next text match
            End If
          End If
        End If
      End If
    Case IDCANCEL                                                                'ESC key / Alt+F4
      GetWindowText(hwndSearchTerm,@sSearchTerm,SizeOf(sSearchTerm))
      iSearchStrategy=SendMessage(hWndStrategy,CB_GETCURSEL,NULL,NULL)
      fMatchCase=SendDlgItemMessage(hDlg,IDC_MATCHCASE,BM_GETCHECK,NULL,NULL)
      fWholeWord=SendDlgItemMessage(hDlg,IDC_WHOLEWORD,BM_GETCHECK,NULL,NULL)
      fSearchFormattings=SendDlgItemMessage(hDlg,IDC_COLOR_DARKGREY,BM_GETCHECK,NULL,NULL)
      GetWindowRect(hDlg,@tCurWindowRect)                                        'Get and store window size and position to restore next time
      DestroyWindow(hDlg)                                                        'Non-modal window! - call EndDialog() if created as a modal window
      hDlgSearch=NULL                                                            'Invalidate global variable
      SetFocus(hwndMain)                                                         'Sometimes useful because the help window is non-modal
    Case IDC_SEARCHTERM
      If HiWord(wParam)=EN_CHANGE Then
        If SendMessage(hwndSearchTerm,WM_GETTEXTLENGTH,NULL,NULL) Then
          EnableWindow(hwndOkButton,TRUE)
        Else
          EnableWindow(hwndOkButton,FALSE)
        End If
      End If
    End Select
    Function=0
  Case Else
    Function=False
  End Select
End Function
'***************************************************************************************************************************************
Function PasswordDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'DialogBox function for the Password input window
  'Changes (or not) the global variable szArchivePassword (always an 8-bit string, not Unicode)
  'Hint: You can't by Windows designenter Double Byte Char in Edit Ctrl w/ Password Mask.
  'The functrion DialogBox() returns TRUE if the user leaves the dialog with "OK" or FALSE otherwise
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static hWndPasswordText As HWND,hWndShowPasswordBtn As HWND,hWndDontAskPasswordAgainBtn As HWND,szwBuffer As WString*MAX_PATH
  '
  Select Case(uMsg)
  Case WM_INITDIALOG
    hWndPasswordText=GetDlgItem(hDlg,IDCE_PASSWORD)
    hWndShowPasswordBtn=GetDlgItem(hDlg,IDCRB_SHOWPASSWORD)
    hWndDontAskPasswordAgainBtn=GetDlgItem(hDlg,IDCRB_DONTASKPASSWORD)
    '
    Select Case UserSettings.wLanguageID                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07                                                 'German: Default language in resources
    Case Else                                                 'Neutral (english)
      SetWindowText(hDlg,"Password:")
      SetWindowText(hWndShowPasswordBtn,"&Show password")
      SetWindowText(hWndShowPasswordBtn,"&Don't ask for password again")
      SetDlgItemText(hDlg,IDOK,"&OK")
      SetDlgItemText(hDlg,IDCANCEL,"&Cancel")
    End Select
    '
    If UserSettings.fShowPassword=BST_CHECKED Then
      SendMessage(hWndShowPasswordBtn,BM_SETCHECK,BST_CHECKED,NULL)
    Else
      SendMessage(hWndPasswordText,EM_SETPASSWORDCHAR,42,NULL)
    End If
    If fDontAskPasswordForThisBook=BST_CHECKED Then SendMessage(hWndDontAskPasswordAgainBtn,BM_SETCHECK,BST_CHECKED,NULL)
    SetWindowText(hWndPasswordText,szArchivePassword)
     '.  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .
    Return TRUE        'TRUE: Let the system set the focus
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDOK
      UserSettings.fShowPassword=SendMessage(hWndShowPasswordBtn,BM_GETCHECK,NULL,NULL)
      fDontAskPasswordForThisBook=SendMessage(hWndDontAskPasswordAgainBtn,BM_GETCHECK,NULL,NULL)
      GetWindowText(hWndPasswordText,@szwBuffer,SizeOf(szwBuffer))                        'Watch: Edit window is Unicode
      szArchivePassword=szwBuffer
      EndDialog(hDlg,1)
    Case IDCANCEL                  'ESC key, Cancel, Alt+F4...                  
      EndDialog(hDlg,0)
    Case IDCRB_SHOWPASSWORD
     If HiWord(wParam)=BN_CLICKED Then
        If SendMessage(hWndShowPasswordBtn,BM_GETCHECK,NULL,NULL)=BST_CHECKED Then
          SendMessage(hWndPasswordText,EM_SETPASSWORDCHAR,NULL,NULL)                   'No masking character
        Else
          SendMessage(hWndPasswordText,EM_SETPASSWORDCHAR,42,NULL)                     'Asterisk=Chr(42) (sets ES_PASSORD flag as a result)
        End If
        InvalidateRect(hWndPasswordText,NULL,True)                                     'Casues a redraw
      End If
    End Select
  Case Else
    Return 0
  End Select
End Function
'***************************************************************************************************************************************
Function HelpDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'DialogBox function for the help window with RichEdit control.
  '
  'Used global variables:
  'Dim Shared swAppFile As wString*MAX_PATH                 'Current (real) path and filename of application
  'Dim Shared swAppName As wString*MAX_PATH                 'Application's name
  '  UserSettings.HelpWinX
  '  UserSettings.HelpWinY
  '  UserSettings.HelpWinWidth
  '  UserSettings.HelpWinHeight 
  '  UserSettings.dwHelpCurScrollPos
  '  UserSettings.dwHelpTextPos
  '
  '#define IDR_HELPTEXT_EN 1
  '#define IDR_HELPTEXT_DE 2
  '
  'Idea for later: Keyboard shortcuts for each tab control item. This isn't possible in a direct way.
  '                Possible solution for TAB control shortcuts: WM_MENUCHAR chUser="ü" (252), fuFlag=MF_SYSMENU
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static hResInfo As HRSRC,hResData As HGLOBAL,pbResource As Byte Ptr,tSetTextEx As SETTEXTEX,i As Long,hWndHlpTxt As HWND,hwndHlpTab As HWND
  Static swzBuffer As WString*2048,dwVersionInfoSize As Dword,pVerInfo As LPVOID,pswFileInfo As wString Ptr,pt As Point,wplc As WINDOWPLACEMENT
  Static pNmHdr As NMHDR Ptr,pEnLink As ENLINK Ptr,tTextRange As TEXTRANGE,tFindTextExInternal As FINDTEXTEX,rct As RECT,tTcItem As TCITEM
  '
  Select Case uMsg
  Case WM_INITDIALOG
    hWndHlpTxt=GetDlgItem(hDlg,IDC_HELPTEXT)
    hwndHlpTab=GetDlgItem(hDlg,IDC_HELPTAB)
    '
    tTcItem.mask=TCIF_TEXT
    tTcItem.pszText=Cast(LPTSTR,StrPtr(WStr(!"History\0")))
    SendMessage(hwndHlpTab,TCM_INSERTITEM,0,cast(lParam,@tTcItem))
    tTcItem.pszText=Cast(LPTSTR,StrPtr(WStr(!"Tips...\0")))
    SendMessage(hwndHlpTab,TCM_INSERTITEM,0,cast(lParam,@tTcItem))
    Select Case UserSettings.wLanguageID                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
    Case &H07                                                 'German: Default language in resources
      tTcItem.pszText=Cast(LPTSTR,StrPtr(WStr(!"RTF books selber machen\0")))
      SendMessage(hwndHlpTab,TCM_INSERTITEM,0,cast(lParam,@tTcItem))
      tTcItem.pszText=Cast(LPTSTR,StrPtr(WStr(!"Über...\0")))
      SendMessage(hwndHlpTab,TCM_INSERTITEM,0,cast(lParam,@tTcItem))
    Case Else                                                 'Neutral (english)
      tTcItem.pszText=Cast(LPTSTR,StrPtr(WStr(!"Create own RTF books\0")))
      SendMessage(hwndHlpTab,TCM_INSERTITEM,0,cast(lParam,@tTcItem))
      tTcItem.pszText=Cast(LPTSTR,StrPtr(WStr(!"About...\0")))
      SendMessage(hwndHlpTab,TCM_INSERTITEM,0,Cast(lParam,@tTcItem))
    End Select
    '
    If UserSettings.dwCurHelpChapter>3 Then UserSettings.dwCurHelpChapter=0                           'Prevent errors, we have only 4 help chapters
    SendMessage(hwndHlpTab,TCM_SETCURSEL,UserSettings.dwCurHelpChapter,NULL)                          'Select last help Tab
    '
    SendMessage(hWndHlpTxt,EM_AUTOURLDETECT,1,0)                                                      'Enable Auto URL recognition 
    SendMessage(hWndHlpTxt,EM_SETEVENTMASK,0,ENM_LINK)                                                'We need events for Link clicks
    '
    'Currently not restored: UserSettings.HelpWinState=SW_SHOWMAXIMIZED/SW_SHOWNORMAL (because SW_SHOWNORMAL can't be changed by user momently)
    '
    If UserSettings.HelpWinX<0 Then UserSettings.HelpWinX=GetSystemMetrics(SM_CXSCREEN)\2
    If UserSettings.HelpWinX>GetSystemMetrics(SM_CXSCREEN) Then UserSettings.HelpWinX=GetSystemMetrics(SM_CXSCREEN)\2
    If UserSettings.HelpWinY<0 Then UserSettings.HelpWinY=0
    If UserSettings.HelpWinY>GetSystemMetrics(SM_CYSCREEN)-20 Then UserSettings.HelpWinY=0
    If UserSettings.HelpWinWidth<10 Then UserSettings.HelpWinWidth=GetSystemMetrics(SM_CXSCREEN)\2
    If UserSettings.HelpWinWidth>GetSystemMetrics(SM_CXSCREEN) Then UserSettings.HelpWinWidth=GetSystemMetrics(SM_CXSCREEN)\2
    If UserSettings.HelpWinHeight<20 Then UserSettings.HelpWinHeight=GetSystemMetrics(SM_CYSCREEN)
    If UserSettings.HelpWinHeight>GetSystemMetrics(SM_CYSCREEN)-20 Then UserSettings.HelpWinHeight=GetSystemMetrics(SM_CYSCREEN)
    MoveWindow(hDlg,UserSettings.HelpWinX,UserSettings.HelpWinY,UserSettings.HelpWinWidth,UserSettings.HelpWinHeight,TRUE)
    '
    'Load the VersionInfo from resource. This is a bit ... fussiness:
    dwVersionInfoSize=GetFileVersionInfoSize(swAppFile,@i)
    pVerInfo=HeapAlloc(GetProcessHeap(),HEAP_ZERO_MEMORY,dwVersionInfoSize)
    GetFileVersionInfo(swAppFile,NULL,dwVersionInfoSize,pVerInfo)
    If VerQueryValue(pVerInfo,"\StringFileInfo\040904B0\FileVersion",@pswFileInfo,@dwVersionInfoSize)<>0 Then
      SetWindowText(hDlg,swAppName+" (Version: "+*pswFileInfo+")")
    Else 'This error should never occur, because what we set into the resource an what not. Nut however ... nobody knows ;-)
      SetWindowText(hDlg,swAppName)
    End If
    '
    PostMessage(hDlg,WM_COMMAND,1000,NULL)                                           'Load the text from resource ("push" virtual button)
    SetFocus(hWndHlpTxt)
    '
    Function=0                                                                       '0: We set the focus
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDOK,IDCANCEL                                                               'ESC key / Alt+F4
      UserSettings.dwCurHelpChapter=SendMessage(hwndHlpTab,TCM_GETCURSEL,NULL,NULL)  'Get selected help Tab and store for next call of help window

      'Store current window state:
      wplc.length=SizeOf(WINDOWPLACEMENT) 
      GetWindowPlacement(hDlg,@wplc)
      If wplc.showCmd<>SW_SHOWMAXIMIZED Then wplc.showCmd=SW_SHOWNORMAL
      UserSettings.HelpWinState=wplc.showCmd                                          'State of the main window: Only SW_SHOWMAXIMIZED or SW_SHOWNORMAL is allowed
      '
      GetWindowRect(hDlg,@rct)                                                        'Get Window rectangle to store the values in ini file
      UserSettings.HelpWinX      = rct.left
      UserSettings.HelpWinY      = rct.top
      UserSettings.HelpWinWidth  = rct.right-rct.left
      UserSettings.HelpWinHeight = rct.bottom-rct.top
      '
      UserSettings.dwHelpTextPos=SendMessage(hWndHlpTxt,EM_LINEINDEX,Cast(wParam,SendMessage(hWndHlpTxt,EM_GETFIRSTVISIBLELINE,0,0)),0) 'For next call of help text
      SendMessage(hWndHlpTxt,EM_GETSCROLLPOS,0,Cast(lparam,@pt))
      UserSettings.dwHelpScrollPos=pt.y
      DestroyWindow(hDlg)                                                              'Non-modal window! - call EndDialog() if created as a modal window
      hDlgHElp=NULL                                                                    'Invalidate global variable
      SetFocus(hwndMain)                                                               'Sometimes useful because the help window is non-modal
    Case 1000                                                      'Virtual "control". Target for EM_COMMAND posting only. Load RTF text from resource.
      Select Case UserSettings.dwCurHelpChapter
      Case 0                                                       '0: Tab0 "Program"
        Select Case UserSettings.wLanguageID                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07 : i=IDR_HELPPROGRAM_DE                           'German: Default language in resources
        Case Else : i=IDR_HELPPROGRAM_EN                           'Neutral (english)
        End Select
      Case 1                                                       '1: Tab1: Make own RTF books"
        Select Case UserSettings.wLanguageID                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07 : i=IDR_HELPBOOKSMAKE_DE                         'German: Default language in resources
        Case Else : i=IDR_HELPBOOKSMAKE_EN                         'Neutral (english)
        End Select
      Case 2                                                       '2: Tab2: Tips"
        Select Case UserSettings.wLanguageID                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07 : i=IDR_HELPTIPS_DE                              'German: Default language in resources
        Case Else : i=IDR_HELPTIPS_EN                              'Neutral (english)
        End Select
      Case Else                                                     '2: Tab3: History"
        i=IDR_HELP_HISTORY                                          'Available in one language only (english)
      End Select
      '
      hResInfo=FindResource(hProgInst,Cast(LPCTSTR,i),RT_RCDATA)                                         'Resource ID of the RTF help text
      hResData=LoadResource(hProgInst,hResInfo)                                                          'Load the RTF text from resource
      pbResource=LockResource(hResData)                                                                  'Get pointer
      tSetTextEx.flags=ST_DEFAULT
      tSetTextEx.codepage=CP_ACP
      SendMessage(hWndHlpTxt,EM_SETTEXTEX,Cast(wParam,@tSetTextEx),Cast(lParam,pbResource))              'Copy help text into RichEdit help control  
      '
      SendMessage(hWndHlpTxt,EM_SETSEL,UserSettings.dwHelpTextPos,UserSettings.dwHelpTextPos)            'Remove selection, scroll to begin of text or last position
      pt.x=0
      pt.y=UserSettings.dwHelpScrollPos
      SendMessage(hWndHlpTxt,EM_SETSCROLLPOS,0,Cast(lParam,@pt))                                         'Restore croll position, if opened multi times
    End Select
    Function=0
  Case WM_GETDLGCODE
    Function=1
  Case WM_NOTIFY                                                          'TAB control, Richedit control...
    Dim pNmHdr As NMHDR Ptr=Cast(NMHDR Ptr,lParam)
    Select Case pNmHdr->idFrom
    Case IDC_HELPTAB                                                      'Notification from the TAB control
      If pNmHdr->code=TCN_SELCHANGE Then                                  'Another Tab selected
        UserSettings.dwCurHelpChapter=SendMessage(pNmHdr->hwndFrom,TCM_GETCURSEL,0,0)
        UserSettings.dwHelpTextPos=0
        UserSettings.dwHelpScrollPos=0
        PostMessage(hDlg,WM_COMMAND,1000,NULL)                             'Load the text from resource ("push" virtual button)
      End If
      Function=0
    Case IDC_HELPTEXT                                                      'Notification from the RichEdit control
      pNmHdr=Cast(NMHDR Ptr,lParam)
      i=0                                                                  'In i we store the default message return value for this branch (decides important further processing of the message)
      If pNmHdr->code=EN_LINK Then                                         'Click at a Hyperlink
        pEnLink=Cast(ENLINK Ptr,lParam)                                    'In this case the NMHDR type is enhanced to an ENLINK type
        If pEnLink->Msg=WM_LBUTTONDOWN Then
          tTextRange.chrg=pEnLink->chrg
          tTextRange.lpstrText=VarPtr(swzBuffer)
          If SendMessage(hWndHlpTxt,EM_GETTEXTRANGE,0,Cast(lParam,@tTextRange))>0 Then  'Get the link text
            ShellExecute(NULL,"open",swzBuffer,"","",SW_SHOW)                           'Assume, this is a web page or else
          End If                                                            'If SendMessage(hWndHlpTxt,EM_GETTEXTRANGE,0,Cast(lParam,@tTextRange))>0 Then
        End If                                                              'If pEnLink->Msg=WM_LBUTTONDOWN Then
      End If                                                                'If pNmHdr->code=EN_LINK Then 
      Function=i
    End Select
  Case WM_SIZE
    GetWindowRect(hWndHlpTab,@rct)                                          'We need the height of the TAB control
    i=Rct.bottom-Rct.top
    MoveWindow(hwndHlpTab,0,0,LoWord(lParam),i,TRUE)                        'Left bound, full width of Dialog client area
    MoveWindow(hWndHlpTxt,0,i+1,LoWord(lParam),HiWord(lParam)-i+1,TRUE)     'Left bound, full width Dialog client area, remaining space below the Tab control
    Function=0
  Case Else
    Function=False
  End Select
End Function
'***************************************************************************************************************************************
Function PictureToRtf(ByVal hInBuffer As HGLOBAL) As HGLOBAL
  'Decompresses a JPEG/GIF/PNG/TIFF/BMP file into a global memory object containing a standard RTF file with the decompressed picture.
  'The RTF file is compatible wih at least Wordpad (based on msftedit.dll v4.1), LibreOffice, Winword 2000, Papyrus.
  'The input memory object must be created using function GlobalAlloc(GMEM_MOVEABLE,...). For decompression GDIplus is used.
  'Altough this function is able to handle DIBs it is much faster, not to use this function for DIBs, but processing them directly.
  'The function returns 0 if error or the handle of a global memory object containing the RTF file otherwise.
  'The calling function must not free the input memory object hInBuffer, but *must* free the returned memory object if not longer used.
  'The function keeps the color deep of 1/4/8 bit color-indexed pictures for speed and size reason and to prevent possible conversion
  'problems. The code looks a bit complicated ... ;-) but the used "flat" way made the code more efficient (we handle a lot of data).
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms533971(v=vs.85).aspx
  'See: https://msdn.microsoft.com/en-us/library/windows/desktop/ms534041(v=vs.85).aspx
  Static pImageStream As IStream Ptr,GdiPlusStartupInput As GdiPlusStartupInput,pGdiPlusToken As ULONG_PTR,pImage As Any Ptr
  Static lImageWidth As Long,lImageHeight As Long,dwPixelFormat As Dword,bih As BITMAPINFOHEADER,pPalette As COLORPALETTE Ptr,dwLenBitmapBits As Dword
  Static dwSizePalette As Dword,BmpData As BITMAPDATA,dwStrideLength As Dword,dwSizeOfPalette As Dword,dwPaddingValue As Dword,sPalette As String
  Static dwSizeOfRtfFile As Dword,sRtfStart As String,sRtfEnd As String,pbDIB As Byte Ptr,dwEndOfScanline As Dword
  Static pCurInByte As UByte Ptr,pCurInWord As UShort Ptr,pCurOutByte As UByte Ptr,pCurOutWord As UShort Ptr,pszCurOut As ZString Ptr
  ' . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  pbDIB=NULL                                                              'Initial function return value
  
  If CreateStreamOnHGlobal(hInBuffer,TRUE,@pImageStream)<>S_OK Then       'Create a stream in global memory from the JPEG/PNG file in hMemStreamObject
    GoTo Done_DecompressPicture
  End If
  '
  GdiPlusStartupInput.GdiplusVersion=1
  If GdiplusStartup(@pGdiPlusToken,@GdiPlusStartupInput,NULL)<>S_OK Then
    GoTo Done_DecompressPicture
  End If
  '
  If GdipCreateBitmapFromStream(pImageStream,@pImage)<>S_OK Then          'Main job of picture conversion
    GoTo Done_DecompressPicture
  End If
  '
  GdipGetImageWidth(pImage,@lImageWidth)
  GdipGetImageHeight(pImage,@lImageHeight)
  GdipGetImagePixelFormat(pImage,@dwPixelFormat)
  pImageStream=NULL                                                        'NULL normally:  "Nothing"
  '
  ZeroMemory(@bih,SizeOf(bih))
  bih.biWidth       = lImageWidth                                          'HIMETRIC units: .01 mm = 2540/inch
  bih.biHeight      = lImageHeight
  bih.biPlanes      = 1
  bih.biClrUsed     = 0                                                    'Number of color indexes in the color table actually used. 0 if maximum.
  bih.biSize        = SizeOf(bih)
  GdipGetImagePaletteSize(pImage,@dwSizePalette)                           'Size, in bytes, of color masks, 8 or 12 bytes, and color palette
  pPalette=GlobalAlloc(GMEM_FIXED Or GMEM_ZEROINIT,dwSizePalette+12)       'Add 12 bytes: color masks, normally 3 DWORDs
  GdipGetImagePalette(pImage,pPalette,dwSizePalette)
  '
  Select Case dwPixelFormat                                                'Convert pixel formats, which we don't support
  Case PIXELFORMAT1BPPINDEXED,PIXELFORMAT4BPPINDEXED,PIXELFORMAT8BPPINDEXED
  Case PIXELFORMAT32BPPRGB,PIXELFORMAT32BPPARGB,PIXELFORMAT32BPPPARGB,PIXELFORMAT48BPPRGB,PIXELFORMAT64BPPARGB,PIXELFORMAT64BPPPARGB
    dwPixelFormat=PIXELFORMAT32BPPRGB
  Case Else
    dwPixelFormat=PIXELFORMAT24BPPRGB
  End Select
  '
  GdipBitmapLockBits(pImage,NULL,IMAGELOCKMODEREAD,dwPixelFormat,@BmpData)  'Request a bitmap with color deep as we set before
  '
  dwStrideLength=Abs(BmpData.stride)                      'Stride width: Scan width in pixels (a scan line), rounded up to a four-byte boundary. If positive: Bitmap is top-down. If negative: Bitmap is bottom-up.
  '
  Select Case dwPixelFormat                               'Calculate DIB bits buffer size, padding correct
  Case PIXELFORMAT1BPPINDEXED                             '1 bpp indexed
    dwSizeOfPalette    = 2*4                              '2 RGB quads = 4 bytes (umber of palette entries depends on the values of the biBitCount and biClrUsed)
    bih.biBitCount     = 1
    dwPaddingValue     = (lImageWidth\8)
  Case PIXELFORMAT4BPPINDEXED                             '4 bpp indexed
    dwSizeOfPalette    = 16*4                             '16 RGB quads = 4 bytes (umber of palette entries depends on the values of the biBitCount and biClrUsed)
    bih.biBitCount     = 4
    dwPaddingValue     = (lImageWidth\2)
  Case PIXELFORMAT8BPPINDEXED                             '8 bpp indexed
    dwSizeOfPalette    = 256*4                            '256 RGB quads = 4 bytes (umber of palette entries depends on the values of the biBitCount and biClrUsed)
    bih.biBitCount     = 8
    dwPaddingValue     = lImageWidth
  Case PIXELFORMAT32BPPRGB                                '32 bit: untested!
    dwSizeOfPalette    = 0                                'RGB bitmaps don't have a palette ("packed bitmap")
    bih.biBitCount     = 32
    dwPaddingValue     = lImageWidth*4
  Case Else                                               '24 bpp (PIXELFORMAT24BPPRGB)
    dwSizeOfPalette    = 0                                'RGB bitmaps don't have a palette ("packed bitmap")
    bih.biBitCount     = 24
    dwPaddingValue     = lImageWidth*3
  End Select
  '
  'Now we create an entire RTF file with an embedded decompressed DIB (bitmap) picture in HEX format. For comparison here are two an example files: PNG/JPG and DIB:
  'Watch: RTF with compressed PNG/JPG works fine with Winword97 upwards, but works *not* with RichEdit based on msftedit.dll up to
  'version 5.41.21.2510 (latest 4.1 release)! Replace "pngblip" by "jpgblip" depending on the file type.
  '{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Arial;}}
  '\fs28 Document with mini PNG compressed graphic\par
  '{\pict\wmetafile1\pngblip\picwgoal240\pichgoal240  89504e470d0a1a0a0000000d4948445200000010000000100103000000253d6d2200000006504c5445000000ff00001bff8d22000000134944415408d7636060
  '60a8fb878eb00a82c501c9e5134d22283ce10000000049454e44ae426082 }
  '\par\fs20 This is some text. 
  '}
  '
  'string mpic = @"{\pict\pngblip\picw" + img.Width.ToString() + @"\pich" + img.Height.ToString() + @"\picwgoal" + width.ToString() + @"\pichgoal" + height.ToString() + @"\bin " + str + "}";
  '\emfblip      Source of the picture is an EMF (enhanced metafile).
  '\pngblip      Source of the picture is a PNG.
  '\jpegblip     Source of the picture is a JPEG.
  '\shppict      Specifies a Word 97-2000 picture. This is a destination control word.
  '\nonshppict   Specifies that Word 97-2000 has written a {\pict destination that it will not read on input. This keyword is for compatibility with other readers.
  '\macpict      Source of the picture is QuickDraw.
  '\pmmetafileN  Source of the picture is an OS/2 metafile. The N argument identifies the metafile type. The N values are described in the \pmmetafile table below.
  '\wmetafileN   Source of the picture is a Windows metafile. The N argument identifies the metafile type (the default is 1).
  '\dibitmapN    Source of the picture is a Windows device-independent bitmap. The N argument identifies the bitmap type (must equal 0).The information to be included in RTF from a Windows device-independent bitmap is the concatenation of the BITMAPINFO structure followed by the actual pixel data.    
  '\wbitmapN     Source of the picture is a Windows device-dependent bitmap. The N argument identifies the bitmap type (must equal 0).The information to be included in RTF from a Windows device-dependent bitmap is the result of the GetBitmapBits function.
  '
  'WordPad ignores any image not stored in the proper Windows MetaFile format. Thus, the previous example on this page will not show up at all (Even though it works just fine in OpenOffice and Word itself). The format that WordPad WILL support is:
  '{/pict/wmetafile8/picw[width]/pich[height]/picwgoal[scaledwidth]/pichgoal[scaledheight] [image-as-string-of-byte-hex-values]} (with terms in square brackets replaced with the appropriate data).
  '
  'Now the downwards compatible uncompressed DIB (the DIB hex bytes are shorted and incomplete in this example):
  '{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Arial;}} {\pict\wmetafile8\picw423\pich423\picwgoal240\pichgoal240 0100090000038a000000000066}\par}
  '
  'As a last example here is an RTF file with text only:
  '{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Arial;}}Hallo1\par Hallo2}
  '
  'We use this way:
  '
  'sRtfStart="{\rtf1{\pict\dibitmap "+Chr(13,10)      'Works fine with msftedit.dll, Wordpad and LibreOffice sWrite, but not with Winword 2000
  'sRtfStart="{\rtf1\qc\vertalc{\pict\dibitmap\wbmwidthbytes64\picw500\pich500 "+Chr(13,10)   'Works more compatible. This example width+height 500 pixels, raster line 64 bytes
  'sRtfStart="{\rtf1\qc\vertalc{\pict\dibitmap\wbmwidthbytes"+Str(dwStrideLength)+"\picw"+Str(bih.biWidth)+"\pich"+Str(bih.biHeight)+" "+Chr(13,10)   'Works more compatible
  sRtfStart="{\rtf1\qc\vertalc{\pict\dibitmap\wbmwidthbytes"+Str(dwStrideLength)+"\picw"+Str(bih.biWidth)+"\pich"+Str(bih.biHeight)+" "+Chr(13,10)   'Works more compatible
           '\wbmwidthbytesN   Specifies the number of bytes in each raster line. This value must be an even number because the Windows graphics device interface (GDI)
           '                  assumes that the bit values of a bitmap form an array of integer (two-byte) values. In other words, \wbmwidthbytes times 8 must be the
           '                  next multiple of 16 greater than or equal to the \picw (bitmap width in pixels) value.
           '\picwN            xExt field if the picture is a Windows metafile; picture width in pixels if the picture is a bitmap. The N argument is a long integer.
           '\pichN            yExt field if the picture is a Windows metafile; picture height in pixels if the picture is a bitmap. The N argument is a long integer.
           '\qc               Centered.
           '\vertalc          Centered vertically.
  '
  sRtfEnd=Chr(13,10)+"}}"
  '
  'Now calculate the needed memory for the entire DIB (bitmap) and allocate the memory:
  dwLenBitmapBits=(lImageHeight*dwStrideLength)+dwPaddingValue
  bih.biSizeImage=SizeOf(BITMAPINFOHEADER)+dwSizeOfPalette+dwLenBitmapBits
  dwSizeOfRtfFile=Len(sRtfStart)+bih.biSizeImage*2+Len(sRtfEnd)                               'No "BM"???,    bih.biSizeImage*2  ---> because each HEX byte needs 2 real bytes
  '
  pbDIB=GlobalAlloc(GMEM_FIXED,dwSizeOfRtfFile)                                               'This memory block receives the enire RTF file containing the DIB
  pszCurOut=pbDIB
  '
  CopyMemory(pszCurOut,StrPtr(sRtfStart),Len(sRtfStart))                                      'Copy the starting RTF file part into memory block
  pszCurOut+=Len(sRtfStart)                                                                   'Update pointer
  '
  For pCurInByte=Cast(UBYTE Ptr,@bih) To Cast(UBYTE Ptr,@bih)+SizeOf(bih)-1                   'First element of the DIB is the BITMAPINFO structure. Copy in HEX format.
    *pszCurOut=Hex(*pCurInByte,2)
    pszCurOut+=2                                                                              'Increment by two bytes (each HEX byte needs space of 2 characters)
  Next pCurInByte
  '
  If dwSizeOfPalette>0 Then                                                                   'Copy palette, if used. RGB bitmaps don't have a palette ("packed bitmap")
    For pCurInByte=Cast(UBYTE Ptr,pPalette)+8 To Cast(UBYTE Ptr,pPalette)+dwSizeOfPalette+8-1 'Without color masks bytes. Question: Normally 12 byts, why 8 bytes in this case?
      *pszCurOut=Hex(*pCurInByte,2)
      pszCurOut+=2                                                                            'Increment by two bytes (each HEX byte needs space of 2 characters)
    Next pCurInByte
  End If
  '
  If BmpData.stride<0 Then                                                                    'Bitmap pixel data are stored bottom-up: We can copy data directly
    For pCurInByte=Cast(UBYTE Ptr,BmpData.Scan0) To Cast(UBYTE Ptr,BmpData.Scan0)+dwLenBitmapBits-1
      *pszCurOut=Hex(*pCurInByte,2)
      pszCurOut+=2                                                                            'Increment by two bytes (each HEX byte needs space of 2 characters)
    Next pCurInByte
  Else                                                                                        'Bitmap pixel data are stored top-down: We must invert the scan lines
    pCurInByte=Cast(UBYTE Ptr,BmpData.Scan0)+dwLenBitmapBits-dwStrideLength-dwPaddingValue    'Now copy the DIB bits into HEX format scanline by scanline
    Do
      dwEndOfScanline=Cast(Dword,pCurInByte)+dwStrideLength
      Do                                                                                      'Next scan line                                                
        *pszCurOut=Hex(*pCurInByte,2)
        pszCurOut+=2                                                                          'Increment by two bytes (each HEX byte needs space of 2 characters)
        pCurInByte+=1
      Loop Until pCurInByte=dwEndOfScanline
      pCurInByte=pCurInByte-(dwStrideLength*2)
    Loop Until pCurInByte<BmpData.Scan0
  End If
  ' 
  GdipBitmapUnlockBits(pImage,@BmpData)                                      'No longer access to the GDI+ image data needed
  If pImage Then GdipDisposeImage(pImage)                                    'Cleanup
  GdiPlusShutDown(Cast(ULONG_PTR,pGdiPlusToken))                             'More cleanup
  '
  CopyMemory(pszCurOut,StrPtr(sRtfEnd),Len(sRtfEnd))                         'Copy the closing RTF file part into memory block
  pszCurOut+=Len(sRtfEnd)                                                    'Update pointer
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  '  'For debugging only:
  '  Static hFile As HFILE
  '  hFile=_lcreat("#testfile.rtf",0)
  '  _hwrite(hFile,pbDIB,dwSizeOfRtfFile)
  '  _lclose(hFile)
  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
  GlobalUnlock(hInBuffer)                                                    'Unlock the input picture memory (PNG/JPEG file or what else in binary format)
  GlobalFree(hInBuffer)                                                      'Free the input picture memory
  hInBuffer=NULL                                                             'Invalidate handle (not really necesssary, but...)
  GlobalFree(pPalette)
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  Done_DecompressPicture:
  Function=pbDIB
End Function
'*************************************************************************************************************************************
Function MetaInfoToRtf(ByVal pInBuffer As Any Ptr) As Any Ptr
  'Input must be UTF-8 plain text containing META info in simplified Dublin Core, output is same, but in RTF format containing a table.
  'Input and output are locked global memory objects. The input object is freed after use. The output object must be freed by calling function.
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'Simplified Dublin code input example (each dataset is one line, separator is ":" character:
  '
  'title: Der siebte Horkrux
  'subtitle: Alternativer Band 7 der Harry-Potter-Reihe
  'relation: The seventh Horcrux 
  'relation: Harry Potter
  '...
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  'Example of a minimal RTF file with table:
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}}\uc1\f0\fs20\trowd\cellx2552\cellx9476\pard\intbl\nowidctlparc            [137 Bytes]
  '\b TEST 1a\b0\cell TEST 1b\cell\row                                                                                                                 [22 bytes without content]
  '\b TEST 2a\b0\cell TEST 2b\cell\row 
  '}
  'Any character >127 we encode as a Unicode character (example: \u9786?)
  '
  'For the output memory block we calculate:
  '140 bytes RTF header (real: 137)
  ' 22 bytes RTF table code each line (without content)
  '  7 bytes RTF code for Unicode each character as a (theoretically) possible maximum
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static pOutBuffer As Wstring Ptr,dwSizeInBuffer As Dword,dwSizeOutBuffer As Dword
  Static pCurInChar As UShort Ptr,pCurOutChar As UShort Ptr,pLastColonChar As UShort Ptr,dwNumCharactersInLine As Dword,dwNumValidColons As Dword
  Static dwNumLines As Dword,dwNumCharacters As Dword,sBuffer As String,pbIn As UByte Ptr,pbOut As UByte Ptr,iPosInRightPartOfLine As Integer
  '
  pOutBuffer=0                                        'Default function return value: No output buffer 
  If pInBuffer=0 Then Goto Done_MetaInfoToRtf
  dwSizeInBuffer=GlobalSize(pInBuffer)
  If dwSizeInBuffer=0 Then Goto Done_MetaInfoToRtf
  '
  'First convert 8 bit text into 16 bit wide character text. Normally UTF8 is prescribed, but we handle ANSI text (with it's possible problems) too:
  dwSizeInBuffer=GlobalSize(pInBuffer)
  If *Cast(UByte Ptr,pInBuffer)=&HEF And *Cast(UByte Ptr,pInBuffer+1)=&HBB And *Cast(UByte Ptr,pInBuffer+2)=&HBF Then   'UTF-8 ByteOrderMark "ï»¿"
    dwNumCharacters=MultiByteToWideChar(CP_UTF8,0,pInBuffer+3,dwSizeInBuffer-3,NULL,NULL)                               'Get needed size of output buffer in wide characters
    dwSizeOutBuffer=dwNumCharacters*2
    pOutBuffer=GlobalAlloc(GMEM_FIXED,dwSizeOutBuffer*2)
    dwNumCharacters=MultiByteToWideChar(CP_UTF8,0,pInBuffer+3,dwSizeInBuffer-3,pOutBuffer,dwNumCharacters)               'Convert UTF-8 into 16-bit Unicode
  Else  'No UTF8 BOM
    dwNumCharacters=MultiByteToWideChar(CP_ACP,0,pInBuffer,dwSizeInBuffer,NULL,NULL)                                     'Get needed size of output buffer in bytes
    dwSizeOutBuffer=dwNumCharacters*2
    pOutBuffer=GlobalAlloc(GMEM_FIXED,dwSizeOutBuffer)
    dwNumCharacters=MultiByteToWideChar(CP_ACP,0,pInBuffer,dwSizeInBuffer,pOutBuffer,dwNumCharacters)                    'Convert local codepage text into 16-bit Unicode
  End If
  *Cast(UShort Ptr,pOutBuffer+dwNumCharacters)=&H0000                                                                  'NULL termination
  GlobalFree(pInBuffer)
  ' . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'Now decide, whether the input file is a Dublin Core Table or a plain text description only:
  '
  dwNumLines=1
  dwNumCharactersInLine=0
  dwNumValidColons=0
  For pCurInChar=pOutBuffer To pOutBuffer+dwNumCharacters-1
    dwNumCharactersInLine+=1
    Select Case *pCurInChar
    Case &H000D                                 'Carriage
      dwNumLines+=1                             'Count-up lines
      dwNumCharactersInLine=0
    Case &H003A                                 'Colon (":") character
      If dwNumCharactersInLine<33 Then          'Count-up first colons in a line if in very left part of line. A Dublin Core element may have a length of max. 32 characters.
        dwNumValidColons+=1
        dwNumCharactersInLine=100000            'Prevent counting more than one colon character per line
      End If
    End Select
  Next pCurInChar
  ' . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'If plain text only, then generate a simple text RTF file only. Otherwise skip and make a nice RTF table:
  '
  If dwNumLines-3>dwNumValidColons Then                                     'This seems to be plain text only, not Dublin Core information.
    pInBuffer=pOutBuffer                                                    'Last output becomes input
    dwSizeOutBuffer=200+(22*dwNumLines)+(7*dwSizeInBuffer)                  'Result might be 16KB for a 1100 bytes input Meta text file for example (lot of reserves...)
    pOutBuffer=GlobalAlloc(GMEM_FIXED,dwSizeOutBuffer)      
    pbOut=Cast(UByte Ptr,pOutBuffer)
    '
    sBuffer="{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}}\uc1\f0\fs20 "+Chr(13)
    CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
    pbOut+=Len(sBuffer)
    '
    For pbIn=pInBuffer To pInBuffer+(dwNumCharacters*2) Step 2                 'Beware of Unicode!
      Select Case *Cast(UShort Ptr,pbIn)                                    'Beware of Unicode!
      Case &H000D                                                           'Carriage
        sBuffer="\par"+Chr(13)
        CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
        pbOut+=Len(sBuffer)
      Case Is<&H0020                                                        'Ignore noise
      Case Is<&H0080                                                        'ANSI characters, we not have to convert into Unicode
        *pbOut=*pbIn
        pbOut+=1                                                            'ANSI characters neet 1 byte in RTF file
      Case Else                                                             'Non-ANSI characters, which we have to convert into Unicode
        sBuffer="\u"+Str(*Cast(UShort Ptr,pbIn))+"?"                        'Example: "ä"-->"\u228?"
        CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
        pbOut+=Len(sBuffer)
      End Select  
    Next pbIn
    sBuffer="}"+Chr(0,0)     'Don't forget end of table if file doesn't end with a Carriage
    CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
    pbOut+=Len(sBuffer)
    '
    GlobalFree(pInBuffer)
    '
    Goto Done_MetaInfoToRtf  'Skip part of function, which creates a table
  End If
  ' . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'This part of function we reach only, if metainfo is assumed to be Dublin Core, we can make a table from
  'As a preparation for RTF table ceation now remove white disturbing space and count number of text lines:
  '
  dwNumLines=0
  iPosInRightPartOfLine=0                                                     'Position after very first colon ":" character in line, where the colon has position 1
  pCurOutChar=pOutBuffer
  For pCurInChar=pOutBuffer To pOutBuffer+dwNumCharacters-1
    Select Case *pCurInChar
    Case &H000D                                                               'Carriage
      If *Cast(UShort Ptr,pCurOutChar-1)<>&H000D  Then                        'Ignore empty lines
        *pCurOutChar=*pCurInChar
        pCurOutChar+=1
        dwNumLines+=1                                                         'Count up number of lines
      End If
      iPosInRightPartOfLine=0
    Case &H0020                                                               'Space
      If *Cast(UShort Ptr,pCurInChar-1)<>&H0020 Then                          'If not repeated space
        If iPosInRightPartOfLine>1 Then                                       'After very first colon copy all spaces, but not the very first space after colon
          *pCurOutChar=*pCurInChar
          pCurOutChar+=1
          If iPosInRightPartOfLine>0 Then iPosInRightPartOfLine+=1            'Count up position if just started (after first colon ":" character)
        End If 
      End If
    Case &H003A                                                               'Colon (":") character
      *pCurOutChar=*pCurInChar
      pCurOutChar+=1
      iPosInRightPartOfLine+=1
    Case Is < &H0020                                                          'Ignore noise
    Case Is > &H0020                                                          'Copy normal character
      *pCurOutChar=*pCurInChar
      pCurOutChar+=1
      If iPosInRightPartOfLine>0 Then iPosInRightPartOfLine+=1                'Count up position if just started (after first colon ":" character)
    End Select
  Next pCurInChar
  If *Cast(UShort Ptr,pCurOutChar-1)=&H000D  Then                             'Empty line at very end?
    pCurOutChar-=1
    dwNumLines-=1                                                             'Count up number of lines
  End If
  dwSizeOutBuffer=Cast(Dword,pCurOutChar)-Cast(Dword,pOutBuffer)              'New size of text in bytes
  *pCurOutChar=&H0000                                                         'New NULL termination
  dwNumLines+=1                                                               'Count up number of lines (end of file is end of last line). Used for output buffer calculation.
  ' . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  'Now create an RTF file on the fly, containing a table with the Dublin Code data:
  '
  pInBuffer=pOutBuffer                                                        'Last output is next input ;-)
  dwSizeInBuffer=dwSizeOutBuffer                                              'Size of last output is size of next input
  '
  'Now Find Line changes for table lines and colon characters for teble rows:
  '
  'The output RTF fiile looks like this:
  'Example of a minimal RTF file with table (\cellx:  x is column width in twips. 15 twips is 1 pixel):
  '{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}}\uc1\f0\fs20\trowd\cellx2552\cellx9476\pard\intbl\nowidctlparc            [137 Bytes]
  '\b TEST 1a\b0\cell TEST 1b\cell\row                                                                                                                 [22 bytes without content]
  '\b TEST 2a\b0\cell TEST 2b\cell\row 
  '}
  'Any character >127 we encode as a Unicode character (example: \u9786?)
  '
  'For the output memory block we calculate:
  '200 bytes RTF header (real: 137 plus closing bracket)
  ' 22 bytes RTF table code each line (without content)
  '  7 bytes Unicode RTF code for each character as a (theoretically) possible worst case maximum
  '
  dwSizeOutBuffer=200+(22*dwNumLines)+(7*dwSizeInBuffer)                  'Result might be 16KB for a 1100 bytes input Meta text file for example (lot of reserves...)
  pOutBuffer=GlobalAlloc(GMEM_FIXED,dwSizeOutBuffer)      
  pbOut=Cast(UByte Ptr,pOutBuffer)
  '
  sBuffer="{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fprq2\fcharset0 Arial;}}\uc1\f0\fs20\trowd\cellx1552\cellx9476\pard\intbl\nowidctlparc "+Chr(13)+"\b "
  CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
  pbOut+=Len(sBuffer)
  '
  iPosInRightPartOfLine=0
  For pbIn=pInBuffer To pInBuffer+dwSizeInBuffer-2 Step 2                 'Beware of Unicode!
    Select Case *Cast(UShort Ptr,pbIn)                                    'Beware of Unicode!
    Case &H000D                                                           'Carriage
      'If *Cast(UShort Ptr,pbIn-4)<>&H000D Then
        If iPosInRightPartOfLine=0 Then
          sBuffer="\b0\cell\cell\row"+Chr(13)+"\b "                       'Including (missed) empty first colon to produce no invalid table
        Else
          sBuffer="\cell\row"+Chr(13)+"\b "
          iPosInRightPartOfLine=0
        End If
        CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
        pbOut+=Len(sBuffer)
      'End If
    Case &H003A                                                           'Colon (":") character
      If iPosInRightPartOfLine>0 Then                                     'Not very first colon: This is text content, handle as a normal character
        *pbOut=*pbIn
        pbOut+=1                                                          'ANSI characters neet 1 byte in RTF file
      Else                                                                'First colon: Dublin code field separator
        iPosInRightPartOfLine+=1
        sBuffer="\b0\cell "
        CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
        pbOut+=Len(sBuffer)
      End If
    Case Is<&H0020                                                        'Ignore noise
    Case Is<&H0080                                                        'ANSI characters, we not have to convert into Unicode
      *pbOut=*pbIn
      pbOut+=1                                                            'ANSI characters neet 1 byte in RTF file
    Case Else                                                             'Non-ANSI characters, which we have to convert into Unicode
      sBuffer="\u"+Str(*Cast(UShort Ptr,pbIn))+"?"                        'Example: "ä"-->"\u228?"
      CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
      pbOut+=Len(sBuffer)
    End Select  
  Next pbIn
  'Don't forget end of table if file doesn't end with a Carriage:
  If iPosInRightPartOfLine=0 Then
    sBuffer="\b0\cell\cell\row}"+Chr(0)                                    'Including (missed) empty first colon to produce no invalid table
  Else
    sBuffer="\cell\row}"+Chr(0)
    iPosInRightPartOfLine=0
  End If
  CopyMemory(pbOut,StrPtr(sBuffer),Len(sBuffer))
  pbOut+=Len(sBuffer)
  '
  GlobalFree(pInBuffer)
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  Done_MetaInfoToRtf:
'  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'  'For debugging only:
'  Static hFile As HFILE
'  hFile=_lcreat("#testfile.rtf",0)
'  _hwrite(hFile,Cast(LPCSTR,pOutBuffer),Cast(Dword,pbOut)-Cast(Dword,pOutBuffer))
'  _lclose(hFile)
'  '/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

  Function=pOutBuffer
End Function
'***************************************************************************************************************************************
Sub ShowSystemErrorMessage(ByVal hWndParent As HWND,ByVal sErrorText As String)
  Static swBuffer As Wstring*2048
  'FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM,NULL,GetLastError(),MAKELANGID(LANG_NEUTRAL,SUBLANG_DEFAULT),Cast(LPWSTR,pAny),0,NULL)
  'LocalFree(pAny)
  FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM,NULL,GetLastError(),0,@swBuffer,SizeOf(swBuffer),NULL)
  
  Select Case UserSettings.wLanguageID                                       'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
  Case &H07 : MessageBox(hWndParent,sErrorText+Chr(13,13)+"System-Fehlermeldung:"+Chr(13)+swBuffer,TXT_APP_NAME,MB_ICONSTOP)   '&H07: German
  Case Else : MessageBox(hWndParent,sErrorText+Chr(13,13)+"System error message:"+Chr(13)+swBuffer,TXT_APP_NAME,MB_ICONSTOP)   '&H00: Neutral (english)
  End Select
End Sub
'***************************************************************************************************************************************
Function StreamOutRtfFile(ByVal hFile As Dword,ByVal pbBuffer As Any Ptr,ByVal dwBytesToWrite As Dword,Byref dwBytesDone As Dword) As LRESULT
  If WriteFile(Cast(HANDLE,hFile),pbBuffer,dwBytesToWrite,@dwBytesDone,NULL) Then
    Function=0  'Success
  Else
    Function=1  'End streaming
  End If
End Function
'*************************************************************************************************************************************
Sub RegisterFileExtension()
  'Registers the file extension TXT_RTFBOOK_EXTENSION=".rbk" for the application, if not just registered.
  'Accessed global variables: swAppFile
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static hKey As HKEY,dwLenBuffer As Dword,swzBuffer As Wstring*MAX_PATH
  '
  swzBuffer=Left(swAppFile,InStr(swAppFile,"\"))                                                       'Gets for example "c:\" (we need root dir of drive)
  If GetDriveType(swzBuffer)<>DRIVE_FIXED Then Goto Done_RegisterFileExtension                         'Don't register file extension, if app's drive is not a fixed rive
  '
  'If extension ".rbk" just registered, then don't do it again:
  If RegOpenKey(HKEY_CLASSES_ROOT,TXT_RTFBOOK_EXTENSION,@hKey) = ERROR_SUCCESS Then                    'Is extension ".rbk" just registered ?
    dwLenBuffer=SizeOf(swzBuffer)
    swzBuffer=""
    RegQueryValueEx(hKey,"",0,NULL,Cast(Byte Ptr,@swzBuffer),@dwLenBuffer)                             'For which application?
    RegCloseKey(hKey)
    If swzBuffer=TXT_APP_NAME Then                                                                     'If registered for own applicatrion 
      If RegOpenKey(HKEY_CLASSES_ROOT,TXT_APP_NAME+"\shell\open\command",@hKey) = ERROR_SUCCESS Then   'Get filename of registered application with path
        dwLenBuffer=SizeOf(swzBuffer)
        swzBuffer=""
        RegQueryValueEx(hKey,"",0,NULL,Cast(Byte Ptr,@swzBuffer),@dwLenBuffer)                         'Get into string
        RegCloseKey(hKey)
        If InStr(swzBuffer,swAppFile)=1 Then Goto Done_RegisterFileExtension              'If registered for own application with same path, then skip rest of sub
      End If
    End If
  End If
  '
  RegCreateKey(HKEY_CLASSES_ROOT,TXT_RTFBOOK_EXTENSION,@hKey)          '#define TXT_RTFBOOK_EXTENSION ".rbk"
  RegSetValue(hKey,"",REG_SZ,TXT_APP_NAME,NULL)
  RegCloseKey(hKey)
  '
  RegCreateKey(HKEY_CLASSES_ROOT,TXT_APP_NAME,@hKey)
  RegSetValue(hKey,"",REG_SZ,TXT_APP_NAME,NULL)
  RegCloseKey(hKey)
  '
  RegCreateKey(HKEY_CLASSES_ROOT,TXT_APP_NAME,@hKey)
  RegSetValue(hKey,"shell\open\command",REG_SZ,swAppFile+" %1",MAX_PATH)
  RegCloseKey(hKey)
  '
  RegCreateKey(HKEY_CLASSES_ROOT,TXT_APP_NAME,@hKey)
  RegSetValue(hKey,"DefaultIcon",REG_SZ,swAppFile,MAX_PATH)   'or swAppFile+",1"   (Resource ID of the icon)
  RegCloseKey(hKey)
  '
  '  Force Icon cache refresh (otherwise the associated icon is shown only after reboot):
  '  Hint: Possibly WinXP seems to require the SHCNF_FLUSHNOWAIT flag in SHChangeNotify.
  '  SHChangeNotify(SHCNE_ASSOCCHANGED,SHCNF_IDLIST,ByVal 0,ByVal 0)
  '. . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .
  Done_RegisterFileExtension:
End Sub
'***************************************************************************************************************************************
Sub UnRegisterFileExtension()
  'Unregisters the file extension TXT_RTFBOOK_EXTENSION=".rbk" for the application
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static hKey As HKEY
  '
  If RegOpenKey(HKEY_CLASSES_ROOT,"",@hKey) = ERROR_SUCCESS Then
    RegDeleteKey(hKey,TXT_RTFBOOK_EXTENSION)
    '
    RegDeleteKey(hKey,TXT_APP_NAME+"\shell\open\command")
    RegDeleteKey(hKey,TXT_APP_NAME+"\shell\open")
    RegDeleteKey(hKey,TXT_APP_NAME+"\shell")
    RegDeleteKey(hKey,TXT_APP_NAME+"\DefaultIcon")
    RegDeleteKey(hKey,TXT_APP_NAME)
    '
    RegCloseKey(hKey)
  End If
  '
  'Now force Icon Refresh (otherwise the associated icon is shown only after reboot):
  'WinXP seems to require the SHCNF_FLUSHNOWAIT flag in SHChangeNotify
  'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/functions/shchangenotify.asp
  SHChangeNotify(SHCNE_ASSOCCHANGED,SHCNF_IDLIST,ByVal 0,ByVal 0)
End Sub
'***************************************************************************************************************************************
Function PrintRichText(ByVal ptPrintInfos As PrintInfos Ptr) As LRESULT
  'This thread function prints the contents of an RTF text box given it's handle, the calling program's handle(s), and the page margins.
  '
  'The type tPrintInfos contains the thread function parameters and has the following members:
  'tPrintInfos.hwndMain As HWND                               'Handle of the main application window
  'tPrintInfos.hInst As HINSTANCE                             'Instance handle of the application
  'tPrintInfos.hwndText As HWND                               'Handle of the RichEdit control, which's text is to print
  'tPrintInfos.lLeftMargin As Long                            'Page margin in Millimeters
  'tPrintInfos.lRightMargin As Long                           'Page margin in Millimeters
  'tPrintInfos.lTopMargin As Long                             'Page margin in Millimeters
  'tPrintInfos.lBottomMargin As Long                          'Page margin in Millimeters
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Static tFormatRange As FORMATRANGE,tDocInfo As DOCINFO,tPrintDlg As PRINTDLG,dwTextDone As Long,dwTextLength As Long,iWidthTwips As Long,iHeightTwips As Long
  Static lPageNumber As Long,szwBuffer As Wstring*255,tGetTextLengthEx As GETTEXTLENGTHEX,sBuffer As String,tCharRange As CHARRANGE,nCopies As Long
  Static szwNameOfPrintJob As Wstring*MAX_PATH,dwPrintedPages As Dword
  Static lLeftMargin As Long,lRightMargin As Long,lTopMargin As Long,lBottomMargin As Long
  '
  'Convert margins from Millimeters into Twips (1 mm = 56,6928 twips):
  lLeftMargin   = ptPrintInfos->lLeftMargin   * 57
  lRightMargin  = ptPrintInfos->lRightMargin  * 57
  lTopMargin    = ptPrintInfos->lTopMargin    * 57
  lBottomMargin = ptPrintInfos->lBottomMargin * 57
  '
  SendMessage(ptPrintInfos->hwndText,EM_EXGETSEL,0,Cast(Lparam,@tCharRange))  'Get possible selection into tCharRange.cpMin and tCharRange.cpMax (all: cpMin=0, cpMax=-1)
  '
  'Setup the print common dialog:
  ZeroMemory(@tPrintDlg,SizeOf(PRINTDLG))     'Saves sevarals code lines "=0" or "=NULL" :-)
  tPrintDlg.lStructSize               = SizeOf(PRINTDLG)
  tPrintDlg.hwndOwner                 = hwndMain
  'tPrintDlg.hDevMode                 = NULL
  'tPrintDlg.hDevNames                = NULL
  'tPrintDlg.hDC                      = NULL
  If tCharRange.cpMin<>tCharRange.cpMax Then     'There exists a selection
    tPrintDlg.Flags                   = PD_RETURNDC Or PD_NOPAGENUMS Or PD_SELECTION   'Preset the "Print selection" Radiobutton
  Else
    tPrintDlg.Flags                   = PD_RETURNDC Or PD_NOPAGENUMS
  End If
  'tPrintDlg.nFromPage                = 0
  'tPrintDlg.nToPage                  = 0
  'tPrintDlg.nMinPage                 = 0
  'tPrintDlg.nMaxPage                 = 0
  'tPrintDlg.nCopies                  = 0
  tPrintDlg.hInstance                 = ptPrintInfos->hInst
  'tPrintDlg.lCustData                = 0
  'tPrintDlg.lpfnPrintHook            = NULL
  'tPrintDlg.lpfnSetupHook            = NULL
  'tPrintDlg.lpPrintTemplateName      = NULL
  'tPrintDlg.lpPrintSetupTemplateName = NULL
  'tPrintDlg.hPrintTemplate           = NULL
  'tPrintDlg.hSetupTemplate           = NULL
  '
  If PrintDlg(@tPrintDlg)=0 Then GoTo Done_PrintRichText  'Call the PrintDlg common dialog to get printer name and a hDC for printer
  '
  Static hwndWaitMessageBox As HWND,hwndWaitText As HWND
  hwndWaitMessageBox=CreateDialog(ptPrintInfos->hInst,MAKEINTRESOURCE(IDD_WAITDLG),hwndMain,@WaitDlgProc)   'Open non-modal "Wait" window, caption and Static text control IDC_WAITTEXT receives the text
  hwndWaitText=GetDlgItem(hwndWaitMessageBox,IDC_WAITTEXT)                                                  'This STATIC control receives the actual text
  ' 
  'Get page dimensions in Twips:
  iWidthTwips           = Int((GetDeviceCaps(tPrintDlg.hDC,HORZRES) / GetDeviceCaps(tPrintDlg.hDC,LOGPIXELSX))*1440)
  iHeightTwips          = Int((GetDeviceCaps(tPrintDlg.hDC,VERTRES) / GetDeviceCaps(tPrintDlg.hDC,LOGPIXELSY))*1440)
  '
  If (tPrintDlg.Flags And PD_SELECTION) Then                                                                  'Print selected text only
    dwTextLength=Abs(tCharRange.cpMax-tCharRange.cpMin)
  Else                                                                                                        'Print all text
    tGetTextLengthEx.flags=GTL_NUMCHARS   'Or GTL_PRECISE 'Or GTL_USECRLF
    tGetTextLengthEx.codepage=1200                                                                            '1200: Unicode
    dwTextLength=SendMessage(ptPrintInfos->hwndText,EM_GETTEXTLENGTHEX,Cast(wparam,@tGetTextLengthEx),NULL)   'Total length of text to be formatted and printed
    tCharRange.cpMin         = 0                                                                              'Range of text to format. 
    tCharRange.cpMax         = dwTextLength                                                                   'Range of text to format.
  End If
  '
  GetWindowText(hWndMain,@szwNameOfPrintJob,SizeOf(szwNameOfPrintJob))
  dwPrintedPages=0
  For nCopies=1 To tPrintDlg.nCopies
    lPageNumber   = 0      'Initialize (currently no page is printed)
    dwTextDone    = 0      'Index of last character printed... not yet, anyway
    '
    tDocInfo.cbSize       = SizeOf(DOCINFO)
    tDocInfo.lpszDocName  = VarPtr(szwNameOfPrintJob)
    tDocInfo.lpszOutput   = NULL
    tDocInfo.fwType       = DI_APPBANDING
    StartDoc(tPrintDlg.hDC,@tDocInfo)                              'We actually do not need to do a startdoc unless we are on page one
    '
    SendMessage(ptPrintInfos->hwndText,EM_FORMATRANGE,1,NULL)      'Free possibly former cached information after last use 
    '
    'Set the fix part of the FORMATRANGE type. Note, that the measurements can change when rendering a lot of pages,
    'so that you get after rendering of >100 pages or so unfortunatels will get a more and more larger bottom margin.
    tFormatRange.hdc                = tPrintDlg.hDC                'Device to render to.
    tFormatRange.hdcTarget          = tPrintDlg.hDC                'Target device to format for.
    tFormatRange.chrg.cpMin         = tCharRange.cpMin             'Range of text to format. 
    tFormatRange.chrg.cpMax         = tCharRange.cpMax             'Range of text to format. 
    Do While dwTextDone<dwTextLength
      'It might look crazy to set all the tFormatRange measurement elements again and again. But if we don't do so,
      'the bottom margin of the pages becomes larger and larger until printing hundreds of pages.
      tFormatRange.rc.Top          = lTopMargin                    'Area to render to. Units in twips. 
      tFormatRange.rc.Left         = lLeftMargin                   'Area to render to. Units in twips. 
      tFormatRange.rc.Right        = iWidthTwips-lRightMargin      'Area to render to. Units in twips. 
      tFormatRange.rc.Bottom       = iHeightTwips-lBottomMargin    'Area to render to. Units in twips. 
      tFormatRange.rcPage.Top      = tFormatRange.rc.Top           'Entire area of rendering device. Units in twips.
      tFormatRange.rcPage.Left     = tFormatRange.rc.Left          'Entire area of rendering device. Units in twips.
      tFormatRange.rcPage.Right    = tFormatRange.rc.Right         'Entire area of rendering device. Units in twips.
      tFormatRange.rcPage.Bottom   = tFormatRange.rc.Bottom        'Entire area of rendering device. Units in twips.

      If tPrintDlg.nCopies=1 Then
        Select Case UserSettings.wLanguageID                        'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07 : szwBuffer="Drucke Seite "+Str(lPageNumber)      'German: Default language in resources
        Case Else : szwBuffer="Print page "+Str(lPageNumber)        'Neutral (english)
        End Select
      Else
        Select Case UserSettings.wLanguageID                        'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
        Case &H07 : szwBuffer="Drucke Seite "+Str(lPageNumber)+" von Kopie "+Str(nCopies)     'German: Default language in resources
        Case Else : szwBuffer="Print page "+Str(lPageNumber)+" of copy "+Str(nCopies)         'Neutral (english)
        End Select
      End If
      SetWindowText(hwndWaitText,szwBuffer)
      '
      StartPage(tPrintDlg.hDC)
      dwTextDone=SendMessage(ptPrintInfos->hwndText,EM_FORMATRANGE,1,Cast(lparam,@tFormatRange))           'Now render the current page
      EndPage(tPrintDlg.hDC)
      tFormatRange.chrg.cpmin = dwTextDone
      tFormatRange.chrg.cpMax = -1
      lPageNumber+=1
      Sleep(0)
    Loop
    '
    SendMessage(ptPrintInfos->hwndText,EM_FORMATRANGE,1,NULL)                                        'Important: Free cached information after last use 
    '
    EndDoc(tPrintDlg.hDC)  'Finish the printing
    '
    dwPrintedPages=dwPrintedPages+lPageNumber      'Count up all pages (including multiples if more than one copy)
    '
    If (tPrintDlg.nCopies>1) And (nCopies=1) Then  'Request onece only if more copies then one
      ShowWindow(hwndWaitMessageBox,SW_HIDE)
      Select Case UserSettings.wLanguageID         'User language ID (&H00=Neutral (english), &H07=German, &H03=Catalan and so on )
      Case &H07 : szwBuffer="Mehr Kopien?"         'German: Default language in resources
      Case Else : szwBuffer="More copies?"         'Neutral (english)
      End Select
      If MessageBox(hwndMain,szwBuffer,TXT_APP_NAME,MB_ICONQUESTION Or MB_YESNO)=IDNO Then Exit For
      ShowWindow(hwndWaitMessageBox,SW_SHOWNORMAL)
    End If
  Next nCopies
  DeleteDC(tPrintDlg.hDC)
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Done_PrintRichText:
  If IsWindow(hwndWaitMessageBox) Then DestroyWindow(hwndWaitMessageBox)   'Close the "Wait" MessageBox window
  Function=dwPrintedPages 'Set return value = # pages printed
End Function
'***************************************************************************************************************************************
Function WaitDlgProc(ByVal hDlg As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As LRESULT
  'DialogBox function for the Wait dialog window (progress while printing/loading/saving)
  'Hint:
  'The one and only control of this dialog, IDC_WAITTEXT, is a STATIC text control with style SS_SIMPLE|SS_NOPREFIX for fast redraw.
  'see here: https://msdn.microsoft.com/en-us/library/ms997560.aspx
  'The system outputs then using ExtTextOut() instead of DrawText(). But watch: No WM_CTLCOLORSTATIC support, text must fit into control,
  'otherwise it's clipped, one line only, no background redraw of former text.
  '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  Select Case uMsg
  Case WM_INITDIALOG
    SetWindowText(hDlg,TXT_APP_NAME)                 'The default caption text
  Case WM_COMMAND
    Select Case LoWord(wParam)
    Case IDOK,IDCANCEL                               'Button or ESC key
      DestroyWindow(hDlg)                            'Non-modal window! - call EndDialog(hDlg,1) if created as a modal window
    End Select
  End Select
  Function=0
End Function
'***************************************************************************************************************************************
Sub RichCom_SetComInterface(ByVal hWndEdit As HWND)                                                                                             'OLE stuff: initialization
  'Inits the OLE stuff necessary to display a picture in the RichEdit control. The used OLE callback functions are, empty or not, the needed minimum.
  'For information look into platform SDK ---> "Rich Edit Controls" ---> "Rich Edit OLE Interfaces", especially "IRichEditOleCallback"^.
  'We implement an absolute minimum only, as needed to show embedded pictures.
  '
  Static dwFunctionPointers(16) As Dword
  '
  If CLng(pObj_RichCom)=0 Then
    dwFunctionPointers(0) =Cast(Dword,@RichCom_Object_QueryInterface)
    dwFunctionPointers(1) =Cast(Dword,@RichCom_Object_AddRef)
    dwFunctionPointers(2) =Cast(Dword,@RichCom_Object_Release)
    dwFunctionPointers(3) =Cast(Dword,@RichCom_Object_GetNewStorage)
    dwFunctionPointers(4) =Cast(Dword,@RichCom_Object_GetInPlaceContext)
    dwFunctionPointers(5) =Cast(Dword,@RichCom_Object_ShowContainerUI)
    dwFunctionPointers(6) =Cast(Dword,@RichCom_Object_QueryInsertObject)
    dwFunctionPointers(7) =Cast(Dword,@RichCom_Object_DeleteObject)
    dwFunctionPointers(8) =Cast(Dword,@RichCom_Object_QueryAcceptData)
    dwFunctionPointers(9) =Cast(Dword,@RichCom_Object_ContextSensitiveHelp)
    dwFunctionPointers(10)=Cast(Dword,@RichCom_Object_GetClipboardData)
    dwFunctionPointers(11)=Cast(Dword,@RichCom_Object_GetDragDropEffect)
    dwFunctionPointers(12)=Cast(Dword,@RichCom_Object_GetContextMenu)
    pObj_RichComObject.pIntf=@dwFunctionPointers(0)
    pObj_RichComObject.Refcount=1
    pObj_RichCom=Cast(Dword Ptr,@pObj_RichComObject)
  End If
  SendMessage(hWndEdit,EM_SETOLECALLBACK,0,Cast(lParam,pObj_RichCom))
End Sub
'***************************************************************************************************************************************
Function RichCom_Object_QueryInterface(pObject As Dword,REFIID As Dword,ppvObj As Dword) As Dword                                               'OLE stuff
  'Useless function, never called..
  Function=S_OK
End Function
'***************************************************************************************************************************************
Function RichCom_Object_AddRef(ByVal pObject As Dword Ptr) As Dword                                                                             'OLE stuff
  Exit Function
  pObject[1]+=1
  Function=pObject[1]
End Function
'***************************************************************************************************************************************
Function RichCom_Object_Release(ByVal pObject As Dword Ptr) As Dword                                                                            'OLE stuff
  Exit Function
  If pObject[1] > 0 Then
    pObject[1]-=1
    Function=pObject[1]
  Else
    pObject=0
  End If
End Function
'***************************************************************************************************************************************
Function RichCom_Object_GetInPlaceContext(ByVal pObject As Dword Ptr,lplpFrame As Dword,lplpDoc As Dword,lpFrameInfo As Dword) As Dword         'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_ShowContainerUI(ByVal pObject As Dword Ptr,fShow As Long) As Dword                                                      'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_QueryInsertObject(ByVal pObject As Dword Ptr,lpclsid As Dword,ByVal lpstg As Dword Ptr,cp As Long) As Dword             'OLE stuff
  Function=S_OK
End Function
'***************************************************************************************************************************************
Function RichCom_Object_DeleteObject(ByVal pObject As Dword Ptr,lpoleobj As Dword) As Dword                                                     'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_QueryAcceptData(ByVal pObject As Dword Ptr,lpdataobj As Dword,lpcfFormat As Dword,reco As Dword,fReally As Long,hMetaPict As Dword) As Dword  'OLE stuff
  Function=E_NOTIMPL
'  Function=S_OK
End Function
'***************************************************************************************************************************************
Function RichCom_Object_ContextSensitiveHelp(ByVal pObject As Dword Ptr,fEnterMode As Long) As Dword                                            'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_GetClipboardData(ByVal pObject As Dword Ptr,lpchrg As Dword,reco As Dword,lplpdataobj As Dword) As Dword                'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_GetDragDropEffect(ByVal pObject As Dword Ptr,fDrag As Long,grfKeyState As Dword,pdwEffect As Dword) As Dword            'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_GetContextMenu(ByVal pObject As Dword Ptr,seltype As Word,lpoleobj As Dword,lpchrg As Dword,lphmenu As Dword) As Dword  'OLE stuff
  Function=E_NOTIMPL
End Function
'***************************************************************************************************************************************
Function RichCom_Object_GetNewStorage(ByVal pObject As Dword Ptr,lplpstg As Dword) As Dword                                                     'OLE stuff
  Static dwRetVal As Dword,pLockBytes As LPLOCKBYTES
  dwRetVal=CreateILockBytesOnHGlobal(NULL,True,@pLockBytes)
  If dwRetVal Then Goto Done_RichCom_Object_GetNewStorage
  dwRetVal=StgCreateDocfileOnILockBytes(pLockBytes,STGM_SHARE_EXCLUSIVE Or STGM_READWRITE Or STGM_CREATE,NULL,Cast(IStorage Ptr Ptr,lplpstg))
  Done_RichCom_Object_GetNewStorage:
  Function=dwRetVal
End Function
'***************************************************************************************************************************************
'Sub TestTime(ByVal iMode As Integer) 'Time measurement for debugging. 0 starts, 1 ends and shows result
'  Static liTime1 As LongInt,liTime2 As LongInt
'  '
'  If iMode=False Then
'    QueryPerformanceCounter(Cast(LARGE_INTEGER ptr,@liTime1))
'  Else  
'    QueryPerformanceCounter(Cast(LARGE_INTEGER Ptr,@liTime2))
'    liTime2=(liTime2-liTime1)*1000000
'    QueryPerformanceFrequency(Cast(LARGE_INTEGER Ptr,@liTime1))
'    liTime2=liTime2\liTime1
'    MessageBox(0,Str(liTime2\1000000)+" Seconds"+Chr(13)+Str(liTime2\1000)+" Milliseconds"+Chr(13)+Str(liTime2)+" Microseconds","Debug: Time elapsed:",0)
'  End If
'End Sub
