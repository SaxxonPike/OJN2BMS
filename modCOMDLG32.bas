Attribute VB_Name = "modCOMDLG32"
Option Explicit


'=============================================================================================================
'
' modCOMDLG32 Module
' ------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : October 17, 2002
' Created On  : July 01, 2000
'
' VB Versions : 5.0 / 6.0
'
' Requires    : COMDLG32.DLL (Microsoft Common Dialog Library)
'               OLEPRO32.DLL (OLE Automation)
'
' Description : This module was created to give you access to the standard Windows dialogs that are most
'               commonly accessed via the COMDLG32.OCX ActiveX Control.  This module allows you all the
'               functionality that the COMDLG32.OCX control does... plus MORE!  And complete documentation
'               to every step of the way is provided to make sure you understand what's going on and what
'               options are available to you.  And because you have access to the code driving the display
'               process of the dialogs, you have more control over what's going on.
'
'               This module gives you access to the following equivelant methods found in the COMDLG32.OCX
'               control:
'
'                 - ShowColor
'                 - ShowFont         * Memory Leak?
'                 - ShowHelp
'                 - ShowOpen
'                 - ShowSave
'                 - ShowPrinter      * Memory Leak?
'
'               This module also gives you access to the following which are NOT available via the
'               COMDLG32.OCX control:
'
'                 - ShowPageSetup    * Memory Leak?
'                 - ShowAbout
'                 - ShowFindComputer
'                 - ShowFindFile
'                 - ShowFolder
'                 - ShowFormat
'                 - ShowIcon
'                 - ShowProperties
'                 - ShowReboot
'                 - ShowRun
'                 - ShowShutDown
'
' NOTE        : The COMDLG32.DLL also contains ShowFind & ShowReplace dialogs, but they have proven to be
'               IMPOSSIBLE to use from within VB 5/6.  These functions can be used fine from within C++, etc.
'               but not VB for some reason.  It crashes the IDE every time.
'
' * WARNING   : In testing of these functions, it was found that the following functions OCCASIONALLY had
'               memory leaks : ShowFont, ShowPrinter, & ShowPageSetup.  This is probibly due to the fact that
'               they all use either the LOGFONT structure, or the DEVMODE / DEVNAMES structure(s).  Even though
'               these are deleted from within the functions as they should be, some memory is retained.  This
'               could possibly cause problems if the user calls one of these functions repeatedly.
'
'               To verify these memory leaks, run the Windows Resource Monitor, then call one or more of these
'               3 functions that sometimes have memory leaks repeatedly and with different flag settings.  If
'               you watch the available System, User, & GDI resources... occasionally they will drop by about
'               2% each time these function(s) are called.  Sometimes this memory is released when the program
'               is shut down, other times it is not.  Sometimes this memory is retained just the first time
'               the function is called, other times it retains the memory every time the function is called.
'               If you have any insite as to the cause of / fix for this, please Email it to me at:
'               VBZ@thevbzone.com
'
' See Also    : Undocumented Windows APIs
'               http://www.geocities.com/SiliconValley/4942/index.html
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


' Types / Enumerations
Public Type POINTAPI
  X                 As Long    ' X Coordinate for the point
  Y                 As Long    ' Y Coordinate for the point
End Type

Public Type RECT
  Left              As Long    ' Left of the rectangle
  Top               As Long    ' Top the rectangle
  Right             As Long    ' Right of the rectangle (Left + Width)
  Bottom            As Long    ' Bottom of the rectangle (Top + Height)
End Type

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 '  Maintenance string for PSS usage
End Type

Private Enum OSTypes
  OS_Unknown = 0     ' "Unknown"
  OS_Win32 = 32      ' "Win 32"
  OS_Win95 = 95      ' "Windows 95"
  OS_Win98 = 98      ' "Windows 98"
  OS_WinNT_351 = 351 ' "Windows NT 3.51"
  OS_WinNT_40 = 40   ' "Windows NT 4.0"
  OS_Win2000 = 2000  ' "Windows 2000"
End Enum

' Windows API FONT information type
Public Type LOGFONT
  lfHeight          As Long    ' Specifies the height, in logical units, of the font’s character cell or character
  lfWidth           As Long    ' Specifies the average width, in logical units, of characters in the font
  lfEscapement      As Long    ' Specifies the angle, in tenths of degrees, between the escapement vector and the x-axis of the device. The escapement vector is parallel to the base line of a row of text
  lfOrientation     As Long    ' Specifies the angle, in tenths of degrees, between each character’s base line and the x-axis of the device
  lfWeight          As Long    ' Specifies the weight of the font in the range 0 through 1000 (See Constants)
  lfItalic          As Byte    ' Specifies an italic font if set to TRUE (1)
  lfUnderline       As Byte    ' Specifies an underlined font if set to TRUE (1)
  lfStrikeOut       As Byte    ' Specifies a strikeout font if set to TRUE (1)
  lfCharSet         As Byte    ' Specifies the character set (See Constants)
  lfOutPrecision    As Byte    ' Specifies the output precision. The output precision defines how closely the output must match the requested font’s height, width, character orientation, escapement, pitch, and font type (See Constants)
  lfClipPrecision   As Byte    ' Specifies the clipping precision. The clipping precision defines how to clip characters that are partially outside the clipping region
  lfQuality         As Byte    ' Specifies the output quality. The output quality defines how carefully the graphics device interface (GDI) must attempt to match the logical-font attributes to those of an actual physical font (See Constants)
  lfPitchAndFamily  As Byte    ' Specifies the pitch and family of the font. The two low-order bits specify the pitch of the font (See Constants) - Bits 4 through 7 of the member specify the font family (See Constants)
  lfFaceName        As String * 31 ' A null-terminated string that specifies the typeface name of the font. The length of this string must not exceed 32 characters, including the null terminator. The EnumFontFamilies function can be used to enumerate the typeface names of all currently available fonts. If lfFaceName is an empty string, GDI uses the first font that matches the other specified attributes.
End Type

' Window API PRINTER information type
Private Type DEVMODE
  ' Standard
  dmDeviceName      As String * 32 ' Specifies the the "friendly" name of the printer (ie - "PCL/HP LaserJet").  This string is unique among device drivers. Note that this name may be truncated to fit in the dmDeviceName array
  dmSpecVersion     As Integer ' Specifies the version number of the initialization data specification on which the structure is based
  dmDriverVersion   As Integer ' Specifies the printer driver version number assigned by the printer driver developer
  dmSize            As Integer ' Specifies the size, in bytes, of the DEVMODE structure, not including any private driver-specific data that might follow the structure’s public members. You can use this member to determine the number of bytes of public data regardless of the version of the DEVMODE structure being used
  dmDriverExtra     As Integer ' Contains the number of bytes of private driver-data that follow this structure. If a device driver does not use device-specific information, set this member to zero
  dmFields          As Long    ' Sets which members of the type are active.  A set of bit flags that specify whether certain members of the DEVMODE structure have been initialized. If a field is initialized, its corresponding bit flag is set, otherwise the bit flag is clear.  A printer driver supports only those DEVMODE structure members that are appropriate for the printer technology (See Constants)
  dmOrientation     As Integer ' Selects the orientation of the paper - This can be either DMORIENT_PORTRAIT (1) or DMORIENT_LANDSCAPE (2)
  dmPaperSize       As Integer ' Selects the size of the paper to print on. This member can be set to zero if the length and width of the paper are both set by the dmPaperLength and dmPaperWidth members (See Constants)
  dmPaperLength     As Integer ' Selects the size of the paper to print on. This member can be set to zero if the length and width of the paper are both set by the dmPaperLength and dmPaperWidth members (See Constants)
  dmPaperWidth      As Integer ' Overrides the width of the paper specified by the dmPaperSize member
  dmScale           As Integer ' Specifies the factor by which the printed output is to be scaled. The apparent page size is scaled from the physical page size by a factor of dmScale/100. For example, a letter-sized page with a dmScale value of 50 would contain as much data as a page of 17- by 22-inches because the output text and graphics would be half their original height and width
  dmCopies          As Integer ' Selects the number of copies printed if the device supports multiple-page copies
  dmDefaultSource   As Integer ' Reserved - must be zero
  dmPrintQuality    As Integer ' Specifies the printer resolution (See Constants). If a positive value is given, it specifies the number of dots per inch (DPI) and is therefore device dependent
  dmColor           As Integer ' Switches between color and monochrome on color printers (See Constants)
  dmDuplex          As Integer ' Selects duplex or double-sided printing for printers capable of duplex printing (See Constants)
  dmYResolution     As Integer ' Specifies the Y-resolution, in dots per inch, of the printer. If the printer initializes this member, the dmPrintQuality member specifies the X-resolution, in dots per inch, of the printer
  dmTTOption        As Integer ' Specifies how TrueType® fonts should be printed (See Constants)
 'dmUnusedPadding   As Long    ' ** Used to align the structure to a DWORD boundary. This should not be used or referenced. Its name and usage is reserved, and can change in future releases
  dmCollate         As Integer ' [ TRUE (1) / FALSE (0 ] Specifies whether collation should be used when printing multiple copies. (See Constants) (This member is ignored unless the printer driver indicates support for collation by setting the dmFields member to DM_COLLATE)
  dmFormName        As String * 32 ' Windows NT: Specifies the name of the form to use; for example, "Letter" or "Legal". A complete set of names can be retrieved by using the EnumForms function.  Windows 95: Printer drivers do not use this member
  
  ' Not used by Printer Drivers
  dmUnusedPadding   As Integer ' Specifies the number of pixels per logical inch. Printer drivers do not use this member
  dmBitsPerPel      As Integer ' Specifies the color resolution, in bits per pixel, of the display device (for example: 4 bits for 16 colors, 8 bits for 256 colors, or 16 bits for 65536 colors). Display drivers use this member, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member
  dmPelsWidth       As Long    ' Specifies the width, in pixels, of the visible device surface. Display drivers use this member, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member.
  dmPelsHeight      As Long    ' Specifies the height, in pixels, of the visible device surface. Display drivers use this member, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member.
  dmDisplayFlags    As Long    ' Specifies the device’s display mode (See Constants)
  dmDisplayFrequency As Long   ' Specifies the frequency, in hertz (cycles per second), of the display device in a particular mode. This value is also known as the display device’s vertical refresh rate. Display drivers use this member. It is used, for example, in the ChangeDisplaySettings function. Printer drivers do not use this member
  
  ' Windows 95 Only
  dmICMMethod       As Long    ' Specifies how ICM is handled. For a non-ICM application, this member determines if ICM is enabled or disabled. For ICM applications, Windows examines this member to determine how to handle ICM support. This member can be one of the constant values, or a driver-defined value greater than the value of DMICMMETHOD_USER (See Constants)
  dmICMIntent       As Long    ' Specifies which of the three possible color matching methods, or intents, should be used by default. This member is primarily for non-ICM applications. ICM applications can establish intents by using the ICM functions. This member can be one of the constant values, or a driver defined value greater than the value of DMICM_USER (See Constants)
  dmMediaType       As Long    ' Specifies the type of media being printed on. The member can be one of the constant values, or a driver-defined value greater than the value of DMMEDIA_USER (See Constants)
  dmDitherType      As Long    ' Specifies how dithering is to be done. The member can be one of the constant values, or a driver-defined value greater than the value of DMDITHER_USER (See Constants)
  dmReserved1       As Long    ' Windows 95: Not used - must be zero.  Windows NT: This member is not supported on Windows NT.
  dmReserved2       As Long    ' Windows 95: Not used - must be zero.  Windows NT: This member is not supported on Windows NT.
End Type

' Window API PRINTER information type
Private Type DEVNAMES
  wDriverOffset     As Integer ' (Input/Output) Specifies the offset in characters from the beginning of this structure to a null-terminated string that contains the filename (without the extension) of the device driver. On input, this string is used to determine the printer to display initially in the dialog box.
  wDeviceOffset     As Integer ' (Input/Output) Specifies the offset in characters from the beginning of this structure to the null-terminated string (maximum of 32 bytes including the null) that contains the name of the device. This string must be identical to the dmDeviceName member of the DEVMODE structure.
  wOutputOffset     As Integer ' (Input/Output) Specifies the offset in characters from the beginning of this structure to the null-terminated string that contains the device name for the physical output medium (output port).
  wDefault          As Integer ' Specifies whether the strings contained in the DEVNAMES structure identify the default printer. This string is used to verify that the default printer has not changed since the last print operation. If any of the strings do not match, a warning message is displayed informing the user that the document may need to be reformatted.
                               ' On output, the wDefault member is changed only if the Print Setup dialog box was displayed and the user chose the OK button. The DN_DEFAULTPRN flag is used if the default printer was selected. If a specific printer is selected, the flag is not used. All other flags in this member are reserved for internal use by the Print Dialog box procedure.
  Extra             As String * 100
End Type

' Windows API FIND/REPLACE information type
Private Type MSG
  hWnd              As Long    ' Identifies the window whose window procedure receives the message.
  Message           As Long    ' Specifies the message number.
  wParam            As Long    ' Specifies additional information about the message. The exact meaning depends on the value of the message member.
  lParam            As Long    ' Specifies additional information about the message. The exact meaning depends on the value of the message member.
  Time              As Long    ' Specifies the time at which the message was posted.
  ptX               As Long    ' Specifies the X cursor position (in screen coordinates) when the message was posted.
  ptY               As Long    ' Specifies the Y cursor position (in screen coordinates) when the message was posted.
End Type


'----------------------------------------------------------------------------------
'                      Common Dialog Type Declarations
'----------------------------------------------------------------------------------

' Common Dialog OPEN/SAVE information type
Public Type OPENFILENAME
  lStructSize       As Long    ' Size of this type / structure
  hwndOwner         As Long    ' Handle to the owner of the dialog
  hInstance         As Long    ' Instance handle of .EXE that contains custom dialog template
  lpstrFilter       As String  ' File type filter
  lpstrCustomFilter As String  ' Sets / returns the custom file type
  nMaxCustFilter    As Long    ' Length of the lpstrCustomFilter buffer
  nFilterIndex      As Long    ' Sets / returns the index of the filter in the dialog
  lpstrFile         As String  ' Sets / returns the full path to the file
  nMaxFile          As Long    ' Length of the lpstrFile buffer
  lpstrFileTitle    As String  ' Returns the name of the file selected (w/o the path)
  nMaxFileTitle     As Long    ' Length of the lpstrFileTitle buffer
  lpstrInitialDir   As String  ' Sets the initial browsing directory for the dialog
  lpstrTitle        As String  ' Sets the title of the dialog to be displayed
  Flags             As Long    ' Sets / returns the flags used with the dialog
  nFileOffset       As Integer ' Number of characters from the beginning of the full path to the first letter of the file name
  nFileExtension    As Integer ' Number of characters from the beginning of the full path to the file extension
  lpstrDefExt       As String  ' Sets the default extention of the file
  lCustData         As Long    ' Data passed to hook function
  lpfnHook          As Long    ' Pointer to hook function
  lpTemplateName    As String  ' Custom template name
End Type

' Common Dialog COLOR information type
Private Type CHOOSECOLOR
  lStructSize       As Long    ' Size of this type / structure
  hwndOwner         As Long    ' Handle to the owner of the dialog
  hInstance         As Long    ' Instance handle of .EXE that contains custom dialog template
  rgbResult         As Long    ' Sets / returns the selected color
  lpCustColors      As String  ' Sets / returns custom color to use
  Flags             As Long    ' Sets / returns the flags used with the dialog
  lCustData         As Long    ' Data passed to hook function
  lpfnHook          As Long    ' Pointer to hook function
  lpTemplateName    As String  ' Custom template name
End Type

' Common Dialog FONT information type
Public Type CHOOSEFONT
  lStructSize       As Long    ' Size of this type / structure
  hwndOwner         As Long    ' Handle to the owner of the dialog
  hDC               As Long    ' Printer DC/IC or NULL
  lpLogFont         As Long    ' Pointer to a LOGFONT variable that contains information on the font
  iPointSize        As Long    ' 10 * size in points of selected font
  Flags             As Long    ' Sets / returns the flags for the dialog
  RGBColors         As Long    ' Returned text color
  lCustData         As Long    ' Data passed to hook function
  lpfnHook          As Long    ' Pointer to hook function
  lpTemplateName    As String  ' Custom template name
  hInstance         As Long    ' Instance handle of .EXE that contains custom dialog template
  lpszStyle         As String  ' Return the style field here must be LF_FACESIZE or bigger
  nFontType         As Integer ' Same value reported to the EnumFonts call back with the extra FONTTYPE_ bits added
  MISSING_ALIGNMENT As Integer ' ** Used to align the structure to a WORD boundary. This should not be used or referenced.
  nSizeMin          As Long    ' Minimum pt size allowed
  nSizeMax          As Long    ' Maximum pt size allowed if CF_LIMITSIZE is used
End Type

' Common Dialog PRINT information type
Public Type PRINTDLG
  lStructSize       As Long    ' Size of this type / structure
  hwndOwner         As Long    ' Handle to the owner of the dialog
  hDevMode          As Long    ' Handle to a memory buffer containing DEVMODE information - Can be NULL
  hDevNames         As Long    ' Handle to a memory buffer containing DEVNAMES information - Can be NULL
  hDC               As Long    ' Printer DC/IC or NULL
  Flags             As Long    ' Sets / returns the flags for the dialog
  nFromPage         As Integer ' Sets / returns the value for the starting page edit control
  nToPage           As Integer ' Sets / returns the value for the ending page edit control
  nMinPage          As Integer ' Sets / returns the minimum value for the range of pages specified in the From and To page edit controls
  nMaxPage          As Integer ' Sets / returns the maximum value for the range of pages specified in the From and To page edit controls
  nCopies           As Integer ' Sets / returns the number of copies for the Copies edit control
  hInstance         As Long    ' Instance handle of .EXE that contains custom dialog template
  lCustData         As Long    ' Data passed to hook function
  lpfnPrintHook     As Long    ' Pointer to printer hook function or NULL
  lpfnSetupHook     As Long    ' Pointer to setup hook function or NULL
  lpPrintTemplateName As String ' Custom print template name
  lpSetupTemplateName As String ' Custom setup template name
  hPrintTemplate    As Long    ' Handle of a memory object containing a dialog box template (If PD_ENABLEPRINTTEMPLATEHANDLE flag is set)
  hSetupTemplate    As Long    ' Handle of a memory object containing a dialog box template (If PD_ENABLESETUPTEMPLATEHANDLE flag is set)
End Type

' Common Dialog PAGE SETUP information type
Public Type PAGESETUPDLG
  lStructSize       As Long    ' Size of this type / structure
  hwndOwner         As Long    ' Handle to the owner of the dialog
  hDevMode          As Long    ' Handle to a global memory object that contains a DEVMODE structure
  hDevNames         As Long    ' Handle to a global memory object that contains a DEVNAMES structure
  Flags             As Long    ' Sets / returns the flags for the dialog
  ptPaperSize       As POINTAPI ' Specifies the dimensions of the paper selected by the user
  rtMinMargin       As RECT    ' Specifies the minimum allowable widths for the left, top, right, and bottom margins
  rtMargin          As RECT    ' Specifies the widths of the left, top, right, and bottom margins
  hInstance         As Long    ' Instance handle of .EXE that contains custom dialog template
  lCustData         As Long    ' Data passed to hook function
  lpfnPageSetupHook As Long    ' Pointer to page setup hook function or NULL
  lpfnPagePaintHook As Long    ' Pointer to page paint hook function or NULL
  lpPageSetupTemplateName As String ' Custom setup template name
  hPageSetupTemplate As Long   ' Handle of a memory object containing a dialog box template
End Type

' Common Dialog FIND / REPLACE information type
Private Type FINDREPLACE
  lStructSize       As Long    ' Specifies the length, in bytes, of the structure.
  hwndOwner         As Long    ' Identifies the window that owns the dialog box. Can not be NULL.
  hInstance         As Long    ' If the FR_ENABLETEMPLATEHANDLE flag is set in the Flags member, hInstance is the handle of a memory object containing a dialog box template. If the FR_ENABLETEMPLATE flag is set, hInstance identifies a module that contains a dialog box template named by the lpTemplateName member. If neither flag is set, this member is ignored.
  Flags             As Long    ' Sets / returns the flags for the dialog
  lpstrFindWhat     As Long    ' Pointer to a buffer that a FINDMSGSTRING message uses to pass the null terminated search string that the user typed in the "Find What:" edit control.
  lpstrReplaceWith  As Long    ' Pointer to a buffer that a FINDMSGSTRING message uses to pass the null terminated replacement string that the user typed in the "Replace With:" edit control.
  wFindWhatLen      As Integer ' Specifies the length, in bytes, of the buffer pointed to by the lpstrFindWhat member.
  wReplaceWithLen   As Integer ' Specifies the length, in bytes, of the buffer pointed to by the lpstrReplaceWith member.
  lCustData         As Long    ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member.
  lpfnHook          As Long    ' Pointer to an FRHookProc hook procedure that can process messages intended for the dialog box.
  lpTemplateName    As String  ' Pointer to a null-terminated string that names the dialog box template resource in the module identified by the hInstance member.
End Type

' Common Dialog BROWSE FOLDER information type
Private Type BROWSEINFO
   hwndOwner        As Long    ' Handle of the owner window for the dialog box
   pidlRoot         As Long    ' Pointer to an item identifier list (an ITEMIDLIST structure) specifying the location of the "root" folder to browse from. Only the specified folder and its subfolders appear in the dialog box. This member can be NULL, and in that case, the name space root (the desktop folder) is used
   pszDisplayName   As String  ' Pointer to a buffer that receives the display name of the folder selected by the user. The size of this buffer is assumed to be MAX_PATH bytes
   lpszTitle        As String  ' Pointer to a null-terminated string that is displayed above the tree view control in the dialog box. This string can be used to specify instructions to the user
   ulFlags          As Long    ' Value specifying the types of folders to be listed in the dialog box as well as other options (See Constants)
   lpfnCALLBACK     As Long    ' Address an application-defined function that the dialog box calls when events occur. For more information, see the description of the BrowseCallbackProc function. This member can be NULL
   lParam           As Long    ' Application-defined value that the dialog box passes to the callback function (if one is specified)
   iImage           As Long    ' Variable that receives the image associated with the selected folder. The image is specified as an index to the system image list.
End Type

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

' Maximum buffer size constant
Public Const MAX_PATH = 260

' Translate color constant
Public Const CLR_INVALID = -1

' Operating System Constants
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Global Memory Function Constants
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000&
Public Const GMEM_LOCKCOUNT = &HFF
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_LOWER = GMEM_NOT_BANKED

' DEVMODE dmField Constants (PRINT)
Public Const DM_ORIENTATION = &H1&
Public Const DM_PAPERSIZE = &H2&
Public Const DM_PAPERLENGTH = &H4&
Public Const DM_PAPERWIDTH = &H8&
Public Const DM_SCALE = &H10&
Public Const DM_COPIES = &H100&
Public Const DM_DEFAULTSOURCE = &H200&
Public Const DM_PRINTQUALITY = &H400&
Public Const DM_COLOR = &H800&
Public Const DM_DUPLEX = &H1000&
Public Const DM_YRESOLUTION = &H2000&
Public Const DM_TTOPTION = &H4000&
Public Const DM_COLLATE = &H8000&
Public Const DM_FORMNAME = &H10000
Public Const DM_LOGPIXELS = &H20000
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const DM_DISPLAYFLAGS = &H200000
Public Const DM_DISPLAYFREQUENCY = &H400000
Public Const DM_ICMMETHOD = &H2000000   ' Windows 95 Only
Public Const DM_ICMINTENT = &H4000000   ' Windows 95 Only
Public Const DM_MEDIATYPE = &H8000000   ' Windows 95 Only
Public Const DM_DITHERTYPE = &H10000000 ' Windows 95 Only

' DEVMODE dmOrientation Constants (PRINT)
Public Const DMORIENT_LANDSCAPE = 2
Public Const DMORIENT_PORTRAIT = 1

' DEVMODE dmPaperSize Constants (PRINT)
Public Const DMPAPER_LETTER = 1        ' Letter, 8 1/2- by 11-inches
Public Const DMPAPER_LEGAL = 5         ' Legal, 8 1/2- by 14-inches
Public Const DMPAPER_A4 = 12           ' A4 Sheet, 210- by 297-millimeters
Public Const DMPAPER_CSHEET = 24       ' C Sheet, 17- by 22-inches
Public Const DMPAPER_DSHEET = 25       ' D Sheet, 22- by 34-inches
Public Const DMPAPER_ESHEET = 26       ' E Sheet, 34- by 44-inches
Public Const DMPAPER_LETTERSMALL = 2   ' Letter Small, 8 1/2- by 11-inches
Public Const DMPAPER_TABLOID = 3       ' Tabloid, 11- by 17-inches
Public Const DMPAPER_LEDGER = 4        ' Ledger, 17- by 11-inches
Public Const DMPAPER_STATEMENT = 6     ' Statement, 5 1/2- by 8 1/2-inches
Public Const DMPAPER_EXECUTIVE = 7     ' Executive, 7 1/4- by 10 1/2-inches
Public Const DMPAPER_A3 = 8            ' A3 sheet, 297- by 420-millimeters
Public Const DMPAPER_A4SMALL = 10      ' A4 small sheet, 210- by 297-millimeters
Public Const DMPAPER_A5 = 11           ' A5 sheet, 148- by 210-millimeters
Public Const DMPAPER_B4 = 12           ' B4 sheet, 250- by 354-millimeters
Public Const DMPAPER_B5 = 13           ' B5 sheet, 182- by 257-millimeter paper
Public Const DMPAPER_FOLIO = 14        ' Folio, 8 1/2- by 13-inch paper
Public Const DMPAPER_QUARTO = 15       ' Quarto, 215- by 275-millimeter paper
Public Const DMPAPER_10X14 = 16        ' 10- by 14-inch sheet
Public Const DMPAPER_11X17 = 17        ' 11- by 17-inch sheet
Public Const DMPAPER_NOTE = 18         ' Note, 8 1/2- by 11-inches
Public Const DMPAPER_ENV_9 = 19        ' #9 Envelope, 3 7/8- by 8 7/8-inches
Public Const DMPAPER_ENV_10 = 20       ' #10 Envelope, 4 1/8- by 9 1/2-inches
Public Const DMPAPER_ENV_11 = 21       ' #11 Envelope, 4 1/2- by 10 3/8-inches
Public Const DMPAPER_ENV_12 = 22       ' #12 Envelope, 4 3/4- by 11-inches
Public Const DMPAPER_ENV_14 = 23       ' #14 Envelope, 5- by 11 1/2-inches
Public Const DMPAPER_ENV_DL = 27       ' DL Envelope, 110- by 220-millimeters
Public Const DMPAPER_ENV_C5 = 28       ' C5 Envelope, 162- by 229-millimeters
Public Const DMPAPER_ENV_C3 = 29       ' C3 Envelope,  324- by 458-millimeters
Public Const DMPAPER_ENV_C4 = 30       ' C4 Envelope,  229- by 324-millimeters
Public Const DMPAPER_ENV_C6 = 31       ' C6 Envelope,  114- by 162-millimeters
Public Const DMPAPER_ENV_C65 = 32      ' C65 Envelope, 114- by 229-millimeters
Public Const DMPAPER_ENV_B4 = 33       ' B4 Envelope,  250- by 353-millimeters
Public Const DMPAPER_ENV_B5 = 34       ' B5 Envelope,  176- by 250-millimeters
Public Const DMPAPER_ENV_B6 = 35       ' B6 Envelope,  176- by 125-millimeters
Public Const DMPAPER_ENV_ITALY = 36    ' Italy Envelope, 110- by 230-millimeters
Public Const DMPAPER_ENV_MONARCH = 37  ' Monarch Envelope, 3 7/8- by 7 1/2-inches
Public Const DMPAPER_ENV_PERSONAL = 38 ' 6 3/4 Envelope, 3 5/8- by 6 1/2-inches
Public Const DMPAPER_FANFOLD_US = 39   ' US Std Fanfold, 14 7/8- by 11-inches
Public Const DMPAPER_FANFOLD_STD_GERMAN = 40 ' German Std Fanfold, 8 1/2- by 12-inches
Public Const DMPAPER_FANFOLD_LGL_GERMAN = 41 ' German Legal Fanfold, 8 1/2- by 13-inches
Public Const DMPAPER_USER = 256        ' User defined

' DEVMODE dmQuality Constants
Public Const DMRES_HIGH = (-4)         ' High quality
Public Const DMRES_MEDIUM = (-3)       ' Mediaum quality
Public Const DMRES_LOW = (-2)          ' Loq quality
Public Const DMRES_DRAFT = (-1)        ' Rough draft quality

' DEVMODE dmColor Constants
Public Const DMCOLOR_COLOR = 2         ' Color printing
Public Const DMCOLOR_MONOCHROME = 1    ' Black and white printing

' DEVMODE dmDuplex Constants  (Selects duplex or double-sided printing for printers capable of duplex printing)
Public Const DMDUP_SIMPLEX = 1
Public Const DMDUP_HORIZONTAL = 3
Public Const DMDUP_VERTICAL = 2

' DEVMODE dmTTOption Constants
Public Const DMTT_BITMAP = 1        ' Prints TrueType fonts as graphics. This is the default action for dot-matrix printers.
Public Const DMTT_DOWNLOAD = 2      ' Downloads TrueType fonts as soft fonts. This is the default action for Hewlett-Packard printers that use Printer Control Language (PCL).
Public Const DMTT_SUBDEV = 3        ' Substitute device fonts for TrueType fonts. This is the default action for PostScript® printers.

' DEVMODE dmCollate Constants
Public Const DMCOLLATE_TRUE = 1     ' Collate when printing multiple copies.
Public Const DMCOLLATE_FALSE = 0    ' Do not collate when printing multiple copies.

' DEVMODE dmDisplayFlags Constants
Public Const DM_GRAYSCALE = &H1     ' Specifies that the display is a noncolor device. If this flag is not set, color is assumed.
Public Const DM_INTERLACED = &H2    ' Specifies that the display mode is interlaced. If the flag is not set, noninterlaced is assumed.

' DEVMODE dmICMMethod Constants
Public Const DMICMMETHOD_NONE = 1   ' Windows 95 only: Specifies that ICM is disabled.
Public Const DMICMMETHOD_SYSTEM = 2 ' Windows 95 only: Specifies that ICM is handled by Windows.
Public Const DMICMMETHOD_DRIVER = 3 ' Windows 95 only: Specifies that ICM is handled by the device driver.
Public Const DMICMMETHOD_DEVICE = 4 ' Windows 95 only: Specifies that ICM is handled by the destination device.
Public Const DMICMMETHOD_USER = 256 ' Windows 95 only: User defined

' DEVMODE dmICMIntent Constants
Public Const DMICM_SATURATE = 1     ' Windows 95 only: Color matching should optimize for color saturation. This value is the most appropriate choice for business graphs when dithering is not desired.
Public Const DMICM_CONTRAST = 2     ' Windows 95 only: Color matching should optimize for color contrast. This value is the most appropriate choice for scanned or photographic images when dithering is desired.
Public Const DMICM_COLORMETRIC = 3  ' Windows 95 only: Color matching should optimize to match the exact color requested. This value is most appropriate for use with business logos or other images when an exact color match is desired.
Public Const DMICM_USER = 256       ' Windows 95 only: User defined

' DEVMODE dmMediaType Constants
Public Const DMMEDIA_STANDARD = 1   ' Windows 95 only: Plain paper.
Public Const DMMEDIA_GLOSSY = 2     ' Windows 95 only: Glossy paper.
Public Const DMMEDIA_TRANSPARNT = 3 ' Windows 95 only: Transparent film.
Public Const DMMEDIA_USER = 256     ' Windows 95 only: User defined

' DEVMODE dmDitherType Constants
Public Const DMDITHER_NONE = 1      ' Windows 95 only: No dithering.
Public Const DMDITHER_COARSE = 2    ' Windows 95 only: Dithering with a coarse brush.
Public Const DMDITHER_FINE = 3      ' Windows 95 only: Dithering with a fine brush.
Public Const DMDITHER_LINEART = 4   ' Windows 95 only: Line art dithering, a special dithering method that produces well defined borders between black, white, and gray scalings. It is not suitable for images that include continuous graduations in intensisty and hue such as scanned photographs.
Public Const DMDITHER_GRAYSCALE = 5 ' Windows 95 only: Device does grayscaling.
Public Const DMDITHER_USER = 256    ' Windows 95 only: User defined

' LOGFONT Height Constants
Public Const FW_DONTCARE = 0
Public Const FW_THIN = 100
Public Const FW_EXTRALIGHT = 200
Public Const FW_ULTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400
Public Const FW_REGULAR = 400
Public Const FW_MEDIUM = 500
Public Const FW_SEMIBOLD = 600
Public Const FW_DEMIBOLD = 600
Public Const FW_BOLD = 700
Public Const FW_EXTRABOLD = 800
Public Const FW_ULTRABOLD = 800
Public Const FW_HEAVY = 900
Public Const FW_BLACK = 900

' Predefined LOGFONT Character Set Constants
Public Const ANSI_CHARSET = 0
Public Const DEFAULT_CHARSET = 1
Public Const SYMBOL_CHARSET = 2
Public Const SHIFTJIS_CHARSET = 128
Public Const HANGEUL_CHARSET = 129
Public Const HANGUL_CHARSET = 129
Public Const GB2312_CHARSET = 134
Public Const CHINESEBIG5_CHARSET = 136
Public Const OEM_CHARSET = 255
Public Const JOHAB_CHARSET = 130       ' Windows 95 Only
Public Const HEBREW_CHARSET = 177      ' Windows 95 Only
Public Const ARABIC_CHARSET = 178      ' Windows 95 Only
Public Const GREEK_CHARSET = 161       ' Windows 95 Only
Public Const TURKISH_CHARSET = 162     ' Windows 95 Only
Public Const VIETNAMESE_CHARSET = 163  ' Windows 95 Only
Public Const THAI_CHARSET = 222        ' Windows 95 Only
Public Const EASTEUROPE_CHARSET = 238  ' Windows 95 Only
Public Const RUSSIAN_CHARSET = 204     ' Windows 95 Only
Public Const MAC_CHARSET = 77          ' Windows 95 Only
Public Const BALTIC_CHARSET = 186      ' Windows 95 Only

' LOGFONT Output Precision Constants
Public Const OUT_CHARACTER_PRECIS = 2  ' Not used.
Public Const OUT_DEFAULT_PRECIS = 0    ' Specifies the default font mapper behavior.
Public Const OUT_DEVICE_PRECIS = 5     ' Instructs the font mapper to choose a Device font when the system contains multiple fonts with the same name.
Public Const OUT_OUTLINE_PRECIS = 8    ' Windows NT: This value instructs the font mapper to choose from TrueType and other outline-based fonts.  Windows 95: This value is not used.
Public Const OUT_RASTER_PRECIS = 6     ' Instructs the font mapper to choose a raster font when the system contains multiple fonts with the same name.
Public Const OUT_STRING_PRECIS = 1     ' This value is not used by the font mapper, but it is returned when raster fonts are enumerated.
Public Const OUT_STROKE_PRECIS = 3     ' Windows NT: This value is not used by the font mapper, but it is returned when TrueType, other outline-based fonts, and vector fonts are enumerated.  Windows 95: This value is used to map vector fonts, and is returned when TrueType or vector fonts are enumerated.
Public Const OUT_TT_ONLY_PRECIS = 7    ' Instructs the font mapper to choose from only TrueType fonts. If there are no TrueType fonts installed in the system, the font mapper returns to default behavior.
Public Const OUT_TT_PRECIS = 4         ' Instructs the font mapper to choose a TrueType font when the system contains multiple fonts with the same name.

' LOGFONT Clip Precision Constants
Public Const CLIP_DEFAULT_PRECIS = 0   ' Specifies default clipping behavior.
Public Const CLIP_CHARACTER_PRECIS = 1 ' Not used.
Public Const CLIP_STROKE_PRECIS = 2    ' Not used by the font mapper, but is returned when raster, vector, or TrueType fonts are enumerated.  Windows NT: For compatibility, this value is always returned when enumerating fonts.
Public Const CLIP_MASK = &HF           ' Not used.
Public Const CLIP_EMBEDDED = 128       ' You must specify this flag to use an embedded read-only font.
Public Const CLIP_LH_ANGLES = 16       ' When this value is used, the rotation for all fonts depends on whether the orientation of the coordinate system is left-handed or right-handed. If not used, device fonts always rotate counterclockwise, but the rotation of other fonts is dependent on the orientation of the coordinate system.  For more information about the orientation of coordinate systems, see the description of the nOrientation parameter
Public Const CLIP_TT_ALWAYS = 32       ' Not used.

' LOGFONT Quality Constants
Public Const DEFAULT_QUALITY = 0       ' Appearance of the font does not matter.
Public Const DRAFT_QUALITY = 1         ' Appearance of the font is less important than when PROOF_QUALITY is used. For GDI raster fonts, scaling is enabled, which means that more font sizes are available, but the quality may be lower. Bold, italic, underline, and strikeout fonts are synthesized if necessary.
Public Const PROOF_QUALITY = 2         ' Character quality of the font is more important than exact matching of the logical-font attributes. For GDI raster fonts, scaling is disabled and the font closest in size is chosen. Although the chosen font size may not be mapped exactly when PROOF_QUALITY is used, the quality of the font is high and there is no distortion of appearance. Bold, italic, underline, and strikeout fonts are synthesized if necessary.

' LOGFONT Pitch Constants (Specifies the pitch and family of the font)
' The two low-order bits specify the pitch of the font and can be one of the following values:
Public Const DEFAULT_PITCH = 0
Public Const FIXED_PITCH = 1
Public Const VARIABLE_PITCH = 2

' LOGFONT Family Constants
' Bits 4 through 7 of the member specify the font family and can be one of the following values:
Public Const FF_DECORATIVE = 80        ' Novelty fonts. Old English is an example.
Public Const FF_DONTCARE = 0           ' Don’t care or don’t know.
Public Const FF_MODERN = 48            ' Fonts with constant stroke width (monospace), with or without serifs. Monospace fonts are usually modern. Pica, Elite, and CourierNew® are examples.
Public Const FF_ROMAN = 16             ' Fonts with variable stroke width (proportional) and with serifs. MS® Serif is an example.
Public Const FF_SCRIPT = 64            ' Fonts designed to look like handwriting. Script and Cursive are examples.
Public Const FF_SWISS = 32             ' Fonts with variable stroke width (proportional) and without serifs. MS® Sans Serif is an example.

' CHOOSEFONT Font Type Constants
Public Const BOLD_FONTTYPE = &H100     ' The font weight is bold. This information is duplicated in the lfWeight member of the LOGFONT structure and is equivalent to FW_BOLD.
Public Const ITALIC_FONTTYPE = &H200   ' The italic font attribute is set. This information is duplicated in the lfItalic member of the LOGFONT structure.
Public Const PRINTER_FONTTYPE = &H4000 ' The font is a printer font.
Public Const REGULAR_FONTTYPE = &H400  ' The font weight is normal. This information is duplicated in the lfWeight member of the LOGFONT structure and is equivalent to FW_REGULAR.
Public Const SCREEN_FONTTYPE = &H2000  ' The font is a screen font.
Public Const SIMULATED_FONTTYPE = &H8000& ' The font is simulated by the graphics device interface (GDI).


'----------------------------------------------------------------------------------
'                       Common Dialog Flag Codes
'----------------------------------------------------------------------------------

' Common Dialog OPEN/SAVE Flag Constants (See flag documentation below)
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHOWHELP = &H10

' Common Dialog COLOR Flag Constants (See flag documentation below)
Public Const CC_ENABLEHOOK = &H10
Public Const CC_ENABLETEMPLATE = &H20
Public Const CC_ENABLETEMPLATEHANDLE = &H40
Public Const CC_FULLOPEN = &H2
Public Const CC_PREVENTFULLOPEN = &H4
Public Const CC_RGBINIT = &H1
Public Const CC_SHOWHELP = &H8

' Common Dialog FONT Flag Constants (See flag documentation below)
Public Const CF_APPLY = &H200&
Public Const CF_ANSIONLY = &H400&
Public Const CF_TTONLY = &H40000
Public Const CF_EFFECTS = &H100&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOSIZESEL = &H200000
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOVERTFONTS = &H1000000
Public Const CF_PRINTERFONTS = &H2
Public Const CF_SCALABLEONLY = &H20000
Public Const CF_SCREENFONTS = &H1
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_SHOWHELP = &H4&
Public Const CF_USESTYLE = &H80&
Public Const CF_WYSIWYG = &H8000&
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)

' Common Dialog PRINT Flag Constants (See flag documentation below)
Public Const PD_ALLPAGES = &H0
Public Const PD_COLLATE = &H10
Public Const PD_DISABLEPRINTTOFILE = &H80000
Public Const PD_ENABLEPRINTHOOK = &H1000
Public Const PD_ENABLEPRINTTEMPLATE = &H4000
Public Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Public Const PD_ENABLESETUPHOOK = &H2000
Public Const PD_ENABLESETUPTEMPLATE = &H8000&
Public Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Public Const PD_HIDEPRINTTOFILE = &H100000
Public Const PD_NOPAGENUMS = &H8
Public Const PD_NOSELECTION = &H4
Public Const PD_NOWARNING = &H80
Public Const PD_PAGENUMS = &H2
Public Const PD_PRINTSETUP = &H40
Public Const PD_PRINTTOFILE = &H20
Public Const PD_RETURNDC = &H100
Public Const PD_RETURNDEFAULT = &H400
Public Const PD_RETURNIC = &H200
Public Const PD_SELECTION = &H1
Public Const PD_SHOWHELP = &H800
Public Const PD_USEDEVMODECOPIES = &H40000
Public Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000

' Common Dialog PAGE SETUP Flag Constants (See flag documentation below)
Public Const PSD_DEFAULTMINMARGINS = &H0
Public Const PSD_DISABLEMARGINS = &H10
Public Const PSD_DISABLEORIENTATION = &H100
Public Const PSD_DISABLEPAGEPAINTING = &H80000
Public Const PSD_DISABLEPAPER = &H200
Public Const PSD_DISABLEPRINTER = &H20
Public Const PSD_ENABLEPAGEPAINTHOOK = &H40000
Public Const PSD_ENABLEPAGESETUPHOOK = &H2000
Public Const PSD_ENABLEPAGESETUPTEMPLATE = &H8000&
Public Const PSD_ENABLEPAGESETUPTEMPLATEHANDLE = &H20000
Public Const PSD_INHUNDREDTHSOFMILLIMETERS = &H8
Public Const PSD_INTHOUSANDTHSOFINCHES = &H4
Public Const PSD_INWININIINTLMEASURE = &H0
Public Const PSD_MARGINS = &H2
Public Const PSD_MINMARGINS = &H1
Public Const PSD_NOWARNING = &H80
Public Const PSD_RETURNDEFAULT = &H400
Public Const PSD_SHOWHELP = &H800

' Common Dialog FIND / REPLACE Flag Constants (See flag documentation below)
Public Const FR_DIALOGTERM = &H40
Public Const FR_DOWN = &H1
Public Const FR_ENABLEHOOK = &H100
Public Const FR_ENABLETEMPLATE = &H200
Public Const FR_ENABLETEMPLATEHANDLE = &H2000
Public Const FR_FINDNEXT = &H8
Public Const FR_HIDEMATCHCASE = &H8000
Public Const FR_HIDEUPDOWN = &H4000
Public Const FR_HIDEWHOLEWORD = &H10000
Public Const FR_MATCHCASE = &H4
Public Const FR_NOMATCHCASE = &H800
Public Const FR_NOUPDOWN = &H400
Public Const FR_NOWHOLEWORD = &H1000
Public Const FR_REPLACE = &H10
Public Const FR_REPLACEALL = &H20
Public Const FR_SHOWHELP = &H80
Public Const FR_WHOLEWORD = &H2

' Common Dialog BROWSE FOLDER Constants (See flag documentation below)
Public Const BIF_BROWSEFORCOMPUTER = &H1000  'Only return computers. If the user selects anything other than a computer, the OK button is grayed.
Public Const BIF_BROWSEFORPRINTER = &H2000   'Only return printers. If the user selects anything other than a printer, the OK button is grayed.
Public Const BIF_DONTGOBELOWDOMAIN = &H2     'Do not include network folders below the domain level in the dialog box's tree view control.
Public Const BIF_RETURNFSANCESTORS = &H8     'Only return file system ancestors. An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy. If the user selects an ancestor of the root folder that is not part of the file system, the OK button is grayed.
Public Const BIF_RETURNONLYFSDIRS = &H1      'Only return file system directories. If the user selects folders that are not part of the file system, the OK button is grayed.
Public Const BIF_STATUSTEXT = &H4            'Include a status area in the dialog box. The callback function can set the status text by sending messages to the dialog box.
Public Const BIF_BROWSEINCLUDEFILES = &H4000 'Version 4.71. The browse dialog box will display files as well as folders.
Public Const BIF_EDITBOX = &H10              'Version 4.71. Include an edit control in the browse dialog box that allows the user to type the name of an item.
Public Const BIF_VALIDATE = &H20             'Version 4.71. If the user types an invalid name into the edit box, the browse dialog box will call the application's BrowseCallbackProc with the BFFM_VALIDATEFAILED message. This flag is ignored if BIF_EDITBOX is not specified.
Public Const BIF_BROWSEINCLUDEURLS = &H80    'Version 5.0.  The browse dialog box can display URLs. The BIF_USENEWUI and BIF_BROWSEINCLUDEFILES flags must also be set. If these three flags are not set, the browser dialog box will reject URLs. Even when these flags are set, the browse dialog box will only display URLs if the folder that contains the selected item supports them. When the folder's IShellFolder::GetAttributesOf method is called to request the selected item's attributes, the folder must set the SFGAO_FOLDER attribute flag. Otherwise, the browse dialog box will not display the URL.
Public Const BIF_NEWDIALOGSTYLE = &H40       'Version 5.0.  Use the new user interface. Setting this flag provides the user with a larger dialog box that can be resized. The dialog box has several new capabilities including: drag and drop capability within the dialog box, reordering, shortcut menus, new folders, delete, and other shortcut menu commands. To use this flag, you must call OleInitialize or CoInitialize before calling SHBrowseForFolder.
Public Const BIF_SHAREABLE = &H8000          'Version 5.0.  The browse dialog box can display shareable resources on remote systems. It is intended for applications that want to expose remote shares on a local system. The BIF_NEWDIALOGSTYLE flag must also be set.
Public Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX) 'Version 5.0. Use the new user interface, including an edit box. This flag is equivalent to BIF_EDITBOX | BIF_NEWDIALOGSTYLE. To use BIF_USENEWUI, you must call OleInitialize or CoInitialize before calling SHBrowseForFolder.
'Public Const BIF_NONEWFOLDERBUTTON = ?      'Version 6.0.  Do not include the New Folder button in the browse dialog box.
'Public Const BIF_NOTRANSLATETARGETS = ?     'Version 6.0.  When the selected item is a shortcut, return the PIDL of the shortcut itself rather than its target.
'Public Const BIF_UAHINT = ?                 'Version 6.0.  When combined with BIF_NEWDIALOGSTYLE, adds a usage hint to the dialog box in place of the edit box. BIF_EDITBOX overrides this flag.

' Common Dialog RUN Constants
Public Const RFF_NOBROWSE = &H1              ' Removes the browse button
Public Const RFF_NODEFAULT = &H2             ' No default item selected
Public Const RFF_CALCDIRECTORY = &H4         ' Calculates the working directory from the file name
Public Const RFF_NOLABEL = &H8               ' Removes the edit box label
Public Const RFF_NOSEPARATEMEM = &H20        ' Removes the Separate Memory Space check box (Windows NT only)

' Common Dialog PROPERTIES Constants
Public Const OPF_PRINTERNAME = 1
Public Const OPF_PATHNAME = 2

' Common Dialog REBOOT Constants
Public Const EWX_EXITANDEXECAPP = &H44&
Public Const EWX_REBOOTSYSTEM = &H43&
Public Const EWX_RESTARTWINDOWS = &H42&
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const IDYES = 6
Public Const IDNO = 7

' Common Dialog HELP Constants (See flag documentation below)
Public Enum HelpCommands
  HELP_COMMAND = &H102&     ' Run the specified help macro
  HELP_CONTEXT = &H1        ' Displays specified help
  HELP_CONTEXTPOPUP = &H8&  ' Displays specified help in a Pop-Up window
  HELP_FORCEFILE = &H9&     ' Display the first page of the help system

  HELP_HELPONHELP = &H4     ' Displays help on how to use the WinHelp system
  HELP_QUIT = &H2           ' Exits all WinHelp help systems
  
  HELP_KEY = &H101          ' Opens the help file to the "Index" tab
  HELP_TAB = &HF            ' Opens the help file to the "Contents" tab
  
' HELP_CONTENTS = &H3&      ' Doesn't apply
' HELP_CONTEXTMENU = &HA    ' Doesn't apply
' HELP_FINDER = &HB         ' Doesn't apply
' HELP_INDEX = &H3          ' Doesn't apply
' HELP_MULTIKEY = &H201&    ' Doesn't apply
' HELP_PARTIALKEY = &H105&  ' Doesn't apply
' HELP_SETCONTENTS = &H5&   ' Doesn't apply
' HELP_SETINDEX = &H5       ' Doesn't apply
' HELP_SETWINPOS = &H203&   ' Doesn't apply
End Enum

'----------------------------------------------------------------------------------
'                       Common Dialog Error Return Codes
'----------------------------------------------------------------------------------

' Cancel Error Constants
Public Const CDERR_CANCEL = 32755
Public Const CDERR_CANCELMSG = "Cancel was selected"

' General Common Dialog Return Codes
Public Const CDERR_NOERROR = 0
Public Const CDERR_DIALOGFAILURE = &HFFFF    ' The dialog box could not be created. The common dialog box function’s call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle.
Public Const CDERR_FINDRESFAILURE = &H6      ' The common dialog box function failed to find a specified resource.
Public Const CDERR_INITIALIZATION = &H2      ' The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available.
Public Const CDERR_LOADRESFAILURE = &H7      ' The common dialog box function failed to load a specified resource.
Public Const CDERR_LOADSTRFAILURE = &H5      ' The common dialog box function failed to load a specified string.
Public Const CDERR_LOCKRESFAILURE = &H8      ' The common dialog box function failed to lock a specified resource.
Public Const CDERR_MEMALLOCFAILURE = &H9     ' The common dialog box function was unable to allocate memory for internal structures.
Public Const CDERR_MEMLOCKFAILURE = &HA      ' The common dialog box function was unable to lock the memory associated with a handle.
Public Const CDERR_NOHINSTANCE = &H4         ' The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle.
Public Const CDERR_NOHOOK = &HB              ' The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure.
Public Const CDERR_NOTEMPLATE = &H3          ' The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template.
Public Const CDERR_REGISTERMSGFAIL = &HC     ' The RegisterWindowMessage function returned an error code when it was called by the common dialog box function.
Public Const CDERR_STRUCTSIZE = &H1          ' The lStructSize member of the initialization structure for the corresponding common dialog box is invalid.

' DLG_PrintDialog Function Return Codes
Public Const PDERR_CREATEICFAILURE = &H100A  ' The DLG_PrintDialog function failed when it attempted to create an information context.
Public Const PDERR_DEFAULTDIFFERENT = &H100C ' You called the DLG_PrintDialog function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer. (This error occurs when you store the DEVNAMES structure and the user changes the default printer by using the Control Panel.)
                                             ' To use the printer described by the DEVNAMES structure, clear the DN_DEFAULTPRN flag and call DLG_PrintDialog again. To use the default printer, replace the DEVNAMES structure (and the DEVMODE structure, if one exists) with NULL; and call DLG_PrintDialog again.
Public Const PDERR_DNDMMISMATCH = &H1009     ' The data in the DEVMODE and DEVNAMES structures describes two different printers.
Public Const PDERR_GETDEVMODEFAIL = &H1005   ' The printer driver failed to initialize a DEVMODE structure. (This error code applies only to printer drivers written for Windows versions 3.0 and later.)
Public Const PDERR_INITFAILURE = &H1006      ' The DLG_PrintDialog function failed during initialization, and there is no more specific extended error code to describe the failure. This is the generic default error code for the function.
Public Const PDERR_LOADDRVFAILURE = &H1004   ' The DLG_PrintDialog function failed to load the device driver for the specified printer.
Public Const PDERR_NODEFAULTPRN = &H1008     ' A default printer does not exist.
Public Const PDERR_NODEVICES = &H1007        ' No printer drivers were found.
Public Const PDERR_PARSEFAILURE = &H1002     ' The DLG_PrintDialog function failed to parse the strings in the [devices] section of the WIN.INI file.
Public Const PDERR_PRINTERNOTFOUND = &H100B  ' The [devices] section of the WIN.INI file did not contain an entry for the requested printer.
Public Const PDERR_RETDEFFAILURE = &H1003    ' The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was not NULL.
Public Const PDERR_SETUPFAILURE = &H1001     ' The DLG_PrintDialog function failed to load the required resources.

' DLG_ChooseFont Function Return Codes
Public Const CFERR_MAXLESSTHANMIN = &H2002   ' The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member.
Public Const CFERR_NOFONTS = &H2001          ' No fonts exist.

' DLG_GetOpenFileName / DLG_GetSaveFileName Function Return Codes
Public Const FNERR_BUFFERTOOSMALL = &H3003   ' The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the filename specified by the user. The first two bytes of the lpstrFile buffer contain an integer value specifying the size, in bytes (ANSI version) or characters (Unicode version), required to receive the full name.
Public Const FNERR_INVALIDFILENAME = &H3002  ' A filename is invalid.
Public Const FNERR_SUBCLASSFAILURE = &H3001  ' An attempt to subclass a list box failed because sufficient memory was not available.

' Constantced used with the FIND / REPLACE functionality
Private Const WM_CLOSE = &H10
Private Const GWL_WNDPROC = (-4)
Private Const HEAP_ZERO_MEMORY = &H8
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const HELPMSGSTRING = "commdlg_help"
Private FINDMESSAGE As Long
Private HELPMESSAGE As Long

' Variables used with the FIND/REPLACE functionality
Private ReturnFR As FINDREPLACE
Private TheMessage As MSG
Private hDialog As Long
Private lHeap As Long
Private OldProc As Long
Private arrFind() As Byte
Private arrReplace() As Byte

' Operating System Variables
Private Win_OS          As OSTypes
Private Win_Version     As String
Private Win_Build       As String
Private CantGetOSInfo   As Boolean

'----------------------------------------------------------------------------------

' Common Dialog API Declarationsf
Public Declare Function DLG_About Lib "SHELL32.DLL" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal strApp As String, ByVal strOther As String, ByVal hIcon As Long) As Long
Public Declare Function DLG_BrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolder" (lpbi As BROWSEINFO) As Long
Public Declare Function DLG_ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Public Declare Function DLG_ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Public Declare Function DLG_FindComputer Lib "SHELL32.DLL" Alias "#91" (ByVal pidlRoot As Long, ByVal pidlSavedSearch As Long) As Long
Public Declare Function DLG_FindFile Lib "SHELL32.DLL" Alias "#90" (ByVal pidlRoot As Long, ByVal pidlSavedSearch As Long) As Long
Public Declare Function DLG_FindText Lib "comdlg32.dll" Alias "FindTextA" (ByRef pFindreplace As Long) As Long
Public Declare Function DLG_FormatDrive Lib "SHELL32.DLL" Alias "SHFormatDrive" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Public Declare Function DLG_GetIcon Lib "SHELL32.DLL" Alias "#62" (ByVal hOwner As Long, ByVal szFilename As String, ByVal Reserved As Long, ByRef lpIconIndex As Long) As Long
Public Declare Function DLG_GetLastError Lib "comdlg32.dll" Alias "CommDlgExtendedError" () As Long
Public Declare Function DLG_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function DLG_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function DLG_PageSetupDialog Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
Public Declare Function DLG_PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG) As Long
Public Declare Function DLG_Properties Lib "SHELL32.DLL" Alias "#178" (ByVal hwndOwner As Long, ByVal uFlags As Long, ByVal lpstrName As String, ByVal lpstrParameters As String) As Long
Public Declare Function DLG_Reboot Lib "Shell32" Alias "#59" (ByVal hOwner As Long, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function DLG_ReplaceText Lib "comdlg32.dll" Alias "ReplaceTextA" (ByRef pFindreplace As Long) As Long
Public Declare Function DLG_Run Lib "SHELL32.DLL" Alias "#61" (ByVal hwndOwner As Long, ByVal hIcon As Long, ByVal lpstrDirectory As String, ByVal lpstrTitle As String, ByVal lpstrDescription As String, ByVal uFlags As Long) As Long
Public Declare Function DLG_ShutDown Lib "Shell32" Alias "#60" (ByVal hOwner As Long) As Long
Public Declare Function DLG_WinHelp_LNG Lib "USER32.DLL" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Declare Function DLG_WinHelp_STR Lib "USER32.DLL" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As String) As Long

' Other Related Windows API Declarations
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSource As Any, ByVal cbCopy As Long)
Public Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal hMem As Long)
Public Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CopyPointer2String Lib "KERNEL32" Alias "lstrcpyA" (ByVal NewString As String, ByVal OldString As Long) As Long
Public Declare Function DispatchMessage Lib "USER32" Alias "DispatchMessageA" (ByRef lpMSG As MSG) As Long
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetMessage Lib "USER32" Alias "GetMessageA" (ByRef lpMSG As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Function GetProcessHeap Lib "KERNEL32" () As Long
Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function HeapAlloc Lib "KERNEL32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "KERNEL32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByRef lpMem As Any) As Long
Public Declare Function IsDialogMessage Lib "USER32" Alias "IsDialogMessageA" (ByVal hDlg As Long, ByRef lpMSG As MSG) As Long
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pColorRef As Long) As Long
Public Declare Function RegisterWindowMessage Lib "USER32" Alias "RegisterWindowMessageA" (ByVal LPString As String) As Long
Public Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function SHGetIDListFromPath Lib "SHELL32.DLL" Alias "#162" (ByVal szPath As String) As Long
Public Declare Function TranslateMessage Lib "USER32" (ByRef lpMSG As MSG) As Long



'=============================================================================================================
' CD_ShowAbout
'
' Purpose :
' Shows the default Windows "About" screen with a custom titlebar caption,
' custom icon, and custom information.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' TitlebarCaption      Optional. New About dialog caption
' AppName              Optional. Application name to display
' OtherInfo            Optional. Other information to be displayed
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowAbout(ByVal OwnerHandle As Long, _
                             Optional ByVal TitlebarCaption As String, _
                             Optional ByVal AppName As String, _
                             Optional ByVal OtherInfo As String, _
                             Optional ByVal hIcon As Long) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  
  ' Test parameters
  If AppName = vbNullString Then
    AppName = Chr(0)
  ElseIf Len(AppName) > 31 Then
    AppName = Left(AppName, 31)
  End If
  If Len(OtherInfo) > 36 Then
    OtherInfo = Left(OtherInfo, 36)
  End If
  
  ' Combine the TitlebarCaption and the AppName
  AppName = TitlebarCaption & "#" & AppName
  
  ' Display the about box
  ReturnCode = DLG_About(OwnerHandle, AppName, OtherInfo, hIcon)
  
  ' Return results
  If ReturnCode = 0 Then
    CD_ShowAbout = False
  Else
    CD_ShowAbout = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowAbout function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowAbout = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowColor
'
' Purpose :
' Shows the standard Windows "Select Color" dialog.
'
' Compare to COMDLG32.OCX ShowColor method.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( CC_... )
' TheColor             Optional. Sets / recieves the color selected in the dialog
' RGB_Red              Optional. Recieves the RED value of the RGB return color
' RGB_Green            Optional. Recieves the GREEN value of the RGB return color
' RGB_Blue             Optional. Recieves the RED value of the RGB return color
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags     = Dialog return flags
'                    TheColor  = Selected Color
'                    RGB_Red   = RED value of "TheColor = RGB(R,G,B)"
'                    RGB_Green = GREEN value of "TheColor = RGB(R,G,B)"
'                    RGB_Blue  = BLUE value of "TheColor = RGB(R,G,B)"
'
'=============================================================================================================
Public Function CD_ShowColor(ByVal OwnerHandle As Long, _
                             Optional ByRef Flags As Long, _
                             Optional ByRef TheColor As Long, _
                             Optional ByRef RGB_Red As Byte, _
                             Optional ByRef RGB_Green As Byte, _
                             Optional ByRef RGB_Blue As Byte) As Boolean
On Error GoTo ErrorTrap
  
  Dim ColorInfo As CHOOSECOLOR
  Dim CustomColor As String
  Dim ReturnCode As Long
  
  ' Make sure the color is passed in the correct format
  TheColor = TranslateColor(TheColor)
  CustomColor = StrConv(TheColor, vbUnicode)
  
  With ColorInfo
    .lStructSize = Len(ColorInfo)  ' Specifies the length of this type/structure
    .hwndOwner = OwnerHandle       ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .rgbResult = TheColor          ' > Sets / recieves the color selected.  If the CC_RGBINIT flag is set, this the color initially selected when the dialog box is created. If the specified color value is not among the available colors, the system selects the nearest solid color available. If rgbResult is zero or CC_RGBINIT is not set, the initially selected color is black.
    .lpCustColors = CustomColor    ' > Sets / receives the custom RGB colors for the specified color
    .Flags = Flags                 ' > Sets / recieves the flags for the dialog
'------
    .hInstance = 0                 ' If CC_ENABLETEMPLATEHANDLE is set - hInstance is the handle of a memory object containing a dialog box template. If CC_ENABLETEMPLATE is set, hInstance identifies a module that contains a dialog box template named by the lpTemplateName member. If neither CC_ENABLETEMPLATEHANDLE nor CC_ENABLETEMPLATE is set, this member is ignored.
    .lCustData = 0                 ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the CHOOSECOLOR structure specified when the dialog was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnHook = 0                  ' Pointer to a CCHookProc hook procedure that can process messages intended for the dialog box. This member is ignored unless the CC_ENABLEHOOK flag is set in the Flags member.
    .lpTemplateName = vbNullString ' Pointer to a null-terminated string that names the dialog box template resource in the module identified by the hInstance member. This template is substituted for the standard dialog box template. For numbered dialog box resources, lpTemplateName can be a value returned by the MAKEINTRESOURCE macro. This member is ignored unless the CC_ENABLETEMPLATE flag is set in the Flags member.
  End With
  
  ' Display the dialog
  ReturnCode = DLG_ChooseColor(ColorInfo)
  
  ' Check if dialog canceled or error occured
  If GetLastError_CDLG(True, "DLG_ChooseColor") = True Or ReturnCode = 0 Then
    Flags = 0
    TheColor = 0
    RGB_Red = 0
    RGB_Green = 0
    RGB_Blue = 0
    CD_ShowColor = False
    
  ' Return values
  Else
    Flags = ColorInfo.Flags
    TheColor = ColorInfo.rgbResult
    Convert_LNG_RGB TheColor, RGB_Red, RGB_Green, RGB_Blue
    CD_ShowColor = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowColor function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowColor = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowFindComputer
'
' Purpose :
' Shows a find dialog box designed to locate a specified file.
'
' Param                Use
' ------------------------------------
' DefaultDir           Optional. The default location to start finding
' SavedSearch          Optional. Saved search file (.FND) path to use
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowFindFile(Optional ByVal DefaultDir As String, _
                                Optional ByVal SavedSearch As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  Dim IDList_Path As Long
  Dim IDList_Saved As Long
  
  ' Setup the IDLists from the paths provided
  If DefaultDir = "" Then
    DefaultDir = "C:\"
  End If
  IDList_Path = SHGetIDListFromPath(DefaultDir)
  
  If SavedSearch = "" Then
    IDList_Saved = 0
  Else
    IDList_Saved = SHGetIDListFromPath(SavedSearch)
  End If
  
  ' Display the dialog
  ReturnCode = DLG_FindFile(IDList_Path, IDList_Saved)
  
  ' Check if function call failed
  If ReturnCode = 0 Then
    CoTaskMemFree IDList_Path
    CoTaskMemFree IDList_Saved
    CD_ShowFindFile = False
  Else
    CoTaskMemFree IDList_Path
    CD_ShowFindFile = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowFindFile function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CoTaskMemFree IDList_Path
    CoTaskMemFree IDList_Saved
    CD_ShowFindFile = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowFindReplace
'
' Purpose :
' Shows the standard Windows "Find/Replace" dialog and allows you to use it via the
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( CC_... )
' TheColor             Optional. Sets / recieves the color selected in the dialog
' RGB_Red              Optional. Recieves the RED value of the RGB return color
' RGB_Green            Optional. Recieves the GREEN value of the RGB return color
' RGB_Blue             Optional. Recieves the RED value of the RGB return color
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags     = Dialog return flags
'                    TheColor  = Selected Color
'                    RGB_Red   = RED value of "TheColor = RGB(R,G,B)"
'                    RGB_Green = GREEN value of "TheColor = RGB(R,G,B)"
'                    RGB_Blue  = BLUE value of "TheColor = RGB(R,G,B)"
'
'=============================================================================================================
Public Sub CD_ShowFindReplace(ByVal OwnerHandle As Long, _
                              Optional ByVal Flags As Long, _
                              Optional ByVal FindString As String = "", _
                              Optional ByVal ShowReplace As Boolean = False, _
                              Optional ByVal ReplaceString As String = "")
On Error GoTo ErrorTrap
  
  Dim FRInfo As FINDREPLACE
  
  ' If there is already one "Find/Replace" dialog open, exit
  If hDialog > 0 Then
    Exit Sub
  End If
  
  ' If no owner was specified, exit... the Find/Replace dialog requires an owner to send Windows messages to
  If OwnerHandle = 0 Then
    Err.Raise -1, "modCOMDLG32.bas - CD_ShowFindReplace", "Owner handle must be a valid window handle."
    Exit Sub
  End If
  
  ' Make sure the Windows messages to be used to trap dialog events are defined
  FINDMESSAGE = RegisterWindowMessage(FINDMSGSTRING)
  HELPMESSAGE = RegisterWindowMessage(HELPMSGSTRING)
  
  ' Setup the string BYTE arrays with the right data
  arrFind = StrConv(FindString & Chr(0), vbFromUnicode)
  arrReplace = StrConv(ReplaceString & Chr(0), vbFromUnicode)
  
  With FRInfo
    .lStructSize = Len(FRInfo)                 ' Size of this type / structure
    .hwndOwner = OwnerHandle                   ' This can NOT be NULL. Identifies the window that owns the dialog box. The window procedure of the specified window receives FINDMSGSTRING messages from the dialog box. This member can be any valid window handle.
    .hInstance = 0                             ' If the FR_ENABLETEMPLATEHANDLE flag is set in the Flags member, hInstance is the handle of a memory object containing a dialog box template. If the FR_ENABLETEMPLATE flag is set, hInstance identifies a module that contains a dialog box template named by the lpTemplateName member. If neither flag is set, this member is ignored.
    .Flags = Flags                             ' A set of bit flags that you can use to initialize the dialog box. The dialog box sets these flags when it sends the FINDMSGSTRING registered message to indicate the user’s input.
    .lpstrFindWhat = VarPtr(arrFind(0))        ' Pointer to a buffer that a FINDMSGSTRING message uses to pass the null terminated search string that the user typed in the “Find What:” edit control. You must dynamically allocate the buffer or use a global or static array so it does not go out of scope before the dialog box closes. The buffer should be at least 80 characters long. If the buffer contains a string when you initialize the dialog box, the string is displayed in the “Find What:” edit control.
                                               ' If a FINDMSGSTRING message specifies the FR_FINDNEXT flag, lpstrFindWhat contains the string to search for. The FR_DOWN, FR_WHOLEWORD, and FR_MATCHCASE flags indicate the direction and type of search. If a FINDMSGSTRING message specifies the FR_REPLACE or FR_REPLACE flags, lpstrFindWhat contains the string to be replaced.
    .lpstrReplaceWith = VarPtr(arrReplace(0))  ' Pointer to a buffer that a FINDMSGSTRING message uses to pass the null terminated replacement string that the user typed in the “Replace With:” edit control. You must dynamically allocate the buffer or use a global or static array so it does not go out of scope before the dialog box closes. If the buffer contains a string when you initialize the dialog box, the string is displayed in the “Replace With:” edit control.
                                               ' If a FINDMSGSTRING message specifies the FR_REPLACE or FR_REPLACEALL flags, lpstrReplaceWith contains the replacement string .
                                               ' The DLG_FindText function ignores this member.
    .wFindWhatLen = MAX_PATH                   ' Specifies the length, in bytes, of the buffer pointed to by the lpstrFindWhat member.
    .wReplaceWithLen = MAX_PATH                ' Specifies the length, in bytes, of the buffer pointed to by the lpstrReplaceWith member.
    .lCustData = 0                             ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the FINDREPLACE structure specified when the dialog was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnHook = 0                              ' Pointer to an FRHookProc hook procedure that can process messages intended for the dialog box. This member is ignored unless the FR_ENABLEHOOK flag is set in the Flags member.
                                               ' If the hook procedure returns FALSE in response to the WM_INITDIALOG message, the hook procedure must display the dialog box or else the dialog box will not be shown. To do this, first perform any other paint operations, and then call the ShowWindow and UpdateWindow functions.
    .lpTemplateName = ""                       ' Pointer to a null-terminated string that names the dialog box template resource in the module identified by the hInstance member. This template is substituted for the standard dialog box template. For numbered dialog box resources, this can be a value returned by the MAKEINTRESOURCE macro. This member is ignored unless the FR_ENABLETEMPLATE flag is set in the Flags member.
  End With
  
  ' The HeapAlloc function allocates a block of memory from a heap. The allocated memory is not movable.
  lHeap = HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, FRInfo.lStructSize)
  
  ' Move the structure just created to the new heap block
  CopyMemory ByVal lHeap, FRInfo, Len(FRInfo)
  
  ' Subclass events sent to the owner of the dialog - this lets you know what's going on with the dialog
  OldProc = SetWindowLong(OwnerHandle, GWL_WNDPROC, AddressOf FindReplaceProc)
  
  ' If the user wants to show the "Find/Replace" dialog, show it... otherwise just show the "Find" dialog
  If ShowReplace = True Then
    hDialog = DLG_ReplaceText(ByVal lHeap)
  Else
    hDialog = DLG_FindText(ByVal lHeap)
  End If
  
  ' Start looking for messages sent to the Find/Replace dialog that just appeared
  MessageLoop
  
  Exit Sub
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowFindReplace function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
  End If
  
End Sub

'=============================================================================================================
' CD_ShowFindComputer
'
' Purpose :
' Shows a find dialog box designed to locate a computer name.
'
' Param                Use
' ------------------------------------
' ( None )
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowFindComputer() As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  
  ' Display the dialog
  ReturnCode = DLG_FindComputer(0, 0)
  
  ' Check if function call failed
  If ReturnCode = 0 Then
    CD_ShowFindComputer = False
  Else
    CD_ShowFindComputer = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowFindComputer function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowFindComputer = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowFolder
'
' Purpose :
' Shows the standard Windows "Select Folder" dialog.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( BIF_... )
' FolderName           Optional. Recieves the selected folder (without path)
' FolderPath           Optional. Sets / recieves the selected folder path
' Prompt               Optional. Sets the dialog's prompt text
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags      = Dialog return flags
'                    FolderName = Selected folder's name (w/o path)
'                    FolderPath = Selected folder's full path
'
'=============================================================================================================
Public Function CD_ShowFolder(ByVal OwnerHandle As Long, _
                              Optional ByRef Flags As Long, _
                              Optional ByRef FolderName As String, _
                              Optional ByRef FolderPath As String, _
                              Optional ByVal Prompt As String) As Boolean
On Error GoTo ErrorTrap
  
  ' Declare variables to be used
  Dim BrwsInfo As BROWSEINFO
  Dim ReturnCode As Long
  
  ' Set default values
  If Flags = 0 Then
    Flags = BIF_RETURNONLYFSDIRS
  End If
  If FolderPath = "" Then
    FolderPath = vbNullChar
  End If
  If Prompt = "" Then
    Prompt = "Select Folder:"
  End If
  
  ' Add NULL character to the end of the titlebar caption
  Prompt = Prompt & vbNullChar
  
  ' Make sure starting path is valid
  If FolderPath = "" Then
    FolderPath = vbNullChar
  ElseIf (Flags And BIF_BROWSEFORCOMPUTER) <> 0 Or (Flags And BIF_BROWSEFORPRINTER) <> 0 Then
    FolderPath = vbNullChar
  End If
  
  ' Initialise variables
  With BrwsInfo
    .hwndOwner = OwnerHandle                    ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .pidlRoot = SHGetIDListFromPath(FolderPath) ' Specifies the "root" to browse from.  Can be NULL
    .pszDisplayName = String(MAX_PATH, Chr(0))  ' Pointer to a buffer that receives the display name of the folder selected by the user
    .lpszTitle = Prompt                         ' Pointer to a null-terminated string that is displayed above the tree view control in the dialog box. This string can be used to specify instructions to the user
    .ulFlags = Flags                            ' Value specifying the types of folders to be listed in the dialog box as well as other options (See Constants)
    
    .lpfnCALLBACK = 0                           ' Address an application-defined function that the dialog box calls when events occur. For more information, see the description of the BrowseCallbackProc function. This member can be NULL.
    .lParam = 0                                 ' Application-defined value that the dialog box passes to the callback function (if one is specified).
    .iImage = 0                                 ' Variable that receives the image associated with the selected folder. The image is specified as an index to the system image list.
  End With
  
  ' Display the dialog
  ReturnCode = DLG_BrowseForFolder(BrwsInfo)
  
  ' Check if dialog canceled or error occured
  If ReturnCode = 0 Then
    Flags = 0
    FolderName = ""
    FolderPath = ""
    CD_ShowFolder = False
    
  ' Return the information
  Else
    Flags = BrwsInfo.ulFlags
    If BrwsInfo.pszDisplayName <> "" Then
      FolderName = BrwsInfo.pszDisplayName
      FolderName = Left(FolderName, InStr(FolderName, Chr(0)) - 1)
    Else
      FolderName = ""
    End If
    FolderPath = String(MAX_PATH, Chr(0))
    If SHGetPathFromIDList(ReturnCode, FolderPath) <> 0 Then
      FolderPath = Left(FolderPath, InStr(FolderPath, Chr(0)) - 1)
      CoTaskMemFree ReturnCode
      ReturnCode = 0
    End If
    CD_ShowFolder = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowFolder function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowFolder = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowFont
'
' Purpose :
' Shows the standard Windows "Select Font" dialog.
'
' Compare to COMDLG32.OCX ShowFont method.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( CF_... )
' FontName             Optional. Sets / recieves the selected Font Name
' FontSize             Optional. Sets / recieves the selected Font Size
' FontBold             Optional. Sets / recieves the selected Font Bold
' FontItalic           Optional. Sets / recieves the selected Font Italic
' FontStrikeThru       Optional. Sets / recieves the selected Font StrikeThru
' FontUnderline        Optional. Sets / recieves the selected Font Underline
' Color                Optional. Sets / recieves the selected Font Color
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags          = Dialog return flags
'                    FontName       = The selected Font Name
'                    FontSize       = The selected Font Size
'                    FontBold       = The selected Font Bold
'                    FontItalic     = The selected Font Italic
'                    FontStrikeThru = The selected Font StrikeThru
'                    FontUnderline  = The selected Font Underline
'                    Color          = The selected Font Color
'
'=============================================================================================================
Public Function CD_ShowFont(ByVal OwnerHandle As Long, _
                            Optional ByRef Flags As Long, _
                            Optional ByRef fontName As String, _
                            Optional ByRef FontSize As Long, _
                            Optional ByRef FontBold As Boolean, _
                            Optional ByRef FontItalic As Boolean, _
                            Optional ByRef FontStrikethru As Boolean, _
                            Optional ByRef FontUnderline As Boolean, _
                            Optional ByRef Color As Long) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  Dim FontInfo As CHOOSEFONT
  Dim TheLogFont As LOGFONT
  Dim FontType As Integer
  Dim StyleBuffer As String
  Dim Temp_FontBold As Byte
  Dim Temp_FontItalic As Byte
  Dim Temp_FontStrikeThru As Byte
  Dim Temp_FontUnderline As Byte
  Dim BoldFlag As Long
  Dim hLogFont As Long ' Handle to TheLogFont memory buffer
  Dim pLogFont As Long ' Pointer to TheLogFont memory buffer
  
  ' Make sure that the flags has CF_INITTOLOGFONTSTRUCT in it to force the
  ' initial settings to show up when the dialog appears
  If (Flags And CF_INITTOLOGFONTSTRUCT) = 0 Then
    Flags = (Flags Or CF_INITTOLOGFONTSTRUCT)
  End If
  
  ' Set the style buffer
  StyleBuffer = String(MAX_PATH, Chr(0))
  
  ' Set the FontBold information
  If FontBold = True Then
    Temp_FontBold = 1
    FontType = FontType Or BOLD_FONTTYPE
    BoldFlag = FW_BOLD
  Else
    Temp_FontBold = 0
    FontType = FontType Or REGULAR_FONTTYPE
    BoldFlag = FW_NORMAL
  End If
  
  ' Set the FontItalic information
  If FontItalic = True Then
    Temp_FontItalic = 1
    FontType = FontType Or ITALIC_FONTTYPE
  Else
    Temp_FontItalic = 0
  End If
  
  ' Set the FontStrikeThru information
  If FontStrikethru = True Then
    Temp_FontStrikeThru = 1
  Else
    Temp_FontStrikeThru = 0
  End If
  
  ' Set the FontUnderline information
  If FontUnderline = True Then
    Temp_FontUnderline = 1
  Else
    Temp_FontUnderline = 0
  End If
  
  ' Make sure the color is passed in the correct way
  Color = TranslateColor(Color)
  
  ' Setup the LOGFONT variable to be passed
  With TheLogFont
    .lfHeight = 0
    .lfWeight = 0
    .lfEscapement = 0
    .lfOrientation = 0
    .lfWeight = BoldFlag
    .lfItalic = Temp_FontItalic
    .lfUnderline = Temp_FontUnderline
    .lfStrikeOut = Temp_FontStrikeThru
    .lfCharSet = DEFAULT_CHARSET
    .lfOutPrecision = OUT_DEFAULT_PRECIS
    .lfClipPrecision = CLIP_DEFAULT_PRECIS
    .lfQuality = DEFAULT_QUALITY
    .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    .lfFaceName = fontName & vbNullChar
  End With
  
  ' Copy the LOGFONT variable to a memory address to get it's pointer
  hLogFont = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(TheLogFont)) ' Allocate a memory buffer to store the LOGFONT variable in
  pLogFont = GlobalLock(hLogFont)                                         ' Lock the allocated memory buffer just created and get a pointer to it
  If pLogFont <> 0 Then
    CopyMemory ByVal pLogFont, TheLogFont, Len(TheLogFont)                  ' Copy structure's content into the newly created memory buffer
  End If
  
  ' Setup the CHOOSEFONT variable to be passed
  With FontInfo
    .lStructSize = Len(FontInfo) ' Specifies the length of this type/structure
    .hwndOwner = OwnerHandle     ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .hDC = Printer.hDC           ' Identifies the device context (or information context) of the printer whose fonts will be listed in the dialog box. This member is used only if the Flags member specifies the CF_PRINTERFONTS or CF_BOTH flag; otherwise, this member is ignored.
    .lpLogFont = pLogFont        ' > Sets / recieves a LOGFONT structure that defines the font. If you set the CF_INITTOLOGFONTSTRUCT flag in the Flags member and initialize the LOGFONT members, the ChooseFont function initializes the dialog box with a font that is the closest possible match. If the user clicks the OK button, ChooseFont sets the members of the LOGFONT structure based on the user’s selections.
    .iPointSize = FontSize * 10  ' > Sets / recieves the size of the selected font, in units of 1/10 of a point. The ChooseFont function sets this value after the user closes the dialog box.
    .Flags = Flags               ' > Sets / recieves the flags for the dialog
    .RGBColors = Color           ' > Sets / recieves the font color if the CF_EFFECTS flag is set
    .lpszStyle = StyleBuffer     ' Pointer to a buffer that contains style data. If the CF_USESTYLE flag is specified, ChooseFont uses the data in this buffer to initialize the font style combo box. When the user closes the dialog box, ChooseFont copies the string in the font style combo box into this buffer.
    .nFontType = FontType        ' Sets the type of the selected font when ChooseFont returns.
    .nSizeMin = 6                ' Sets the minimum point size a user can select. ChooseFont recognizes this member only if the CF_LIMITSIZE flag is specified.
    .nSizeMax = 72               ' Sets the maximum point size a user can select. ChooseFont recognizes this member only if the CF_LIMITSIZE flag is specified.
'------
    .hInstance = 0               ' If the CF_ENABLETEMPLATEHANDLE flag is set in the Flags member, hInstance is the handle of a memory object containing a dialog box template. If the CF_ENABLETEMPLATE flag is set, hInstance identifies a module that contains a dialog box template named by the lpTemplateName member. If neither CF_ENABLETEMPLATEHANDLE nor CF_ENABLETEMPLATE is set, this member is ignored.
    .lCustData = 0               ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the CHOOSEFONT structure specified when the dialog was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnHook = 0                ' Pointer to a CFHookProc hook procedure that can process messages intended for the dialog box. This member is ignored unless the CF_ENABLEHOOK flag is set in the Flags member.
    .lpTemplateName = ""         ' Pointer to a null-terminated string that names the dialog box template resource in the module identified by the hInstance member. This template is substituted for the standard dialog box template. For numbered dialog box resources, lpTemplateName can be a value returned by the MAKEINTRESOURCE macro. This member is ignored unless the CF_ENABLETEMPLATE flag is set in the Flags member.
  End With
  
  ' Display the dialog
  ReturnCode = DLG_ChooseFont(FontInfo)
  
  ' Check if dialog canceled or error occured
  If GetLastError_CDLG(True, "DLG_ChooseFont") = True Or ReturnCode = 0 Then
    Flags = 0
    fontName = ""
    FontSize = 0
    FontBold = False
    FontItalic = False
    FontStrikethru = False
    FontUnderline = False
    Color = 0
    CD_ShowFont = False
    
  ' Return the information
  Else
    ' Copy the LOGFONT information back to the original variable
    CopyMemory TheLogFont, ByVal pLogFont, Len(TheLogFont)
    With TheLogFont
      If .lfWeight > FW_NORMAL Then
        FontBold = True
      Else
        FontBold = False
      End If
      Flags = FontInfo.Flags
      fontName = Left(.lfFaceName, InStr(.lfFaceName, Chr(0)) - 1)
      FontSize = FontInfo.iPointSize / 10
      FontItalic = .lfItalic
      FontStrikethru = .lfStrikeOut
      FontUnderline = .lfUnderline
      Color = FontInfo.RGBColors
    End With
    CD_ShowFont = True
  End If
  
FreeMemory:
  
  GlobalUnlock hLogFont
  GlobalFree hLogFont
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowFont function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowFont = False
    Resume FreeMemory
  End If
  
End Function

'=============================================================================================================
' CD_ShowFormat
'
' Purpose :
' Shows the standard Windows "Format Drive" dialog with the specified drive
' index.  Drive 0 = A:\, Drive 1 = B:\, Drive 2 = C:\, etc.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' DriveIndex           Optional. Sets which drive index to format
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowFormat(ByVal OwnerHandle As Long, _
                              Optional ByVal DriveIndex As Long) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnValue As Long
  
  ' Display the dialog
  ReturnValue = DLG_FormatDrive(OwnerHandle, DriveIndex, 0, 0)
  
  ' Return Values = Cancel : -2 /  OK : 6
  
  ' Check if dialog canceled or error occured
  If ReturnValue <= 0 Then
    CD_ShowFormat = False
  Else
    CD_ShowFormat = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowFormat function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowFormat = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowHelp
'
' Purpose :
' Displays the selected help file.  This function displays it in different
' ways, or only displays part of it depending on the flag passed to this
' function.
'
' Compare to the COMDLG32.OCX ShowHelp method.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' HelpFilePath         Sets the help file (.HLP) to display
' HelpCommand          Optional. Sets the help flag to use.  This determines what to do with the help file
' ContextID            Optional. Sets the context ID to use.  This is only used for certain flags.
' MacroName_or_Key     Optional. Sets the Macro to use, or Key to search for.
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowHelp(ByVal OwnerHandle As Long, _
                            ByVal HelpFilePath As String, _
                            Optional ByVal HelpCommand As HelpCommands = HELP_TAB, _
                            Optional ByVal ContextID As Long = 0, _
                            Optional ByVal MacroName_or_Key As String = "") As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  
  ' Make sure the help file exists and is valid
  If HelpCommand <> HELP_HELPONHELP And HelpCommand <> HELP_QUIT Then
    If HelpFilePath = "" Then
      Exit Function
    ElseIf Dir(HelpFilePath) = "" Then
      MsgBox HelpFilePath & Chr(13) & Chr(13) & "Could not locate this file to open.  Make sure this file has not been renamed, moved, or deleted.", vbOKOnly + vbExclamation, "  File Not Found"
      Exit Function
    ElseIf UCase(Right(HelpFilePath, 3)) <> "HLP" Then
      MsgBox HelpFilePath & Chr(13) & Chr(13) & "Can not open this file because this file does not appear to be a WinHelp (.HLP) help file.", vbOKOnly + vbExclamation, "  Invalid WinHelp File"
      Exit Function
    End If
  End If
  
  ' Display the help file according to the command
  Select Case HelpCommand
    Case HELP_COMMAND
      ReturnCode = DLG_WinHelp_STR(OwnerHandle, HelpFilePath, HELP_COMMAND, MacroName_or_Key)
    Case HELP_CONTEXT
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, HelpFilePath, HELP_CONTEXT, ContextID)
    Case HELP_CONTEXTPOPUP
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, HelpFilePath, HELP_CONTEXTPOPUP, ContextID)
    Case HELP_FORCEFILE
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, HelpFilePath, HELP_FORCEFILE, 0)
    Case HELP_HELPONHELP
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, vbNullString, HELP_HELPONHELP, 0)
    Case HELP_KEY
      ReturnCode = DLG_WinHelp_STR(OwnerHandle, HelpFilePath, HELP_KEY, MacroName_or_Key)
    Case HELP_QUIT
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, vbNullString, HELP_QUIT, 0)
    Case HELP_TAB
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, HelpFilePath, HELP_TAB, 0)
    Case Else
      ReturnCode = DLG_WinHelp_LNG(OwnerHandle, HelpFilePath, HELP_TAB, 0)
  End Select
  
  ' Return results
  If ReturnCode = 0 Then
    CD_ShowHelp = False
  Else
    CD_ShowHelp = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowHelp function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowHelp = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowIcon
'
' Purpose :
' Displays the standard Windows "Select Icon" dialog.
'
' NOTE : Use the ExtractIconEx API to extract the selected icon index
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' FileName             Optional. Sets / retrieves the selected icon resource file
' IconIndex            Optional. Sets / retrieves the selected icon index
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : FileName  = The selected file path
'                    IconIndex = The selected icon index
'
'=============================================================================================================
Public Function CD_ShowIcon(ByVal OwnerHandle As Long, _
                            Optional ByRef FileName As String, _
                            Optional ByRef IconIndex As Long) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  
  ' Properly format the string
  FileName = FileName & String(MAX_PATH - Len(FileName), Chr(0))
  FileName = Left(FileName, MAX_PATH)
  
  ' Display the dialog
  ReturnCode = DLG_GetIcon(OwnerHandle, FileName, 0, IconIndex)
  
  ' Check if dialog canceled or error occured
  If ReturnCode = 0 Then
    FileName = ""
    IconIndex = 0
    CD_ShowIcon = False
    
  ' Return the values
  Else
    FileName = Left(FileName, InStr(FileName, Chr(0)) - 1)
    CD_ShowIcon = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowIcon function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowIcon = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowOpen_Save
'
' Purpose :
' Displays the standard Windows "Open File" or "Save File" dialog depending
' on which the user wants to display.
'
' Compare to COMDLG32.OCX ShowOpen & ShowSave methods.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( OFN_... )
' FileName             Optional. Sets / recieves the full file path.
'                      If the OFN_ALLOWMULTISELECT flag is used, this returns the
'                      path where the files where selected from, followed by a
'                      NULL (vbNullChar), followed by a NULL-seperated list of
'                      the files that were selected, followed by a DOUBLE-NULL
'                      terminator at the end.
' FileTitle            Optional. Sets / recieves the file name (w/o path)
'                      If the OFN_ALLOWMULTISELECT flag is used, this returns a
'                      blank string.
' DefaultExt           Optional. Sets the default file extension
' DialogTitle          Optional. Sets the titlebar caption of the dialog
' Filter               Optional. Sets / recieves the browse file filter
' InitDir              Optional. Sets the initial browse directory
' Open_Not_Save        Optional. TRUE  = ShowOpen
'                                FALSE = ShowSave
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags     = Dialog return flags
'                    FileName  = Full file path for selected file
'                    FileTitle = File name selected (w/o path)
'                    Fileter   = Filter used to select the file
'
'=============================================================================================================
Public Function CD_ShowOpen_Save(ByVal OwnerHandle As Long, _
                                 Optional ByRef Flags As Long, _
                                 Optional ByRef FileName As String, _
                                 Optional ByRef FileTitle As String, _
                                 Optional ByVal DefaultExt As String, _
                                 Optional ByVal DialogTitle As String, _
                                 Optional ByRef Filter As String = "All Files (*.*)|*.*", _
                                 Optional ByVal InitDir As String = "C:\", _
                                 Optional ByVal Open_Not_Save As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  Dim FileInfo As OPENFILENAME
  Dim ReturnCode As Long
  Dim ReturnFilePath As String
  Dim ReturnFileName As String
  Dim SelectedFilter As String
  Dim FilterIndex As Long
  Dim FileNameOffset As Long
  Dim ExtensionOffset As Long
  
  ' Initialize the buffers to recieve the paths
  ReturnFilePath = FileName & String(MAX_PATH - Len(FileName), Chr(0))
  ReturnFileName = String(MAX_PATH, Chr(0))
  
  ' Initialize the variables that will be used to return the selected filter
  SelectedFilter = String(MAX_PATH, Chr(0))
  FilterIndex = 1
  
  ' Setup the initial directory
  If InitDir = "" Then InitDir = CurDir
  
  ' Make sure DefaultExt is correct
  If DefaultExt = "" Then DefaultExt = Chr(0)
  
  ' Make sure strings are NULL terminated
  If Right(DefaultExt, 1) <> Chr(0) Then DefaultExt = DefaultExt & Chr(0)
  If Right(InitDir, 1) <> Chr(0) Then InitDir = InitDir & Chr(0)
  If Right(DialogTitle, 1) <> Chr(0) Then DialogTitle = DialogTitle & Chr(0)
  
  ' Setup the information to use
  With FileInfo
    .lStructSize = Len(FileInfo)          ' Specifies the length of this type/structure
    .hwndOwner = OwnerHandle              ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .lpstrFilter = StripFilter(Filter)    ' Sets the specified filter correctly via the StripFilter function
    .lpstrCustomFilter = SelectedFilter   ' > Sets / recieves the selected filter. If this = NULL then no filter is returned
    .nFilterIndex = FilterIndex           ' Sets the default filter index to use.  1,2,3, etc. for standard filters... 0 if .lpstrCustomFilter is used to return the filter selected
    .lpstrFile = ReturnFilePath           ' > Sets / recieves the full path of the selected file(s)
    .lpstrFileTitle = ReturnFileName      ' > Sets / recieves the file name & extention of the selected file (this can be NULL)
    .lpstrInitialDir = InitDir            ' Sets the initial directory to browse from (if this is NULL, the current directory is used)
    .lpstrTitle = DialogTitle             ' Sets the caption of the dialog's titlebar (if this is NULL, the default "Save As" / "Open" is used)
    .Flags = Flags                        ' > Sets / receives the flags for the dialog
    .lpstrDefExt = DefaultExt             ' Sets the extension to attach to a file name if none is specified by the user.  If this is NULL, no extension is attached to the filename if the user doesn't specify one
    .nFileOffset = FileNameOffset         ' > Sets / recieves the number of characters to the left the first letter of the file name is located at (if lpstrFile = "c:\dir1\dir2\file.ext", nFileOffset = 13)
    .nFileExtension = ExtensionOffset     ' > Sets / recieves the number of characters to the left the first letter of the file extention is located at (if lpstrFile = "c:\dir1\dir2\file.ext", nFileExtension = 18)
    .nMaxFile = Len(ReturnFileName)       ' Specifies the length of .lpstrFile
    .nMaxFileTitle = Len(ReturnFileName)  ' Specifies the length of .lpstrFileTitle
    .nMaxCustFilter = Len(SelectedFilter) ' Specifies the length of the returned Filter from .lpstrCustomFilter (Ignored if .lpstrCustomFilter is NULL or vbNullString)
'------
    .hInstance = 0                        ' This is only needed if Flags contains OFN_ENABLETEMPLATEHANDLE, OFN_ENABLETEMPLATE, or OFN_EXPLORER
    .lCustData = 0                        ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the OPENFILENAME structure specified when the dialog box was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnHook = 0                         ' Pointer to a hook procedure. This member is ignored unless the Flags member includes the OFN_ENABLEHOOK flag.
                                          ' If the OFN_EXPLORER flag is not set in the Flags member, lpfnHook is a pointer to an OFNHookProcOldStyle hook procedure that receives messages intended for the dialog box. The hook procedure returns FALSE to pass a message to the default dialog box procedure or TRUE to discard the message.
                                          ' If OFN_EXPLORER is set, lpfnHook is a pointer to an OFNHookProc hook procedure. The hook procedure receives notification messages sent from the dialog box. The hook procedure also receives messages for any additional controls that you defined by specifying a child dialog template. The hook procedure does not receive messages intended for the standard controls of the default dialog box.
    .lpTemplateName = vbNullString        ' Pointer to a null-terminated string that names a dialog template resource in the module identified by the hInstance member. This member is ignored unless the OFN_ENABLETEMPLATE flag is set in the Flags member.
                                          ' If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
  End With
  
  ' Show open or save dialog depending on the use specification
  If Open_Not_Save = True Then
    ReturnCode = DLG_GetOpenFileName(FileInfo)
  Else
    ReturnCode = DLG_GetSaveFileName(FileInfo)
  End If
    
  ' Check if dialog canceled or error occured
  If GetLastError_CDLG(True, "DLG_GetOpenFileName") = True Or ReturnCode = 0 Then
    Flags = 0
    FileName = ""
    FileTitle = ""
    Filter = ""
    CD_ShowOpen_Save = False
    
  ' Return the information
  Else
    
    With FileInfo
      SelectedFilter = .lpstrCustomFilter
      Do While Left(SelectedFilter, 1) = Chr(0)
        SelectedFilter = Right(SelectedFilter, Len(SelectedFilter) - 1)
      Loop
      If SelectedFilter <> "" Then
         SelectedFilter = Left(SelectedFilter, InStr(SelectedFilter, Chr(0)) - 1)
      End If
      ReturnFilePath = .lpstrFile
      ReturnFileName = .lpstrFileTitle
    End With
    If (Flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
      Do While Right(ReturnFilePath, 1) = Chr(0)
        ReturnFilePath = Left(ReturnFilePath, Len(ReturnFilePath) - 1)
      Loop
      FileName = ReturnFilePath & Chr(0) & Chr(0)
    Else
      FileName = Left(ReturnFilePath, InStr(ReturnFilePath, Chr(0)) - 1)
    End If
    FileTitle = Left(ReturnFileName, InStr(ReturnFileName, Chr(0)) - 1)
    Filter = SelectedFilter
    Flags = FileInfo.Flags
    CD_ShowOpen_Save = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowOpen_Save function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowOpen_Save = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowPageSetup
'
' Purpose :
' Displays the standard "Page Setup" dialog used in word processors like MS Word.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( PSD_... )
' Orientation          Optional. Sets / recieves the selected page orientation ( DMORIENT_...)
' PaperSelection       Optional. Sets / recieves the selected paper size ( DMPAPER_...)
' PaperSize_Height     Optional. Recieves the selected paper height (in pixels)
' PaperSize_Width      Optional. Recieves the selected paper width (in pixels)
' Margin_Left          Optional. Sets / recieves the selected LEFT margin
' Margin_Top           Optional. Sets / recieves the selected TOP margin
' Margin_Right         Optional. Sets / recieves the selected RIGHT margin
' Margin_Bottom        Optional. Sets / recieves the selected BOTTOM margin
' MinMargin_Left       Optional. Sets the minimum LEFT margin
' MinMargin_Top        Optional. Sets the minimum TOP margin
' MinMargin_Right      Optional. Sets the minimum RIGHT margin
' MinMargin_Bottom     Optional. Sets the minimum BOTTOM margin
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags            = Dialog return flags
'                    PaperSelection   = Selected paper size
'                    PaperSize_Height = Selected paper height (in pixels)
'                    PaperSize_Width  = Selected paper width (in pixels)
'                    Margin_Left      = Selected LEFT margin
'                    Margin_Top       = Selected TOP margin
'                    Margin_Right     = Selected RIGHT margin
'                    Margin_Bottom    = Selected BOTTOM margin
'
'=============================================================================================================
Public Function CD_ShowPageSetup(ByVal OwnerHandle As Long, _
                                 Optional ByRef Flags As Long, _
                                 Optional ByRef Orientation As Long, _
                                 Optional ByRef PaperSelection As Long, _
                                 Optional ByRef PaperSize_Height As Long, _
                                 Optional ByRef PaperSize_Width As Long, _
                                 Optional ByRef Margin_Left As Long, _
                                 Optional ByRef Margin_Top As Long, _
                                 Optional ByRef Margin_Right As Long, _
                                 Optional ByRef Margin_Bottom As Long, _
                                 Optional ByVal MinMargin_Left As Long, _
                                 Optional ByVal MinMargin_Top As Long, _
                                 Optional ByVal MinMargin_Right As Long, _
                                 Optional ByVal MinMargin_Bottom As Long) As Boolean
On Error GoTo ErrorTrap
  
  Dim PageInfo As PAGESETUPDLG
  Dim PaperSize As POINTAPI
  Dim Margin As RECT
  Dim MinMargin As RECT
  Dim ReturnCode As Long
  Dim DevInfo As DEVMODE
  Dim DevName As DEVNAMES
  Dim hDevMode As Long
  Dim pDevMode As Long
  Dim hDevNames As Long
  Dim pDevNames As Long
  
  ' Make sure that the margins and min-margins are enabled
  If (Flags And PSD_MINMARGINS) = 0 Then
    Flags = (Flags Or PSD_MINMARGINS)
  End If
  If (Flags And PSD_MARGINS) = 0 Then
    Flags = (Flags Or PSD_MARGINS)
  End If
  If (Flags And PSD_DISABLEPRINTER) = 0 Then
    Flags = (Flags Or PSD_DISABLEPRINTER)
  End If
  
  ' Setup the required types / structures
  PaperSize.X = PaperSize_Width
  PaperSize.Y = PaperSize_Height
  Margin.Left = Margin_Left
  Margin.Top = Margin_Top
  Margin.Right = Margin_Right
  Margin.Bottom = Margin_Bottom
  MinMargin.Left = MinMargin_Left
  MinMargin.Top = MinMargin_Top
  MinMargin.Right = MinMargin_Right
  MinMargin.Bottom = MinMargin_Bottom
  
  '----------------------------------------------------------------------------------
  ' * NOTES ABOUT .hDevMode:
  ' Identifies a movable global memory object that contains a DEVMODE structure.
  ' Before a call to the PrintDlg function, the structure members may contain data
  ' used to initialize the dialog controls. When PrintDlg returns, the structure
  ' members specify the state of the dialog box controls.
  '
  ' If you do not use the structure to initialize the dialog box controls, hDevMode
  ' may be NULL. In this case, PrintDlg allocates memory for the structure,
  ' initializes its members, and returns a handle that identifies it.
  '
  ' If the device driver for the specified printer does not support extended device
  ' modes, hDevMode is NULL when PrintDlg returns.
  '
  ' If the device name (specified by the dmDeviceName member of the DEVMODE
  ' structure) does not appear in the [devices] section of WIN.INI, PrintDlg returns
  ' an error.
  '----------------------------------------------------------------------------------
  
  'Set the current orientation and duplex setting
  On Error Resume Next
  With DevInfo
    .dmSize = Len(DevInfo)
   '.dmDeviceName = Printer.DeviceName ' NOT needed for this function
   '.dmDriverExtra = 0                 ' NOT needed for this function
    .dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PRINTQUALITY  ' Or DM_COLOR Or DM_DUPLEX Or DM_COLLATE  Or DM_SCALE
    .dmOrientation = Orientation       ' Printer.Orientation
    .dmPaperSize = PaperSelection      ' Printer.PaperSize
   '.dmScale = <This member not needed for standard printing... but may be needed for more advanced printing>
   '.dmCopies = Printer.Copies         ' NOT needed for this function
   '.dmDefaultSource = 0               ' NOT needed for this function
    .dmPrintQuality = Printer.PrintQuality
   '.dmColor = Printer.ColorMode       ' NOT needed for this function
   '.dmDuplex = Printer.Duplex         ' NOT needed for this function
   '.dmCollate = 0                     ' NOT needed for this function
  End With
  On Error GoTo ErrorTrap
  
  ' Allocate a memory buffer for the DEVMODE variable
  hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevInfo))
  pDevMode = GlobalLock(hDevMode)
  ' Copy the DEVMODE type/structure to the memory buffer
  If pDevMode > 0 Then
    CopyMemory ByVal pDevMode, DevInfo, Len(DevInfo)
    GlobalUnlock hDevMode
  End If
  
  '----------------------------------------------------------------------------------
  ' * NOTES ABOUT .hDevName:
  ' Identifies a movable global memory object that contains a DEVNAMES structure.
  ' This structure contains three strings that specify the driver name, the printer
  ' name, and the output port name. Before the call to PrintDlg, the structure
  ' members contain strings used to initialize dialog box controls. When PrintDlg
  ' returns, the structure members contain the strings typed by the user. The
  ' calling application uses these strings to create a device context or an
  ' information context.
  '
  ' If you do not use the structure to initialize the dialog box controls,
  ' hDevNames may be NULL. In this case, PrintDlg allocates memory for the structure,
  ' initializes its members (by using the printer name specified in the DEVMODE
  ' structure), and returns a handle that identifies it.
  '
  ' PrintDlg uses the first port name that appears in the [devices] section of
  ' WIN.INI when it initializes the members in the DEVNAMES structure. For example,
  ' the function uses "LPT1:" as the port name if the following string appears in
  ' the [devices] section : "PCL / HP LaserJet=HPPCL,LPT1:,LPT2:"
  '----------------------------------------------------------------------------------
  
  'Set the current driver, device, and port name strings
  With DevName
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0
  End With
  With Printer
    DevName.Extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
  End With
  
  ' Allocate a memory buffer for the DEVNAMES variable
  hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
  pDevNames = GlobalLock(hDevNames)
  ' Copy the DEVNAMES type/structure to the memory buffer
  If pDevNames > 0 Then
    CopyMemory ByVal pDevNames, DevName, Len(DevName)
    GlobalUnlock hDevNames
  End If
  
  ' Setup the PageInfo type/structure to pass the function
  With PageInfo
    .lStructSize = Len(PageInfo) ' Size of this type / structure
    .hwndOwner = OwnerHandle     ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .hDevMode = hDevMode         ' Handle to a global memory object that contains a DEVMODE structure. On input, if a handle is given, the values in the corresponding DEVMODE structure are used to initialize the controls in the dialog box. On output, the dialog box sets hDevMode to a global memory handle for a DEVMODE structure that contains values specifying the user’s selections. If the user’s selections are not available, the dialog box sets hDevMode to NULL
    .hDevNames = hDevNames       ' Handle to a global memory object that contains a DEVNAMES structure. This structure contains three strings that specify the driver name, the printer name, and the output port name. On input, if a handle is given, the strings in the corresponding DEVNAMES structure are used to initialize controls in the dialog box. On output, the dialog box sets hDevNames to a  global memory handle for a DEVNAMES structure that contains strings specifying the user’s selections. If the user’s selections are not available, the dialog box sets hDevNames to NULL
    .Flags = Flags               ' > Sets / recieves the flags for the dialog
    .ptPaperSize = PaperSize     ' > Sets / recieves the dimensions of the paper selected by the user. The PSD_INTHOUSANDTHSOFINCHES or PSD_INHUNDREDTHSOFMILLIMETERS flag indicates the units of measurement.
    .rtMargin = Margin           ' > Sets / receives the widths of the left, top, right, and bottom margins. If you set the PSD_MARGINS flag, rtMargin specifies the initial margin values. When PageSetupDlg returns, rtMargin contains the margin widths selected by the user. The PSD_INHUNDREDTHSOFMILLIMETERS or PSD_INTHOUSANDTHSOFINCHES flag indicates the units of measurement.
    .rtMinMargin = MinMargin     ' Sets the minimum allowable widths for the left, top, right, and bottom margins. The system ignores this member if the PSD_MINMARGINS flag is not set. These values must be less than or equal to the values specified in the rtMargin member. The PSD_INTHOUSANDTHSOFINCHES or PSD_INHUNDREDTHSOFMILLIMETERS flag indicates the units of measurement.
'------
    .hInstance = 0               ' If the PSD_ENABLEPAGESETUPTEMPLATE flag is set in the Flags member, hInstance is the handle of the application or module instance that contains the dialog box template named by the lpPageSetupTemplateName member.
    .lCustData = 0               ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnPageSetupHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the PAGESETUPDLG structure specified when the dialog was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnPageSetupHook = 0       ' Pointer to a PageSetupHook hook procedure that can process messages intended for the dialog box. This member is ignored unless the PSD_ENABLEPAGESETUPHOOK flag is set in the Flags member.
    .lpfnPagePaintHook = 0       ' Pointer to a PagePaintHook hook procedure that receives WM_PSD_* messages from the dialog box whenever the sample page is redrawn. By processing the messages, the hook procedure can customize the appearance of the sample page. This member is ignored unless the PSD_ENABLEPAGEPAINTHOOK flag is set in the Flags member.
    .lpPageSetupTemplateName = "" ' Pointer to a null-terminated string that names the dialog box template resource in the module identified by the hInstance member. This template is substituted for the standard dialog box template. For numbered dialog box resources, lpPageSetupTemplateName can be a value returned by the MAKEINTRESOURCE macro. This member is ignored unless the PSD_ENABLEPAGESETUPTEMPLATE flag is set in the Flags member.
    .hPageSetupTemplate = 0      ' If the PSD_ENABLEPAGESETUPTEMPLATEHANDLE flag is set in the Flags member, hPageSetupTemplate is the handle of a memory object containing a dialog box template.
  End With
  
  'Show the pagesetup dialog
  ReturnCode = DLG_PageSetupDialog(PageInfo)
  
  ' Check if dialog canceled or error occured
  If GetLastError_CDLG(True, "DLG_PageSetupDialog") = True Or ReturnCode = 0 Then
    Flags = 0
    Orientation = 0
    PaperSelection = 0
    PaperSize_Height = 0
    PaperSize_Width = 0
    Margin_Left = 0
    Margin_Top = 0
    Margin_Right = 0
    Margin_Bottom = 0
    CD_ShowPageSetup = False
    
  ' Return the information
  Else
    
    ' Get the DevName structure
    pDevNames = GlobalLock(PageInfo.hDevNames)
    CopyMemory DevName, ByVal pDevNames, 45
    
    ' Get the DevInfo structure
    pDevMode = GlobalLock(PageInfo.hDevMode)
    CopyMemory DevInfo, ByVal pDevMode, Len(DevInfo)
    
    Flags = PageInfo.Flags
    Orientation = DevInfo.dmOrientation
    PaperSelection = DevInfo.dmPaperSize
    PaperSize_Height = PageInfo.ptPaperSize.Y
    PaperSize_Width = PageInfo.ptPaperSize.X
    Margin_Left = PageInfo.rtMargin.Left
    Margin_Top = PageInfo.rtMargin.Top
    Margin_Right = PageInfo.rtMargin.Right
    Margin_Bottom = PageInfo.rtMargin.Bottom
    CD_ShowPageSetup = True
    
  End If
  
FreeMemory:
  
  GlobalUnlock PageInfo.hDevNames
  GlobalFree PageInfo.hDevNames
  GlobalUnlock hDevNames
  GlobalFree hDevNames
  GlobalUnlock PageInfo.hDevMode
  GlobalFree PageInfo.hDevMode
  GlobalUnlock hDevMode
  GlobalFree hDevMode
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowPageSetup function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowPageSetup = False
    Resume FreeMemory
  End If
  
End Function

'=============================================================================================================
' CD_ShowPrinter
'
' Purpose :
' Displays the standard "Print" dialog.
'
' Compare to COMDLG32.OCX ShowPrinter method.
'
' IMPORTANT NOTE :
' For the OwnerHandle, you must specify a valid form's handle or error #65535
' will occur.  Apparently the dialog requires a valid window handle, not just
' an App.hInstance handle (i.e. - Form1.hWnd)
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets / recieves flag(s) to pass ( PD_... )
' PrinterName          Optional. Recieves the selected printer to use
' FromPage             Optional. Sets / Recieves the starting page to print
' ToPage               Optional. Sets / Recieves the ending page to print
' Min                  Optional. Sets / recieves the minimum value for FromPage / ToPage
' Max                  Optional. Sets / recieves the maximum value for FromPage / ToPage
' Copies               Optional. Sets / recieves the number of copies to print
' Duplex               Optional. Sets / recieves the printer DUPLEX setting to use ( DMDUP_...)
' Orientation          Optional. Sets / recieves the selected page orientation ( DMORIENT_...)
' PaperSize            Optional. Sets / recieves the selected paper size  ( DMPAPER_...)
' PrintQuality         Optional. Sets / recieves the print quality to use ( DMRES_... )
' ColorMode            Optional. Sets / recieves the print color mode to use ( DMCOLOR_...)
' PaperBin             Optional. Recieves the paper bin to print from
' Collate              Optional. Sets / recieves the selected collate mode
' MakeChangesToPrinter Optional. Sets whether the selected printer settings should be
'                                automatically applied to the Printer object
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds : Flags        = Dialog return flags
'                    PrinterName  = Selected printer to use
'                    FromPage     = Selected print start page
'                    ToPage       = Selected print end page
'                    Min          = Minimum FromPage / ToPage value
'                    Max          = Maximum FromPage / ToPage value
'                    Copies       = Selected number of copies to print
'                    Duplex       = Selected duplex mode to use
'                    Orientation  = Selected paper orientation to use
'                    PaperSize    = Selected paper size to use
'                    PrintQuality = Selected print quality to use
'                    ColorMode    = Selected printer color mode to use
'                    PaperBin     = Selected printer bin to print from
'                    Collate      = Selected collate value
'
'=============================================================================================================
Public Function CD_ShowPrinter(ByVal OwnerHandle As Long, _
                               Optional ByRef Flags As Long, _
                               Optional ByRef PrinterName As String, _
                               Optional ByRef FromPage As Integer, _
                               Optional ByRef ToPage As Integer, _
                               Optional ByRef Min As Integer, _
                               Optional ByRef Max As Integer, _
                               Optional ByRef Copies As Integer, _
                               Optional ByRef Duplex As Integer, _
                               Optional ByRef Orientation As Integer, _
                               Optional ByRef PaperSize As Integer, _
                               Optional ByRef PrintQuality As Integer, _
                               Optional ByRef ColorMode As Integer, _
                               Optional ByRef PaperBin As Integer, _
                               Optional ByRef Collate As Boolean, _
                               Optional ByVal MakeChangesToPrinter As Boolean = True) As Boolean
On Error GoTo ErrorTrap
  
  Dim PrintInfo As PRINTDLG
  Dim DevInfo As DEVMODE
  Dim DevName As DEVNAMES
  Dim pDevMode As Long
  Dim hDevMode As Long
  Dim pDevNames As Long
  Dim hDevNames As Long
  Dim ReturnCode As Long
  Dim objPrinter As Printer
  Dim strPrinterName As String
  Dim TempCollate As Integer
  
  ' Make sure that the "ToPages" param isn't greater than the "Max" param
  If ToPage > Max Then ToPage = Max
  
  ' Make sure that the "FromPages" param isn't less than the "Min" param
  If FromPage < Min Then FromPage = Min
  
  ' Check if the user wants to specify to collate
  If Collate = True Then
    TempCollate = 1
    If (Flags And PD_COLLATE) = 0 Then
      Flags = (Flags Or PD_COLLATE)
    End If
  Else
    If (Flags And PD_COLLATE) <> 0 Then
      Flags = (Flags Xor PD_COLLATE)
    End If
    TempCollate = 0
  End If
  
  ' If the user specifies to use page numbers, make sure that the
  ' PD_PAGENUMS flag is passed or it won't work
  If FromPage <> 0 Or ToPage <> 1 Then
    If (Flags And PD_PAGENUMS) = 0 Then
      Flags = (Flags Or PD_PAGENUMS)
    End If
  End If
  
  '----------------------------------------------------------------------------------
  ' * NOTES ABOUT .hDevMode:
  ' Identifies a movable global memory object that contains a DEVMODE structure.
  ' Before a call to the PrintDlg function, the structure members may contain data
  ' used to initialize the dialog controls. When PrintDlg returns, the structure
  ' members specify the state of the dialog box controls.
  '
  ' If you do not use the structure to initialize the dialog box controls, hDevMode
  ' may be NULL. In this case, PrintDlg allocates memory for the structure,
  ' initializes its members, and returns a handle that identifies it.
  '
  ' If the device driver for the specified printer does not support extended device
  ' modes, hDevMode is NULL when PrintDlg returns.
  '
  ' If the device name (specified by the dmDeviceName member of the DEVMODE
  ' structure) does not appear in the [devices] section of WIN.INI, PrintDlg returns
  ' an error.
  '----------------------------------------------------------------------------------
  
  'Set the current orientation and duplex setting
  On Error Resume Next
  With DevInfo
    .dmSize = Len(DevInfo)
    .dmDeviceName = Printer.DeviceName
    .dmDriverExtra = 0
    .dmFields = DM_ORIENTATION Or DM_PAPERSIZE Or DM_PAPERWIDTH Or DM_COPIES Or DM_PRINTQUALITY Or DM_COLOR Or DM_DUPLEX Or DM_COLLATE  'Or DM_SCALE
    .dmOrientation = Orientation     ' Printer.Orientation
    .dmPaperSize = PaperSize         ' Printer.PaperSize
   '.dmScale = <This member not needed for standard printing... but may be needed for more advanced printing>
    .dmCopies = Copies               ' Printer.Copies
    .dmDefaultSource = 0
    .dmPrintQuality = PrintQuality   ' Printer.PrintQuality
    .dmColor = ColorMode             ' Printer.ColorMode
    .dmDuplex = Duplex               ' Printer.Duplex
    .dmCollate = TempCollate         ' (1 = True, 0 = False)
  End With
  On Error GoTo ErrorTrap
  
  'Allocate memory for the initialization hDevMode structure and copy the settings gathered above into this memory
  hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevInfo))
  pDevMode = GlobalLock(hDevMode)
  If pDevMode > 0 Then
    CopyMemory ByVal pDevMode, DevInfo, Len(DevInfo)
    GlobalUnlock hDevMode
  End If
  
  '----------------------------------------------------------------------------------
  ' * NOTES ABOUT .hDevName:
  ' Identifies a movable global memory object that contains a DEVNAMES structure.
  ' This structure contains three strings that specify the driver name, the printer
  ' name, and the output port name. Before the call to PrintDlg, the structure
  ' members contain strings used to initialize dialog box controls. When PrintDlg
  ' returns, the structure members contain the strings typed by the user. The
  ' calling application uses these strings to create a device context or an
  ' information context.
  '
  ' If you do not use the structure to initialize the dialog box controls,
  ' hDevNames may be NULL. In this case, PrintDlg allocates memory for the structure,
  ' initializes its members (by using the printer name specified in the DEVMODE
  ' structure), and returns a handle that identifies it.
  '
  ' PrintDlg uses the first port name that appears in the [devices] section of
  ' WIN.INI when it initializes the members in the DEVNAMES structure. For example,
  ' the function uses "LPT1:" as the port name if the following string appears in
  ' the [devices] section : "PCL / HP LaserJet=HPPCL,LPT1:,LPT2:"
  '----------------------------------------------------------------------------------
  
  'Set the current driver, device, and port name strings
  With DevName
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0
  End With
  With Printer
    DevName.Extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
  End With
  
  'Allocate memory for the initial hDevName structure and copy the settings gathered above into this memory
  hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
  pDevNames = GlobalLock(hDevNames)
  If pDevNames > 0 Then
    CopyMemory ByVal pDevNames, DevName, Len(DevName)
    GlobalUnlock hDevNames
  End If
  
  ' Use DLG_PrintDialog to get the handle to a memory block with a DevInfo and DevName structures
  With PrintInfo
    .lStructSize = Len(PrintInfo)  ' Size of this type / structure
    .hwndOwner = OwnerHandle       ' Handle of the owner window for the dialog box. If this = NULL then the dialog has no owner
    .hDevMode = hDevMode           ' Identifies a movable global memory object that contains a DEVMODE structure. Before a call to the PrintDlg function, the structure members may contain data used to initialize the dialog controls. When PrintDlg returns, the structure members specify the state of the dialog box controls.
    .hDevNames = hDevNames         ' Identifies a movable global memory object that contains a DEVNAMES structure. This structure contains three strings that specify the driver name, the printer name, and the output port name. Before the call to PrintDlg, the structure members contain strings used to initialize dialog box controls. When PrintDlg returns, the structure members contain the strings typed by the user. The calling application uses these strings to create a device context or an information context.
    .hDC = Printer.hDC             ' Identifies a device context or an information context, depending on whether the Flags member specifies the PD_RETURNDC or PC_RETURNIC flag. If neither flag is specified, the value of this member is NULL. If both flags are specified, PD_RETURNDC has priority.
    .nMinPage = Min                ' > Sets / recieves the minimum value for the range of pages specified in the From and To page edit controls.
    .nMaxPage = Max                ' > Sets / recieves the maximum value for the range of pages specified in the From and To page edit controls.
    .nFromPage = FromPage          ' > Sets / recieves the starting page. This value is valid only if the PD_PAGENUMS flag is specified.
    .nToPage = ToPage              ' > Sets / recieves the ending page. This value is valid only if the PD_PAGENUMS flag is specified.
    .Flags = Flags                 ' > Sets / recieves the flags for the dialog
    .nCopies = Copies              ' > Sets / recieves the number of copies to print.
                                   ' Contains the initial number of copies for the Copies edit control if hDevMode is NULL; otherwise, the dmCopies member of the DEVMODE structure contains the initial value. When PrintDlg returns, this member contains the actual number of copies to print.
                                   ' If the printer driver does not support multiple copies, this value may be greater than one and the application must print all requested copies. If the PD_USEDEVMODECOPIESANDCOLLATE value is set in the Flags member, nCopies is always set to 1 on return and the dmCopies member in the DEVMODE receives the actual number of copies to print.
'------
    .hInstance = 0                 ' If the PD_ENABLEPRINTTEMPLATE or PD_ENABLESETUPTEMPLATE flag is set in the Flags member, hInstance is the handle of the application or module instance that contains the dialog box template named by the lpPrintTemplateName or lpSetupTemplateName member.
    .lCustData = 0                 ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnPrintHook or lpfnSetupHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message’s lParam parameter is a pointer to the PRINTDLG structure specified when the dialog was created. The hook procedure can use this pointer to get the lCustData value.
    .lpfnPrintHook = 0             ' Pointer to a PrintHookProc hook procedure that can process messages intended for the Print dialog box. This member is ignored unless the PD_ENABLEPRINTHOOK flag is set in the Flags member.
    .lpfnSetupHook = 0             ' Pointer to a SetupHookProc hook procedure that can process messages intended for the Print Setup dialog box. This member is ignored unless the PD_ENABLESETUPHOOK flag is set in the Flags member.
    .lpPrintTemplateName = ""      ' Pointer to a null-terminated string that names a dialog box template resource in the module identified by the hInstance member. This template is substituted for the standard Print dialog box template. This member is ignored unless the PD_ENABLEPRINTTEMPLATE flag is set in the Flags member.
    .lpSetupTemplateName = ""      ' Pointer to a null-terminated string that names a dialog box template resource in the module identified by the hInstance member. This template is substituted for the standard Print Setup dialog box template. This member is ignored unless the PD_ENABLESETUPTEMPLATE flag is set in the Flags member.
    .hPrintTemplate = 0            ' If the PD_ENABLEPRINTTEMPLATEHANDLE flag is set in the Flags member, hPrintTemplate is the handle of a memory object containing a dialog box template. This template is substituted for the standard Print dialog box template.
    .hSetupTemplate = 0            ' If the PD_ENABLESETUPTEMPLATEHANDLE flag is set in the Flags member, hSetupTemplate is the handle of a memory object containing a dialog box template. This template is substituted for the standard Print Setup dialog box template.
  End With
  
  ' Display the dialog
  ReturnCode = DLG_PrintDialog(PrintInfo)
  
  ' Check if dialog canceled or error occured
  If GetLastError_CDLG(True, "DLG_PrintDialog") = True Or ReturnCode = 0 Then
    Flags = 0
    PrinterName = ""
    FromPage = 0
    ToPage = 0
    Min = 0
    Max = 0
    Copies = 0
    Duplex = 0
    Orientation = 0
    PaperSize = 0
    PrintQuality = 0
    ColorMode = 0
    PaperBin = 0
    Collate = False
    CD_ShowPrinter = False
    
  ' Return the values
  Else
    On Error Resume Next
    
    ' Get the DevName structure
    pDevNames = GlobalLock(PrintInfo.hDevNames)
    CopyMemory DevName, ByVal pDevNames, 45
    
    ' Get the DevInfo structure
    pDevMode = GlobalLock(PrintInfo.hDevMode)
    CopyMemory DevInfo, ByVal pDevMode, Len(DevInfo)
    
    ' If the user specifies to change the printer info, change it
    strPrinterName = UCase(Left(DevInfo.dmDeviceName, InStr(DevInfo.dmDeviceName, Chr(0)) - 1))
    If MakeChangesToPrinter = True Then
      ' Check if the user selected a printer other than the default
      If UCase(Printer.DeviceName) <> UCase(strPrinterName) Then
        ' Iterate through the Printers to find the one they selected
        For Each objPrinter In Printers
          If UCase(objPrinter.DeviceName) = strPrinterName Then
            ' Set the printer to the selected printer
            Set Printer = objPrinter
          End If
        Next
      End If
      
      'Set printer object properties according to selections made by user
      Printer.Copies = DevInfo.dmCopies
      Printer.Duplex = DevInfo.dmDuplex
      Printer.Orientation = DevInfo.dmOrientation
      Printer.PaperSize = DevInfo.dmPaperSize
      Printer.PrintQuality = DevInfo.dmPrintQuality
      Printer.ColorMode = DevInfo.dmColor
      Printer.PaperBin = DevInfo.dmDefaultSource
    End If
    
    ' Set the return values
    Flags = PrintInfo.Flags
    PrinterName = strPrinterName
    FromPage = PrintInfo.nFromPage
    ToPage = PrintInfo.nToPage
    Min = PrintInfo.nMinPage
    Max = PrintInfo.nMaxPage
    Copies = DevInfo.dmCopies
    Duplex = DevInfo.dmDuplex
    Orientation = DevInfo.dmOrientation
    PaperSize = DevInfo.dmPaperSize
    PrintQuality = DevInfo.dmPrintQuality
    ColorMode = DevInfo.dmColor
    PaperBin = DevInfo.dmDefaultSource
    Collate = DevInfo.dmCollate
    CD_ShowPrinter = True
    
  End If
  
FreeMemory:
  
  GlobalUnlock PrintInfo.hDevNames
  GlobalFree PrintInfo.hDevNames
  GlobalUnlock hDevNames
  GlobalFree hDevNames
  GlobalUnlock PrintInfo.hDevMode
  GlobalFree PrintInfo.hDevMode
  GlobalUnlock hDevMode
  GlobalFree hDevMode
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowPrinter function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowPrinter = False
    Resume FreeMemory
  End If
  
End Function

'=============================================================================================================
' CD_ShowProperties
'
' Purpose :
' Displays the standard "Properties" dialog for the selected file / drive.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' FileName             Optional. Sets the file to display properties of
'                      If this is not specified, PrinterName needs to be
' PrinterName          Optional. Sets the printer to display properties of
'                      If this is not specified, FileName needs to be
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowProperties(ByVal OwnerHandle As Long, _
                                  Optional ByVal FileName As String, _
                                  Optional ByVal PrinterName As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  
  ' Check if params are valid
  If FileName = "" And PrinterName = "" Then
    MsgBox "No path specified to display properties.", vbOKOnly + vbExclamation, "  Missing Parameter"
    Exit Function
  End If
  
  ' Display the dialog
  If FileName <> "" Then
    ReturnCode = DLG_Properties(OwnerHandle, OPF_PATHNAME, FileName, "")
  Else
    ReturnCode = DLG_Properties(OwnerHandle, OPF_PRINTERNAME, PrinterName, "")
  End If
  
  ' Check if error occured
  If ReturnCode = 0 Then
    CD_ShowProperties = False
  Else
    CD_ShowProperties = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowProperties function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowProperties = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowReboot
'
' Purpose :
' Displays the standard "Your computer needs to reboot" message, then restarts
' the computer if the user selects YES.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets flag(s) to pass ( EWX_... )
' Prompt               Optional. Sets the prompt to be displayed in the dialog
'
' Return
' ------
' FALSE if user cancels or error occurs
' TRUE if succeeds and user selects YES
'
'=============================================================================================================
Public Function CD_ShowReboot(ByVal OwnerHandle As Long, _
                              Optional ByVal Flags As Long, _
                              Optional ByVal Prompt As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnCode As Long
  
  ' Make sure there's a line return at end
  If Right(Prompt, 1) <> Chr(13) And Right(Prompt, 1) <> Chr(10) Then
    Prompt = Prompt & Chr(13)
  End If
  
  ' Display the dialog
  ReturnCode = DLG_Reboot(OwnerHandle, Prompt, Flags)
  
  ' Check for errors
  If ReturnCode = 0 Then
    CD_ShowReboot = False
  ElseIf ReturnCode = IDNO Then
    CD_ShowReboot = False
  ElseIf ReturnCode = IDYES Then
    CD_ShowReboot = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowReboot function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowReboot = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowRun
'
' Purpose :
' Shows the standard Windows "Run Program" dialog.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
' Flags                Optional. Sets flag(s) to pass ( RFF_... )
' Prompt               Optional. Sets message to display on the dialog
' Title                Optional. Sets the titlebar caption of the dialog
' hIcon                Optional. Handle to the icon to display in the dialog
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowRun(ByVal OwnerHandle As Long, _
                           Optional ByVal Flags As Long, _
                           Optional ByVal Prompt As String, _
                           Optional ByVal Title As String, _
                           Optional ByVal hIcon As Long) As Boolean
On Error GoTo ErrorTrap
  
  ' Put in default values
  If Prompt = "" Then
    Prompt = "Type the name of a program, folder, document, or internet resource to open:"
  End If
  If Title = "" Then
    Title = "Run"
  End If
  
  ' Display the dialog
  DLG_Run OwnerHandle, hIcon, "", Title, Prompt, Flags
  CD_ShowRun = True
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowRun function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowRun = False
  End If
  
End Function

'=============================================================================================================
' CD_ShowShutDown
'
' Purpose :
' Shows the standard Windows "Shut Down Computer" dialog.
'
' Param                Use
' ------------------------------------
' OwnerHandle          Handle to the owner of the dialog
'
' Return
' ------
' FALSE if error occurs
' TRUE if succeeds
'
'=============================================================================================================
Public Function CD_ShowShutDown(ByVal OwnerHandle As Long) As Boolean
On Error GoTo ErrorTrap
  
  Dim ReturnValue As Long
  
  ' Display the dialog
  ReturnValue = DLG_ShutDown(OwnerHandle)
  
  ' Check if error occured
  If ReturnValue = 0 Then
    CD_ShowShutDown = False
  Else
    CD_ShowShutDown = True
  End If
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " encountered the following error in the CD_ShowShutDown function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
    CD_ShowShutDown = False
  End If
  
End Function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' Function that converts automation colors such as "vbButtonFace" to standard
' color such as "12632256".  It is safest to pass all colors through this
' function to make sure that if a user passes a color like "Me.BackColor" and
' the BackColor is vbButtonFace, it won't mess up any of the API's that are
' expecting a normal color value.
Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
On Error Resume Next
  
  If OleTranslateColor(oClr, hPal, TranslateColor) Then
    TranslateColor = CLR_INVALID
  End If
  
End Function

' Function that takes the ANSI text string passed and converts it to
' UNICODE text if the user's operating system is Windows NT 4 or 5
Private Function Convert_STR_UNI(ByVal ANSI_String As String) As String
On Error Resume Next
  
  ' String parameters for undocumented functions need to be converted
  ' to UNICODE for WinNT systems
  If CheckTheOS = OS_WinNT_40 Or CheckTheOS = OS_Win2000 Then
    Convert_STR_UNI = StrConv(ANSI_String, vbUnicode)
  Else
    Convert_STR_UNI = ANSI_String
  End If
  
End Function

' Function that takes the UNICODE text string passed and converts it to
' ANSI text if the user's operating system is Windows NT 4 or 5
Private Function Convert_UNI_STR(ByVal UNICODE_String As String) As String
On Error Resume Next
  
  ' String parameters for undocumented functions need to be converted
  ' to UNICODE for WinNT systems
  If CheckTheOS = OS_WinNT_40 Or CheckTheOS = OS_Win2000 Then
    Convert_UNI_STR = StrConv(UNICODE_String, vbFromUnicode)
  Else
    Convert_UNI_STR = UNICODE_String
  End If
  
End Function

' Function that takes a pointer and changes it to the string value at that memory location of the pointer
Private Function Convert_PTR_STR(ByVal ThePointer As Long) As String
On Error Resume Next
  
  Dim TheString As String
  
  TheString = String(MAX_PATH, Chr(0))
  CopyPointer2String TheString, ThePointer
  Convert_PTR_STR = Left(TheString, InStr(TheString, Chr(0)) - 1)
  
End Function

' Function that takes a string and returns a pointer to where that string is located in memory
Private Function Convert_STR_PTR(ByRef TheString As Long) As Long
On Error Resume Next
  
  Convert_STR_PTR = StrPtr(TheString)
  
End Function

' Function that takes a standard color that is stored in a LONG variable type
' and breaks it out into it's respective RED, GREEN, & BLUE parts and returns
' these three values as BYTEs (0 to 255).
Private Function Convert_LNG_RGB(ByVal lngColor As Long, ByRef Return_Red As Byte, ByRef Return_Green As Byte, ByRef Return_Blue As Byte) As Boolean
On Error GoTo ErrorTrap
  
  Return_Blue = (lngColor And &HFF0000) / 65536
  Return_Green = (lngColor And &HFF00) / 256 Mod 256
  Return_Red = lngColor And &HFF
  
  Convert_LNG_RGB = True
  
  Exit Function
  
ErrorTrap:

  Err.Clear
  Return_Blue = 0
  Return_Green = 0
  Return_Red = 0
  
End Function

' Function that takes a the Red, Green, & Blue parts of a color as BYTE (0 to 255)
' values and translates it into a long value to represent it.
Private Function Convert_RGB_LNG(ByVal TheRed As Byte, ByVal TheGreen As Byte, ByVal TheBlue As Byte, ByRef Return_Long As Long) As Boolean
On Error GoTo ErrorTrap
  
  Return_Long = RGB(TheRed, TheGreen, TheBlue)
  
  Convert_RGB_LNG = True
  
  Exit Function
  
ErrorTrap:

  Err.Clear
  TheBlue = 0
  TheGreen = 0
  TheRed = 0
  
End Function

' Function that checks what OS the user is running.  This function is designed
' to check the OS only once to save on excess CPU usage.
Private Function CheckTheOS() As OSTypes
On Error Resume Next
  
  ' Check the operating system
  ' (NOTE - This only checks the OS once, and if it fails doesn't try again)
  If Win_OS = OS_Unknown And CantGetOSInfo = False Then
    If GetOS = False Then
      Win_OS = OS_Unknown
      CantGetOSInfo = True
    Else
      CheckTheOS = Win_OS
    End If
  ElseIf Win_OS = OS_Unknown And CantGetOSInfo = True Then
    CheckTheOS = OS_Unknown
  ElseIf CantGetOSInfo = False Then
    CheckTheOS = Win_OS
  Else
    CheckTheOS = OS_Unknown
  End If
  
End Function

' Function to gets version information about the user's Windows OS
Private Function GetOS() As Boolean
On Error GoTo TheEnd
  
  Dim OSinfo As OSVERSIONINFO
  Dim RetValue As Long
  Dim PID As String
  
  OSinfo.dwOSVersionInfoSize = 148
  OSinfo.szCSDVersion = Space(128)
  RetValue = GetVersionEx(OSinfo)
  If RetValue = 0 Then
    Win_Build = ""
    Win_OS = OS_Unknown
    Win_Version = ""
    GetOS = False
    Exit Function
  End If

  With OSinfo
    Select Case .dwPlatformId
      Case VER_PLATFORM_WIN32s
        PID = "Win 32"
        Win_OS = OS_Win32
      Case VER_PLATFORM_WIN32_WINDOWS
        If .dwMinorVersion = 0 Then
          PID = "Windows 95"
          Win_OS = OS_Win95
        ElseIf .dwMinorVersion = 10 Then
          PID = "Windows 98"
          Win_OS = OS_Win98
        End If
      Case VER_PLATFORM_WIN32_NT
        If .dwMajorVersion = 3 Then
          PID = "Windows NT 3.51"
          Win_OS = OS_WinNT_351
        ElseIf .dwMajorVersion = 4 Then
          PID = "Windows NT 4.0"
          Win_OS = OS_WinNT_40
        ElseIf .dwMajorVersion = 5 Then
          PID = "Windows 2000"
          Win_OS = OS_Win2000
        End If
      Case Else
        PID = "Unknown"
        Win_OS = OS_Unknown
    End Select
  End With
  
  Win_Version = Trim(Str(OSinfo.dwMajorVersion) & "." & LTrim(Str(OSinfo.dwMinorVersion)))
  Win_Build = Trim(Str(OSinfo.dwBuildNumber))
  
  GetOS = True
  
  Exit Function
  
TheEnd:
  
  Err.Clear
  GetOS = False
  
End Function

' Checks to see if an error occured in one of the following functions, and if one did,
' displays it in a standard error message dialog:
'
'  - DLG_GetOpenFileName
'  - DLG_GetSaveFileName
'  - DLG_ChooseColor Lib
'  - DLG_ChooseFont Lib
'  - DLG_PrintDialog Lib
'  - DLG_PageSetupDialog
'
Private Function GetLastError_CDLG(Optional ByVal ShowError As Boolean = True, Optional ByVal LastFunctionCalled As String = "last") As Boolean
On Error Resume Next
  
  Dim ReturnCode As Long
  Dim TheMsg As String
  
  ReturnCode = DLG_GetLastError
  
  Select Case ReturnCode
    ' General Common Dialog Return Codes
    Case CDERR_NOERROR          ' 0
      GetLastError_CDLG = False
      Exit Function
    Case CDERR_DIALOGFAILURE, 65535 ' &HFFFF
      TheMsg = "The dialog box could not be created." & Chr(13) & "The common dialog box function’s call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle."
    Case CDERR_FINDRESFAILURE   ' &H6
      TheMsg = "The common dialog box function failed to find a specified resource."
    Case CDERR_INITIALIZATION   ' &H2
      TheMsg = "The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available."
    Case CDERR_LOADRESFAILURE   ' &H7
      TheMsg = "The common dialog box function failed to load a specified resource."
    Case CDERR_LOADSTRFAILURE   ' &H5
      TheMsg = "The common dialog box function failed to load a specified string."
    Case CDERR_LOCKRESFAILURE   ' &H8
      TheMsg = "The common dialog box function failed to lock a specified resource."
    Case CDERR_MEMALLOCFAILURE  ' &H9
      TheMsg = "The common dialog box function was unable to allocate memory for internal structures."
    Case CDERR_MEMLOCKFAILURE   ' &HA
      TheMsg = "The common dialog box function was unable to lock the memory associated with a handle."
    Case CDERR_NOHINSTANCE      ' &H4
      TheMsg = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but a corresponding instance handle was not provide."
    Case CDERR_NOHOOK           ' &HB
      TheMsg = "The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but a pointer to a corresponding hook procedure was not provided."
    Case CDERR_NOTEMPLATE       ' &H3
      TheMsg = "The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but a corresponding template was not provided."
    Case CDERR_REGISTERMSGFAIL  ' &HC
      TheMsg = "The RegisterWindowMessage function returned an error code when it was called by the common dialog box function."
    Case CDERR_STRUCTSIZE       ' &H1
      TheMsg = "The lStructSize member of the initialization structure for the corresponding common dialog box is invalid."
    
    ' DLG_PrintDialog Function Return Codes
    Case PDERR_CREATEICFAILURE  ' &H100A
      TheMsg = "The DLG_PrintDialog function failed when it attempted to create an information context."
    Case PDERR_DEFAULTDIFFERENT ' &H100C
      TheMsg = "You called the DLG_PrintDialog function with the DN_DEFAULTPRN flag specified in the wDefault member of the DEVNAMES structure, but the printer described by the other structure members did not match the current default printer." & Chr(13) & Chr(13) & "This error occurs when the DEVNAMES structure is stored and the user changes the default printer by using the Control Panel."
    Case PDERR_DNDMMISMATCH     ' &H1009
      TheMsg = "The data in the DEVMODE and DEVNAMES structures describes two different printers."
    Case PDERR_GETDEVMODEFAIL   ' &H1005
      TheMsg = "The printer driver failed to initialize a DEVMODE structure." & Chr(13) & "This error code applies only to printer drivers written for Windows versions 3.0 and later."
    Case PDERR_INITFAILURE      ' &H1006
      TheMsg = "The DLG_PrintDialog function failed during initialization, and there is no more specific extended error code to describe the failure."
    Case PDERR_LOADDRVFAILURE   ' &H1004
      TheMsg = "The DLG_PrintDialog function failed to load the device driver for the specified printer."
    Case PDERR_NODEFAULTPRN     ' &H1008
      TheMsg = "A default printer does not exist."
    Case PDERR_NODEVICES        ' &H1007
      TheMsg = "No printer drivers were found."
    Case PDERR_PARSEFAILURE     ' &H1002
      TheMsg = "The DLG_PrintDialog function failed to parse the strings in the [devices] section of the WIN.INI file."
    Case PDERR_PRINTERNOTFOUND  ' &H100B
      TheMsg = "The [devices] section of the WIN.INI file did not contain an entry for the requested printer."
    Case PDERR_RETDEFFAILURE    ' &H1003
      TheMsg = "The PD_RETURNDEFAULT flag was specified in the Flags member of the PRINTDLG structure, but the hDevMode or hDevNames member was not NULL."
    Case PDERR_SETUPFAILURE     ' &H1001
      TheMsg = "The DLG_PrintDialog function failed to load the required resources."
    
    ' DLG_ChooseFont Function Return Codes
    Case CFERR_MAXLESSTHANMIN   ' &H2002
      TheMsg = "The size specified in the nSizeMax member of the CHOOSEFONT structure is less than the size specified in the nSizeMin member."
    Case CFERR_NOFONTS          ' &H2001
      TheMsg = "No fonts exist."
    
    ' DLG_GetOpenFileName / DLG_GetSaveFileName Function Return Codes
    Case FNERR_BUFFERTOOSMALL   ' &H3003
      TheMsg = "The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the filename specified by the user."
    Case FNERR_INVALIDFILENAME  ' &H3002
      TheMsg = "A filename is invalid."
    Case FNERR_SUBCLASSFAILURE  ' &H3001
      TheMsg = "An attempt to subclass a list box failed because sufficient memory was not available."
    
    ' Other error
    Case Else
      TheMsg = "Unknown Error"
  End Select
  
  GetLastError_CDLG = True
  
  If ShowError = True Then
    MsgBox "COMDLG32.DLL encountered the following error while calling the '" & LastFunctionCalled & "' function:" & Chr(13) & Chr(13) & "Error Number = " & CStr(ReturnCode) & Chr(13) & "Error Description = " & TheMsg, vbOKOnly + vbExclamation, "  Common Dialog Error"
  End If
  
End Function

' Function that takes the standard "Pipe Seperated" filter statements for
' the ShowOpen / ShowSave functions and changes them to a format that the
' COMDLG32.DLL API functions expect ("NULL Seperated" + Double NULL End)
Private Function StripFilter(ByVal TheFilter As String) As String
On Error Resume Next
  
  Dim MyCounter As Integer
  Dim CharLeft As String
  Dim CharRight As String
  Dim StringSoFar As String
  
  ' If there is no filter specified, return NULL
  If TheFilter = "" Then
    StripFilter = Chr(0)
    Exit Function
  End If
  
  ' Parse the string looking for the VB standard PIPE "|" separator and
  ' replace it with a NULL if found
  For MyCounter = 1 To Len(TheFilter)
    CharLeft = Left(TheFilter, MyCounter)
    CharRight = Right(CharLeft, 1)
    If CharRight = "|" Then
      CharRight = Chr(0)
    End If
    StringSoFar = StringSoFar & CharRight
  Next
  
  ' The filter always has to be terminated with 2 NULLs
  StringSoFar = StringSoFar & Chr(0) & Chr(0)
  StripFilter = StringSoFar
  
End Function


' Function that checks to see if one specific flag is set in a series of flags
Private Function CheckFlags(ByVal FlagToCheck As Long, ByVal FlagsToSearch As Long) As Boolean
On Error Resume Next
  
  CheckFlags = ((FlagsToSearch And FlagToCheck) = FlagToCheck)
  
End Function

' Message loop that continually looks for messages being sent to either the Find/Replace
' dialog, or the dialog's owner form.  If the message is for the owner form, then process
' it because that's where the main messages for the Find/Replace dialog are sent to.
Private Sub MessageLoop()
On Error Resume Next
  
  ' Keep looking for messages while the Find/Replace dialog is still open
  Do While GetMessage(TheMessage, 0&, 0&, 0&) And hDialog > 0
    
    ' If the current message is a message from the Find/Replace Dialog's owner, process it here
    If IsDialogMessage(hDialog, TheMessage) = 0 Then
     'Debug.Print "Calling Form Message - hDialog = " & CStr(hDialog) & ", TheMessage = " & CStr(TheMessage.Message)
      TranslateMessage TheMessage
      DispatchMessage TheMessage
      
    ' If the current message is a message from the Find/Replace Dialog it self, process it here
    Else
     'Debug.Print "Dialog Message - hDialog = " & CStr(hDialog) & ", TheMessage = " & CStr(TheMessage.Message)
     
    End If
  Loop
  
End Sub

' Function that handles the subclassed Windows messages to the Find/Replace dialog
Public Function FindReplaceProc(ByVal hOwner As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
  
  Select Case wMsg
    
    ' Find button pressed
    Case FINDMESSAGE
      CopyMemory ReturnFR, ByVal lParam, Len(ReturnFR)
      If (ReturnFR.Flags And FR_DIALOGTERM) = FR_DIALOGTERM Then
        GoTo UnSubClass
      Else
        FindReplace_Event ReturnFR, False
      End If
      
    ' Help button pressed
    Case HELPMESSAGE
      FindReplace_Event ReturnFR, True
      
    ' Closing subclassed form... make sure it's unsubclassed first
    Case WM_CLOSE
      GoTo UnSubClass
      
    ' Other messages, re-route them to their original destination
    Case Else
      FindReplaceProc = CallWindowProc(OldProc, hOwner, wMsg, wParam, lParam)
      
  End Select
  
  Exit Function
  
UnSubClass:
  
  ' Unsubclass the dialog and release the memory associated with it
  SetWindowLong hOwner, GWL_WNDPROC, OldProc
  HeapFree GetProcessHeap(), 0, lHeap
  hDialog = 0
  
End Function

' Function that is called when the user clicks the
Private Function FindReplace_Event(ByRef FRInfo As FINDREPLACE, ByVal HelpButtonPressed As Boolean)
On Error Resume Next
  
  
'******************************************************************************************
'                PUT A CALL TO YOUR FIND/REPLACE PROCESSING FUNCTION HERE
'******************************************************************************************
  
  
  Dim sTemp As String
  
  If HelpButtonPressed = True Then
    MsgBox "Help Button Pressed", vbOKOnly + vbInformation, "Find/Replace Parameters"
    Exit Function
  End If
  
  With FRInfo
    sTemp = "Here is your code for Find/Replace with parameters:" & vbCrLf & vbCrLf
    sTemp = sTemp & "Find string: " & Convert_PTR_STR(.lpstrFindWhat) & vbCrLf
    sTemp = sTemp & "Replace string: " & Convert_PTR_STR(.lpstrReplaceWith) & vbCrLf & vbCrLf
    sTemp = sTemp & "Current Flags: " & vbCrLf & vbCrLf
    sTemp = sTemp & "FR_FINDNEXT = " & CheckFlags(FR_FINDNEXT, .Flags) & vbCrLf
    sTemp = sTemp & "FR_REPLACE = " & CheckFlags(FR_REPLACE, .Flags) & vbCrLf
    sTemp = sTemp & "FR_REPLACEALL = " & CheckFlags(FR_REPLACEALL, .Flags) & vbCrLf
    sTemp = sTemp & "FR_DOWN = " & CheckFlags(FR_DOWN, .Flags) & vbCrLf
    sTemp = sTemp & "FR_MATCHCASE = " & CheckFlags(FR_MATCHCASE, .Flags) & vbCrLf
    sTemp = sTemp & "FR_WHOLEWORD = " & CheckFlags(FR_WHOLEWORD, .Flags)
    MsgBox sTemp, vbOKOnly + vbInformation, "Find/Replace Parameters"
  End With
  
End Function




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




'=============================================================================================================
'                           SAVE / OPEN FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' OFN_ALLOWMULTISELECT
'          Specifies that the File Name list box allows multiple selections.
'          If you also set the OFN_EXPLORER flag, the dialog box uses the
'          Explorer-style user interface; otherwise, it uses the old-style
'          user interface.
'
'          If the user selects more than one file, the strFile buffer
'          returns the path to the current directory followed by the
'          filenames of the selected files. The nFileOffset member is the
'          offset to the first filename, and the nFileExtension member is
'          not used. For Explorer-style dialog boxes, the directory and
'          filename strings are NULL separated, with an extra NULL character
'          after the last filename. This format enables the Explorer-style
'          dialogs to return long filenames that include spaces. For old-style
'          dialog boxes, the directory and filename strings are separated by
'          spaces and the function uses short filenames for filenames with
'          spaces. You can use the FindFirstFile function to convert between
'          long and short filenames.
'
'          If you specify a custom template for an old-style dialog box, the
'          definition of the File Name list box must contain the
'          LBS_EXTENDEDSEL value.
'
' OFN_CREATEPROMPT
'          If the user specifies a file that does not exist, this flag causes
'          the dialog box to prompt the user for permission to create the file.
'          If the user chooses to create the file, the dialog box closes and the
'          function returns the specified name; otherwise, the dialog box
'          remains open.
'
' OFN_ENABLEHOOK
'          Enables the hook function specified in the lpfnHook member.
'
' OFN_ENABLETEMPLATE
'          Indicates that the lpTemplateName member points to the name of a
'          dialog template resource in the module identified by the hInstance
'          member.
'
'          If the OFN_EXPLORER flag is set, the system uses the specified
'          template to create a dialog box that is a child of the default
'          Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the
'          system uses the template to create an old-style dialog box that
'          replaces the default dialog box.
'
' OFN_ENABLETEMPLATEHANDLE
'          Indicates that the hInstance member identifies a data block that
'          contains a preloaded dialog box template. The system ignores the
'          lpTemplateName if this flag is specified.
'
'          If the OFN_EXPLORER flag is set, the system uses the specified
'          template to create a dialog box that is a child of the default
'          Explorer-style dialog box. If the OFN_EXPLORER flag is not set,
'          the system uses the template to create an old-style dialog box
'          that replaces the default dialog box.
'
' OFN_EXPLORER
'          Indicates that any customizations made to the Open or Save As dialog
'           box use the new Explorer-style customization methods. For more
'          information, see the "Explorer-Style Hook Procedures" and
'          "Explorer-Style Custom Templates" sections of the Common Dialog Box
'          Library overview.
'
'          By default, the Open and Save As dialog boxes use the Explorer-style
'          user interface regardless of whether this flag is set. This flag
'          is necessary only if you provide a hook procedure or custom
'          template, or set the OFN_ALLOWMULTISELECT flag.
'
'          If you want the old-style user interface, omit the OFN_EXPLORER
'          flag and provide a replacement old-style template or hook
'          procedure. If you want the old style but do not need a custom
'          template or hook procedure, simply provide a hook procedure
'          that always returns FALSE.
'
' OFN_EXTENSIONDIFFERENT
'          Specifies that the user typed a filename extension that differs from
'          the extension specified by lpstrDefExt. The function does not use this
'          flag if lpstrDefExt is NULL.
'
' OFN_FILEMUSTEXIST
'          Specifies that the user can type only names of existing files in the
'          File Name entry field. If this flag is specified and the user enters
'          an invalid name, the dialog box procedure displays a warning in a
'          message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is
'          also used.
'
' OFN_HIDEREADONLY
'          Hides the Read Only check box.
'
' OFN_LONGNAMES
'          For old-style dialog boxes, this flag causes the dialog box to use
'          long filenames. If this flag is not specified, or if the
'          OFN_ALLOWMULTISELECT flag is also set, old-style dialog boxes use
'          short filenames (8.3 format) for filenames with spaces.
'
'          Explorer-style dialog boxes ignore this flag and always display
'          long filenames.
'
' OFN_NOCHANGEDIR
'          Restores the current directory to its original value if the user
'          changed the directory while searching for files.
'
' OFN_NODEREFERENCELINKS
'          Directs the dialog box to return the path and filename of the
'          selected shortcut (.LNK) file. If this value is not given, the
'          dialog box returns the path and filename of the file referenced
'          by the shortcut
'
' OFN_NOLONGNAMES
'          For old-style dialog boxes, this flag causes the dialog box to
'          use short filenames (8.3 format).
'
'          Explorer-style dialog boxes ignore this flag and always display
'          long filenames.
'
' OFN_NONETWORKBUTTON
'          Hides and disables the Network button.
'
' OFN_NOREADONLYRETURN
'          Specifies that the returned file does not have the Read Only
'          check box checked and is not in a write-protected directory.
'
' OFN_NOTESTFILECREATE
'          Specifies that the file is not created before the dialog box is
'          closed. This flag should be specified if the application saves
'          the file on a create-nonmodify network sharepoint. When an
'          application specifies this flag, the library does not check for
'          write protection, a full disk, an open drive door, or network
'          protection. Applications using this flag must perform file
'          operations carefully, because a file cannot be reopened once it
'          is closed.
'
' OFN_NOVALIDATE
'          Specifies that the common dialog boxes allow invalid characters
'          in the returned filename. Typically, the calling application uses
'          a hook procedure that checks the filename by using the
'          FILEOKSTRING message. If the text box in the edit control is
'          empty or contains nothing but spaces, the lists of files and
'          directories are updated. If the text box in the edit control
'          contains anything else, nFileOffset and nFileExtension are set
'          to values generated by parsing the text. No default extension
'          is added to the text, nor is text copied to the buffer specified
'          by lpstrFileTitle.
'
'          If the value specified by nFileOffset is less than zero, the
'          filename is invalid. Otherwise, the filename is valid, and
'          nFileExtension and nFileOffset can be used as if the
'          OFN_NOVALIDATE flag had not been specified.
'
' OFN_OVERWRITEPROMPT
'          Causes the Save As dialog box to generate a message box if the
'          selected file already exists. The user must confirm whether to
'          overwrite the file.
'
' OFN_PATHMUSTEXIST
'          Specifies that the user can type only valid paths and filenames.
'          If this flag is used and the user types an invalid path and
'          filename in the File Name entry field, the dialog box function
'          displays a warning in a message box.
'
' OFN_READONLY
'          Causes the Read Only check box to be checked initially when the
'          dialog box is created. This flag indicates the state of the
'          Read Only check box when the dialog box is closed.
'
' OFN_SHAREAWARE
'          Specifies that if a call to the OpenFile function fails because
'          of a network sharing violation, the error is ignored and the
'          dialog box returns the selected filename.
'
'          If this flag is not set, the dialog box notifies your hook
'          procedure when a network sharing violation occurs for the
'          filename specified by the user. If you set the OFN_EXPLORER
'          flag, the dialog box sends the CDN_SHAREVIOLATION message to
'          the hook procedure. If you do not set OFN_EXPLORER, the dialog
'          box sends the SHAREVISTRING registered message to the hook
'          procedure.
'
' OFN_SHOWHELP
'          Causes the dialog box to display the Help button. The hwndOwner
'          member must specify the window to receive the HELPMSGSTRING
'          registered messages that the dialog box sends when the user
'          clicks the Help button.
'
'          An Explorer-style dialog box sends a CDN_HELP notification
'          message to your hook procedure when the user clicks the Help
'          button.
'
'
'=============================================================================================================
'                             COLOR FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' CC_ENABLEHOOK
'          Enables the hook procedure specified in the lpfnHook member of
'          this structure. This flag is used only to initialize the dialog
'          box.
'
' CC_ENABLETEMPLATE
'          Indicates that the hInstance and lpTemplateName members specify
'          a dialog box template to use in place of the default template.
'          This flag is used only to initialize the dialog box.
'
' CC_ENABLETEMPLATEHANDLE
'          Indicates that the hInstance member identifies a data block that
'          contains a preloaded dialog box template. The system ignores the
'          lpTemplateName member if this flag is specified. This flag is
'          used only to initialize the dialog box.
'
' CC_FULLOPEN
'          Causes the dialog box to display the additional controls that
'          allow the user to create custom colors. If this flag is not set,
'          the user must click the Define Custom Color button to display the
'          custom color controls.
'
' CC_PREVENTFULLOPEN
'          Disables the Define Custom Colors button.
'
' CC_RGBINIT  Causes the dialog box to use the color specified in the
'          rgbResult member as the initial color selection.
'
' CC_SHOWHELP Causes the dialog box to display the Help button. The hwndOwner
'          member must specify the window to receive the HELPMSGSTRING
'          registered messages that the dialog box sends when the user clicks
'          the Help button.
'
'
'=============================================================================================================
'                              FONT FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' CF_APPLY
'          Causes the dialog box to display the Apply button. You should
'          provide a hook procedure to process WM_COMMAND messages for the Apply
'          button. The hook procedure can send the WM_CHOOSEFONT_GETLOGFONT
'          message to the dialog box to retrieve the address of the LOGFONT
'          structure that contains the current selections for the font.
'
' CF_ANSIONLY
'          This flag is obsolete. To limit font selections to all scripts
'          except those that use the OEM or Symbol character sets, use
'          CF_SCRIPTSONLY. To get the Windows 3.1 CF_ANSIONLY behavior, use
'          CF_SELECTSCRIPT and specify ANSI_CHARSET in the lfCharSet member
'          of the LOGFONT structure pointed to by lpLogFont.
'
' CF_BOTH
'          Causes the dialog box to list the available printer and screen
'          fonts.  The hDC member identifies the device context (or
'          information context) associated with the printer. This flag is
'          a combination of the CF_SCREENFONTS and CF_PRINTERFONTS flags.
'
' CF_TTONLY
'          Specifies that ChooseFont should only enumerate and allow the
'          selection of TrueType fonts.
'
' CF_EFFECTS
'          Causes the dialog box to display the controls that allow the
'          user to specify strikeout, underline, and text color options.
'          If this flag is set, you can use the rgbColors member to specify
'          the initial text color. You can use the lfStrikeOut and
'          lfUnderline members of the LOGFONT structure pointed to by
'          lpLogFont to specify the initial settings of the strikeout and
'          underline check boxes.  ChooseFont can use these members to
'          return the user’s selections.
'
' CF_ENABLEHOOK
'          Enables the hook procedure specified in the lpfnHook member
'          of this structure.
'
' CF_ENABLETEMPLATE
'          Indicates that the hInstance and lpTemplateName members specify
'          a dialog box template to use in place of the default template.
'
' CF_ENABLETEMPLATEHANDLE
'          Indicates that the hInstance member identifies a
'          data block that contains a preloaded dialog box template. The
'          system ignores the lpTemplateName member if this flag is
'          specified.
'
' CF_FIXEDPITCHONLY
'          Specifies that ChooseFont should select only fixed-pitch fonts.
'
' CF_FORCEFONTEXIST
'          Specifies that ChooseFont should indicate an error condition if
'          the user attempts to select a font or style that does not exist.
'
' CF_INITTOLOGFONTSTRUCT
'          Specifies that ChooseFont should use the LOGFONT structure pointed
'          to by the lpLogFont member to initialize the dialog box controls.
'
' CF_LIMITSIZE
'          Specifies that ChooseFont should select only font sizes within the
'          range specified by the nSizeMin and nSizeMax members.
'
' CF_NOOEMFONTS
'          Same as the CF_NOVECTORFONTS flag.
'
' CF_NOFACESEL
'          When using a LOGFONT structure to initialize the dialog box
'          controls, use this flag to selectively prevent the dialog box from
'          displaying an initial selection for the font name combo box. This
'          is useful when there is no single font name that applies to the
'          text selection.
'
' CF_NOSCRIPTSEL
'          Disables the Script combo box. When this flag is set, the lfCharSet
'          member of the LOGFONT structure is set to DEFAULT_CHARSET when
'          ChooseFont returns. This flag is used only to initialize the dialog
'          box.
'
' CF_NOSTYLESEL
'          When using a LOGFONT structure to initialize the dialog box controls,
'          use this flag to selectively prevent the dialog box from displaying
'          an initial selection for the font style combo box. This is useful
'          when there is no single font style that applies to the text selection.
'
' CF_NOSIZESEL
'          When using a LOGFONT structure to initialize the dialog box controls,
'          use this flag to selectively prevent the dialog box from displaying
'          an initial selection for the font size combo box. This is useful when
'          there is no single font size that applies to the text selection.
'
' CF_NOSIMULATIONS
'          Specifies that ChooseFont should not allow graphics device interface
'          (GDI) font simulations.
'
' CF_NOVECTORFONTS
'          Specifies that ChooseFont should not allow vector font selections.
'
' CF_NOVERTFONTS
'          Causes the Font dialog box to list only horizontally oriented fonts.
'
' CF_PRINTERFONTS
'          Causes the dialog box to list only the fonts supported by the printer
'          associated with the device context (or information context) identified
'          by the hDC member.
'
' CF_SCALABLEONLY
'          Specifies that ChooseFont should allow only the selection of scalable
'          fonts. (Scalable fonts include vector fonts, scalable printer fonts,
'          TrueType fonts, and fonts scaled by other technologies.)
'
' CF_SCREENFONTS
'          Causes the dialog box to list only the screen fonts supported by the
'          system.
'
' CF_SCRIPTSONLY
'          Specifies that ChooseFont should allow selection of fonts for all
'          non-OEM and Symbol character sets, as well as the ANSI character set.
'          This supersedes the CF_ANSIONLY value.
'
' CF_SELECTSCRIPT
'          When specified on input, only fonts with the character set identified
'          in the lfCharSet member of the LOGFONT structure are displayed. The
'          user will not be allowed to change the character set specified in the
'          Scripts combo box.
'
' CF_SHOWHELP
'          Causes the dialog box to display the Help button. The hwndOwner member
'          must specify the window to receive the HELPMSGSTRING registered
'          messages that the dialog box sends when the user clicks the Help
'          button.
'
' CF_USESTYLE
'          Specifies that the lpszStyle member points to a buffer that contains
'          style data that ChooseFont should use to initialize the Font Style
'          combo box. When the user closes the dialog box, ChooseFont copies
'          style data for the user’s selection to this buffer.
'
' CF_WYSIWYG
'          Specifies that ChooseFont should allow only the selection of fonts
'          available on both the printer and the display. If this flag is
'          specified, the CF_BOTH and CF_SCALABLEONLY flags should also be
'          specified.
'
'
'=============================================================================================================
'                              PRINT FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' PD_ALLPAGES
'          The default flag that indicates that the All radio button is
'          initially selected. This flag is used as a placeholder to
'          indicate that the PD_PAGENUMS and PD_SELECTION flags are not
'          specified.
'
' PD_COLLATE
'          Places a checkmark in the Collate check box when set on input.
'          When the PrintDlg function returns, this flag indicates that the
'          user selected the Collate option and the printer driver does not
'          support collation. In this case, the application must provide
'          collation. If PrintDlg sets the PD_COLLATE flag on return, the
'          dmCollate member of the DEVMODE structure is undefined.
'
' PD_DISABLEPRINTTOFILE
'          Disables the Print to File check box.
'
' PD_ENABLEPRINTHOOK
'          Enables the hook procedure specified in the lpfnPrintHook member.
'          This enables the hook procedure for the Print dialog box.
'
' PD_ENABLEPRINTTEMPLATE
'          Indicates that the hInstance and lpPrintTemplateName members
'          specify a dialog box template to use in place of the default
'          template for the Print dialog box.
'
' PD_ENABLEPRINTTEMPLATEHANDLE
'          Indicates that the hPrintTemplate member identifies a data
'          block that contains a preloaded dialog box template. The
'          system uses this template in place of the default template for
'          the Print dialog box. The system ignores the lpPrintTemplateName
'          member if this flag is specified.
'
' PD_ENABLESETUPHOOK
'          Enables the hook procedure specified in the lpfnSetupHook member.
'          This enables the hook procedure for the Print Setup dialog box.
'
' PD_ENABLESETUPTEMPLATE
'          Indicates that the hInstance and lpSetupTemplateName members
'          specify a dialog box template to use in place of the default
'          template for the Print Setup dialog box.
'
' PD_ENABLESETUPTEMPLATEHANDLE
'          Indicates that the hSetupTemplate member identifies a data block
'          that contains a preloaded dialog box template. The system uses
'          this template in place of the default template for the Print
'          Setup dialog box. The system ignores the lpSetupTemplateName
'          member if this flag is specified.
'
' PD_HIDEPRINTTOFILE
'          Hides the Print to File check box.
'
' PD_NOPAGENUMS
'          Disables the Pages radio button and the associated edit controls.
'
' PD_NOSELECTION
'          Disables the Selection radio button.
'
' PD_NOWARNING
'          Prevents the warning message from being displayed when there is no
'          default printer.
'
' PD_PAGENUMS
'          Causes the Pages radio button to be in the selected state when the
'          dialog box is created. When PrintDlg returns, this flag is set if
'          the Pages radio button is in the selected state.
'
' PD_PRINTSETUP
'          Causes the system to display the Print Setup dialog box rather
'          than the Print dialog box.
'
' PD_PRINTTOFILE
'          Causes the Print to File check box to be checked when the dialog
'          box is created.
'
'          When PrintDlg returns, this flag is set if the check box is
'          checked. In this case, the offset indicated by the wOutputOffset
'          member of the DEVNAMES structure contains the string "FILE:".
'          When you call the StartDoc function to start the printing
'          operation, specify this "FILE:" string in the lpszOutput member
'          of the DOCINFO structure. Specifying this string causes the print
'          subsystem to query the user for the name of the output file.
'
' PD_RETURNDC
'          Causes PrintDlg to return a device context matching the selections
'          the user made in the dialog box. The device context is returned in
'          hDC.
'
' PD_RETURNDEFAULT
'          The PrintDlg function does not display the dialog box. Instead, it
'          sets the hDevNames and hDevMode members to handles to DEVMODE and
'          DEVNAMES structures that are initialized for the system default
'          printer. Both hDevNames or hDevMode must be NULL, or PrintDlg
'          returns an error. If the system default printer is supported by an
'          old printer driver (earlier than Windows version 3.0), only
'          hDevNames is returned; hDevMode is NULL.
'
' PD_RETURNIC
'          Similar to the PD_RETURNDC flag, except that this flag returns an
'          information context rather than a device context. If neither
'          PD_RETURNDC nor PD_RETURNIC is specified, hDC is undefined on
'          output.
'
' PD_SELECTION
'          Causes the Selection radio button to be in the selected state when
'          the dialog box is created. When PrintDlg returns, this flag is
'          specified if the Selection radio button is selected. If neither
'          PD_PAGENUMS nor PD_SELECTION is set, the All radio button is
'          selected.
'
' PD_SHOWHELP
'          Causes the dialog box to display the Help button. The hwndOwner
'          member must specify the window to receive the HELPMSGSTRING
'          registered messages that the dialog box sends when the user clicks
'          the Help button.
'
' PD_USEDEVMODECOPIES
'          Same As PD_USEDEVMODECOPIESANDCOLLATE
'
' PD_USEDEVMODECOPIESANDCOLLATE
'          Disables the Copies edit control if the printer driver does not
'          support multiple copies, and disables the Collate checkbox if the
'          printer driver does not support collation. If this flag is not
'          specified, PrintDlg stores the user selections for the Copies and
'          Collate options in the dmCopies and dmCollate members of the
'          DEVMODE structure.
'
'          If this flag isn’t set, the copies and collate information is
'          returned in the DEVMODE structure if the driver supports multiple
'          copies and collation. If the driver doesn’t support multiple
'          copies and collation, the information is returned in the PRINTDLG
'          structure. This means that an application only has to look at
'          nCopies and PD_COLLATE to determine how many copies it needs to
'          render and whether it needs to print them collated.
'
'
'=============================================================================================================
'                            PAGE SETUP FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' PSD_DEFAULTMINMARGINS
'          Sets the minimum values that the user can specify for the page
'          margins to be the minimum margins allowed by the printer. This
'          is the default. This flag is ignored if the PSD_MARGINS and
'          PSD_MINMARGINS flags are also specified.
'
' PSD_DISABLEMARGINS
'          Disables the margin controls, preventing the user from setting
'          the margins.
'
' PSD_DISABLEORIENTATION
'          Disables the orientation controls, preventing the user from
'          setting the page orientation.
'
' PSD_DISABLEPAGEPAINTING
'          Prevents the dialog box from drawing the contents of the sample
'          page. If you enable a PagePaintHook hook procedure, you can
'          still draw the contents of the sample page.
'
' PSD_DISABLEPAPER
'          Disables the paper controls, preventing the user from setting
'          page parameters such as the paper size and source.
'
' PSD_DISABLEPRINTER
'          Disables the Printer button, preventing the user from invoking
'          a dialog box that contains additional printer setup information.
'
' PSD_ENABLEPAGEPAINTHOOK
'          Enables the hook procedure specified in the lpfnPagePaintHook
'          member.
'
' PSD_ENABLEPAGESETUPHOOK
'          Enables the hook procedure specified in the lpfnPageSetupHook
'          member.
'
' PSD_ENABLEPAGESETUPTEMPLATE
'          Indicates that the hInstance and lpPageSetupTemplateName
'          members specify a dialog box template to use in place of the
'          default template.
'
' PSD_ENABLEPAGESETUPTEMPLATEHANDLE
'          Indicates that the hPageSetupTemplate member identifies a data
'          block that contains a preloaded dialog box template. The system
'          ignores the lpPageSetupTemplateName member if this flag is
'          specified.
'
' PSD_INHUNDREDTHSOFMILLIMETERS
'          Indicates that hundredths of millimeters are the unit of
'          measurement for margins and paper size. The values in the
'          rtMargin, rtMinMargin, and ptPaperSize members are in hundredths
'          of millimeters. You can set this flag on input to override the
'          default unit of measurement for the user’s locale. When the
'          function returns, the dialog box sets this flag to indicate the
'          units used.
'
' PSD_INTHOUSANDTHSOFINCHES
'          Indicates that thousandths of inches are the unit of measurement
'          for margins and paper size. The values in the rtMargin,
'          rtMinMargin, and ptPaperSize members are in thousandths of
'          inches. You can set this flag on input to override the default
'          unit of measurement for the user’s locale. When the function
'          returns, the dialog box sets this flag to indicate the units
'          used.
'
' PSD_INWININIINTLMEASURE
'          Not implemented.
'
' PSD_MARGINS
'          Causes the system to use the values specified in the rtMargin
'          member as the initial widths for the left, top, right, and
'          bottom margins. If PSD_MARGINS is not set, the system sets the
'          initial widths to one inch for all margins.
'
' PSD_MINMARGINS
'          Causes the system to use the values specified in the rtMinMargin
'          member as the minimum allowable widths for the left, top, right,
'          and bottom margins. The system prevents the user from entering a
'          width that is less than the specified minimum. If PSD_MINMARGINS
'          is not specified, the system sets the minimum allowable widths to
'          those allowed by the printer.
'
' PSD_NOWARNING
'          Prevents the system from displaying a warning message when there
'          is no default printer.
'
' PSD_RETURNDEFAULT
'          PageSetupDlg does not display the dialog box. Instead, it sets
'          the hDevNames and hDevMode members to handles to DEVMODE and
'          DEVNAMES structures that are initialized for the system default
'          printer. PageSetupDlg returns an error if either hDevNames or
'          hDevMode is not NULL.
'
' PSD_SHOWHELP
'          Causes the dialog box to display the Help button. The hwndOwner
'          member must specify the window to receive the HELPMSGSTRING
'          registered messages that the dialog box sends when the user
'          clicks the Help button.
'
'
'=============================================================================================================
'                            FIND / REPLACE FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' FR_DIALOGTERM
'          If set in a FINDMSGSTRING message, indicates that the
'          dialog box is closing. When you receive a message with
'          this flag set, the dialog box window handle returned by
'          the DLG_FindText or DLG_ReplaceText function is no longer valid.
'
' FR_DOWN
'          If set, the Down button of the direction radio buttons in
'          a Find dialog box is selected indicating that you should
'          search from the current location to the end of the document.
'          If not set, the Up button is selected so you should search
'          to the beginning of the document. You can set this flag to
'          initialize the dialog box. If set in a FINDMSGSTRING message,
'          indicates the user’s selection.
'
' FR_ENABLEHOOK
'          Enables the hook function specified in the lpfnHook member.
'          This flag is used only to initialize the dialog box.
'
' FR_ENABLETEMPLATE
'          Indicates that the hInstance and lpTemplateName members
'          specify a dialog box template to use in place of the default
'          template. This flag is used only to initialize the dialog
'          box.
'
' FR_ENABLETEMPLATEHANDLE
'          Indicates that the hInstance member identifies a data block
'          that contains a preloaded dialog box template. The system
'          ignores the lpTemplateName member if this flag is specified.
'
' FR_FINDNEXT
'          If set in a FINDMSGSTRING message, indicates that the user
'          clicked the Find Next button in a Find or Replace dialog box.
'          The lpstrFindWhat member specifies the string to search for.
'
' FR_HIDEUPDOWN
'          If set when initializing a Find dialog box, hides the search
'          direction radio buttons.
'
' FR_HIDEMATCHCASE
'          If set when initializing a Find or Replace dialog box, hides
'          the Match Case check box.
'
' FR_HIDEWHOLEWORD
'          If set when initializing a Find or Replace dialog box, hides
'          the Match Whole Word Only check box.
'
' FR_MATCHCASE
'          If set, the Match Case check box is checked indicating that
'          the search should be case-sensitive. If not set, the check box
'          is unchecked so the search should be case-insensitive. You can
'          set this flag to initialize the dialog box. If set in a
'          FINDMSGSTRING message, indicates the user’s selection.
'
' FR_NOMATCHCASE
'          If set when initializing a Find or Replace dialog box, disables
'          the Match Case check box.
'
' FR_NOUPDOWN
'          If set when initializing a Find dialog box, disables the search
'          direction radio buttons.
'
' FR_NOWHOLEWORD
'          If set when initializing a Find or Replace dialog box, disables
'          the Whole Word check box.
'
' FR_REPLACE
'          If set in a FINDMSGSTRING message, indicates that the user clicked
'          the Replace button in a Replace dialog box. The lpstrFindWhat member
'          specifies the string to be replaced and the lpstrReplaceWith member
'          specifies the replacement string.
'
' FR_REPLACEALL
'          If set in a FINDMSGSTRING message, indicates that the user clicked
'          the Replace All button in a Replace dialog box. The lpstrFindWhat
'          member specifies the string to be replaced and the lpstrReplaceWith
'          member specifies the replacement string.
'
' FR_SHOWHELP
'          Causes the dialog box to display the Help button. The hwndOwner
'          member must specify the window to receive the HELPMSGSTRING
'          registered messages that the dialog box sends when the user clicks
'          the Help button.
'
' FR_WHOLEWORD
'          If set, the Match Whole Word Only check box is checked indicating that
'          you should search only for whole words that match the search string.
'          If not set, the check box is unchecked so you should also search for
'          word fragments that match the search string. You can set this flag to
'          initialize the dialog box. If set in a FINDMSGSTRING message,
'          indicates the user’s selection.
'
'
'=============================================================================================================
'                         BROWSE FOR FOLDER FLAGS
'=============================================================================================================
' Flag     Meaning
'=============================================================================================================
' BIF_BROWSEFORCOMPUTER
'          Only returns computers. If the user selects anything other than a
'          computer, the OK button is grayed.
'
' BIF_BROWSEFORPRINTER
'          Only returns printers. If the user selects anything other than a
'          printer, the OK button is grayed.
'
' BIF_DONTGOBELOWDOMAIN
'          Does not include network folders below the domain level in the tree
'          view control.
'
' BIF_RETURNFSANCESTORS
'          Only returns file system ancestors. If the user selects anything other
'          than a file system ancestor, the OK button is grayed.
'
' BIF_RETURNONLYFSDIRS
'          Only returns file system directories. If the user selects folders that
'          are not part of the file system, the OK button is grayed.
'
' BIF_STATUSTEXT
'          Includes a status area in the dialog box. The callback function can
'          set the status text by sending messages to the dialog box.
'
' BIF_BROWSEINCLUDEFILES
'          UNDOCUMENTED : Browse for everything
' BIF_EDITBOX
'          UNDOCUMENTED : Displays an editable TextBox on the dialog that
'          displays the current directory name
'
' BIF_VALIDATE
'          UNDOCUMENTED : Insist on valid result (or CANCEL)
'
'
'=============================================================================================================
'                            WINHELP FLAGS
'-------------------------------------------------------------------------------------------------------------
'
'BOOL WinHelp(
'    HWND hWndMain,    // handle of window requesting Help
'    LPCTSTR lpszHelp, // address of directory-path string
'    UINT uCommand,    // type of Help
'    DWORD dwData      // additional data
');
'
'=============================================================================================================
'  uCommand        Action                           dwData
'=============================================================================================================
'
' HELP_COMMAND     Executes a Help macro or         Address of a string that specifies
'                  macro string.                    the name of the Help macro(s) to
'                                                   execute. If the string specifies
'                                                   multiple macros names, the names
'                                                   must be separated by semicolons.
'                                                   You must use the short form of the
'                                                   macro name for some macros because
'                                                   Help does not support the long name.
'
' HELP_CONTENTS    Displays the topic specified     Ignored, set to 0.
'                  by the Contents option in the
'                  [OPTIONS] section of the .HPJ
'                  file. This is for backward
'                  compatibility. New applica-
'                  tions should provide a .CNT
'                  file and use the HELP_FINDER
'                  command.
'
' HELP_CONTEXT     Displays the topic identified    Unsigned long integer containing the
'                  by the specified context         context identifier for the topic.
'                  identifier defined in the
'                  [MAP] section of the .HPJ
'                  file.
'
' HELP_CONTEXTPOPUP  Displays, in a pop-up window,  Unsigned long integer containing the
'                  the topic identified by the      context identifier for a topic.
'                  specified context identifier
'                  defined in the [MAP] section
'                  of the .HPJ file.
'
' HELP_FORCEFILE   Ensures that WinHelp is          Ignored, set to 0.
'                  displaying the correct help
'                  file. If the incorrect help
'                  file is being displayed,
'                  WinHelp opens the correct
'                  one; otherwise, there is no
'                  action.
'
' HELP_HELPONHELP  Displays help on how to use      Ignored, set to 0.
'                  Windows Help, if the
'                  WINHELP.HLP file is
'                  available.
'
' HELP_INDEX       Displays the Index in the Help   Ignored, set to 0.
'                  Topics dialog box. This command
'                  is for backward compatibility.
'                  New applications should use the
'                  HELP_FINDER command.
'
' HELP_KEY         Displays the topic in the        Address of a keyword string.
'                  keyword table that matches the
'                  specified keyword, if there is
'                  an exact match. If there is
'                  more than one match, displays
'                  the Index with the topics
'                  listed in the Topics Found list
'                  box.
'
' HELP_MULTIKEY    Displays the topic specified     Address of a MULTIKEYHELP structure
'                  by a keyword in an alternative   that specifies a table footnote
'                  keyword table.                   character and a keyword.
'
' HELP_PARTIALKEY  Displays the topic in the        Address of a keyword string.
'                  keyword table that matches the
'                  specified keyword, if there is
'                  an exact match. If there is
'                  more than one match, displays
'                  the Index tab. To display the
'                  Index without passing a keyword
'                  (the third result), you should
'                  use a pointer to an empty
'                  string.
'
' HELP_QUIT        Informs the Help application     Ignored, set to 0.
'                  that it is no longer needed.
'                  If no other applications have
'                  asked for Help, Windows closes
'                  the Help application.
'
' HELP_SETCONTENTS Specifies the Contents topic.    Unsigned long integer containing
'                  The Help application displays    the context identifier for the
'                  this topic when the user clicks  Contents topic.
'                  the Contents button.
'
' HELP_SETINDEX    Specifies a keyword table to be  Unsigned long integer containing
'                  displayed in the Index of the    the context identifier for the
'                  Help Topics dialog box.          Index topic.
'
' HELP_SETWINPOS   Displays the Help window, if it  Address of a HELPWININFO structure
'                  is minimized or in memory, and   that specifies the size and
'                  sets its size and position as    position of either a primary or
'                  specified.                       secondary Help window.
'
' HELP_TAB         UNDOCUMENTED - Opens the         Ignored, set to 0.
'                  contents tab of the help
'                  file specified.
'
