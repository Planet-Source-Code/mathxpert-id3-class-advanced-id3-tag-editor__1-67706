VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDIPlusCandy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' I give credit to all the websites for uncovering the pieces
' of the puzzle to create this wonderful GDI+ candy class. It
' is a dream come true!

Private Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"

Private Enum GpStatus
    Ok
    GenericError
    InvalidParameter
    OutOfMemory
    ObjectBusy
    InsufficientBuffer
    NotImplemented
    Win32Error
    WrongState
    Aborted
    FileNotFound
    ValueOverflow
    AccessDenied
    UnknownImageFormat
    FontFamilyNotFound
    FontStyleNotFound
    NotTrueTypeFont
    UnsupportedlusVersion
    lusNotInitialized
    PropertyNotFound
    PropertyNotSupported
End Enum

Private Type CLSID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Enum ImageCodecFlags
    ImageCodecFlagsEncoder = &H1
    ImageCodecFlagsDecoder = &H2
    ImageCodecFlagsSupportBitmap = &H4
    ImageCodecFlagsSupportVector = &H8
    ImageCodecFlagsSeekableEncode = &H10
    ImageCodecFlagsBlockingDecode = &H20
    ImageCodecFlagsBuiltin = &H10000
    ImageCodecFlagsSystem = &H20000
    ImageCodecFlagsUser = &H40000
End Enum

Private Type ImageCodecInfo
    CLSID As CLSID
    FormatDescriptionPtr As Long
    CodecNamePtr As Long
    DllNamePtr As Long
    FormatID As CLSID
    FilenameExtensionPtr As Long
    MimeTypePtr As Long
    Flags As ImageCodecFlags
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPatternPtr As Long
    SigMaskPtr As Long
End Type

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
    Guid As CLSID
    NumberOfValues As Long
    Type As Long
    ValuePtr As Long
End Type

Private Type EncoderParameters
    Count As Long
    Parameter(0 To 15) As EncoderParameter
End Type

Private Type PICTDESC
    size As Long
    Type As Long
    hBmpOrIcon As Long
    hPal As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, lngInput As GdiplusStartupInput, Optional ByVal Output As Long = 0) As GpStatus
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal NumEncoders As Long, ByVal size As Long, Encoders As Any) As GpStatus
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (NumEncoders As Long, size As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal Stream As Long, Bitmap As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal Hbm As Long, ByVal hPal As Long, Bitmap As Long) As GpStatus
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, HbmReturn As Long, ByVal Background As Long) As GpStatus
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As Long, ClsidEncoder As CLSID, EncoderParams As Any) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long
Private Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal pstm As Any, phglobal As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32" (lpPictDesc As PICTDESC, riid As CLSID, ByVal fOwn As Long, lplpvObj As Object)
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As CLSID) As Long

Private Const GMEM_ZEROINIT As Long = &H40

Dim m_Initialized As Boolean
Dim m_Token As Long

Public Property Get SuccessfullyInitialized() As Boolean
    SuccessfullyInitialized = m_Initialized
End Property

Private Function PtrToStr(ByVal Pointer As Long) As String
    Dim sBuf As String
    sBuf = Space$(lstrlenW(ByVal Pointer))
    lstrcpyW ByVal StrPtr(sBuf), ByVal Pointer
    PtrToStr = sBuf
End Function

Private Function GetEncoderCLSID(ByVal MIMEType As String, CLSID As CLSID) As Long
    Dim num&, size&, idx&, i&, ICI() As ImageCodecInfo, buffer() As Byte
    If m_Initialized Then
        GdipGetImageEncodersSize num, size
        If size > 0 Then
            ReDim ICI(1 To size)
            ReDim buffer(1 To size)
            GdipGetImageEncoders num, size, buffer(1)
            RtlMoveMemory ICI(1), buffer(1), Len(ICI(1)) * num
            For i = 1 To num
                If StrComp(PtrToStr(ICI(i).MimeTypePtr), MIMEType, vbTextCompare) = 0 Then
                    CLSID = ICI(i).CLSID
                    idx = i
                    Exit For
                End If
            Next
            Erase ICI
            Erase buffer
        End If
    End If
    GetEncoderCLSID = idx
End Function

Private Sub InitializeGDIPlus()
    Dim GpInput As GdiplusStartupInput
    GpInput.GdiplusVersion = 1
    m_Initialized = (GdiplusStartup(m_Token, GpInput) = Ok)
End Sub

Private Function HandleToPicture(ByVal hGDIHandle As Long, ByVal ObjectType As PictureTypeConstants, Optional ByVal hPal As Long = 0) As StdPicture
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As CLSID
    Dim oPicture As IPicture
    
    With tPictDesc
        .size = Len(tPictDesc)
        .Type = ObjectType
        .hBmpOrIcon = hGDIHandle
        .hPal = hPal
    End With
    
    With IID_IPicture
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    OleCreatePictureIndirect tPictDesc, IID_IPicture, True, oPicture
    Set HandleToPicture = oPicture
    Set oPicture = Nothing
End Function

Private Function CreateStreamOnData(bytData() As Byte, oStream As IUnknown) As Long
    ' Returns 0 if successful, or -1 if unsuccessful
    On Error GoTo err
    
    Dim lB As Long
    Dim lngByteCount As Long
    Dim hMem As Long
    Dim lpMem As Long
    Dim ret As Long
    
    ret = -1
    lB = LBound(bytData)
    lngByteCount = UBound(bytData) - lB + 1
    hMem = GlobalAlloc(&H2 Or GMEM_ZEROINIT, lngByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            RtlMoveMemory ByVal lpMem, bytData(lB), lngByteCount
            GlobalUnlock hMem
            If CreateStreamOnHGlobal(hMem, True, oStream) = 0 Then
                ret = 0
            End If
        End If
    End If
err:
    CreateStreamOnData = ret
End Function

Private Function GetDataFromStream(oStream As IUnknown, bytData() As Byte) As Long
    ' Returns 0 if successful, or -1 if unsuccessful
    On Error GoTo err
    
    Dim lngByteCount As Long
    Dim hMem As Long
    Dim lpMem As Long
    Dim ret As Long
    
    ret = -1
    If GetHGlobalFromStream(oStream, hMem) = 0 Then
        lngByteCount = GlobalSize(hMem)
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            ReDim bytData(0 To lngByteCount - 1)
            RtlMoveMemory bytData(0), ByVal lpMem, lngByteCount
            GlobalUnlock hMem
            ret = 0
        End If
    End If
err:
    GetDataFromStream = ret
End Function

Public Function DataToImage(ByVal Data As String) As StdPicture
    Dim bData() As Byte
    Dim lRet As Long
    Dim lBmp As Long
    Dim hBmp As Long
    Dim oStream As IUnknown
    
    If m_Initialized Then
        bData = StrConv(Data, vbFromUnicode)
        If CreateStreamOnData(bData, oStream) = 0 Then
            Erase bData
            lRet = GdipCreateBitmapFromStream(ObjPtr(oStream), lBmp)
            If lRet = Ok Then
                lRet = GdipCreateHBITMAPFromBitmap(lBmp, hBmp, 0)
                If lRet = Ok Then
                    Set DataToImage = HandleToPicture(hBmp, vbPicTypeBitmap)
                End If
            End If
        End If
    End If
End Function

Public Function ImageToData(ByVal Image As StdPicture, Optional ByVal MIMEType As String = "image/jpeg") As String
    Dim bData() As Byte
    Dim lRet As Long
    Dim lBmp As Long
    Dim oStream As IUnknown
    Dim PicEncoder As CLSID
    Dim tParams As EncoderParameters
    
    If m_Initialized Then
        lRet = GdipCreateBitmapFromHBITMAP(Image.Handle, 0, lBmp)
        If lRet = Ok Then
            lRet = GetEncoderCLSID(MIMEType, PicEncoder)
            If lRet <> 0 Then
                tParams.Count = 1
                With tParams.Parameter(0)
                    CLSIDFromString StrPtr(EncoderQuality), .Guid
                    .NumberOfValues = 1
                    .Type = 4 ' EncoderParameterValueTypeLong
                    .ValuePtr = VarPtr(75) ' Set the quality to 75
                End With
                CreateStreamOnHGlobal 0, True, oStream 'Create an empty stream
                lRet = GdipSaveImageToStream(lBmp, ObjPtr(oStream), PicEncoder, tParams)
                If lRet = Ok Then
                    If GetDataFromStream(oStream, bData) = 0 Then
                        ImageToData = StrConv(bData, vbUnicode)
                        Erase bData
                    End If
                End If
            End If
            GdipDisposeImage lBmp
        End If
    End If
End Function

Private Sub TerminateGDIPlus()
    GdiplusShutdown m_Token
End Sub

Private Sub Class_Initialize()
    InitializeGDIPlus
End Sub

Private Sub Class_Terminate()
    TerminateGDIPlus
End Sub
