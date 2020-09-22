Attribute VB_Name = "m_imutils"
Public Declare Function BitBlt Lib "gdi32" _
  (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetDesktopWindow Lib _
   "user32" () As Long

Public Declare Function GetWindowDC Lib _
   "user32" (ByVal hWnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32" _
   (ByVal hWnd As Long, ByVal hdc As Long) As Long
'--end block--'

Public Const SRCCOPY = &HCC0020

Public Declare Function GrayScale Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean) As Integer
Public Declare Function getDesktop Lib "ImageUtils.dll" (ByVal strFileName As String, ByVal blnEnableOverWrite As Boolean, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal blnJpeg As Boolean, ByVal JPGCompressQuality As Integer) As Integer
Public Declare Function ConvertBMPtoJPG Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal JPGCompressQuality As Integer, ByVal blnKeepBMP As Boolean) As Integer
Public Declare Function ConvertJPGtoBMP Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean, ByVal blnKeepJPG As Boolean) As Integer
Public Declare Function RotateRight Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean) As Integer
Public Declare Function RotateLeft Lib "ImageUtils.dll" (ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnEnableOverWrite As Boolean) As Integer
