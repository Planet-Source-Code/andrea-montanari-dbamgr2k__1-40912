VERSION 5.00
Begin VB.Form fCredits 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   305
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "fCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020    ' (DWORD) dest = source

Const kCredits = "Alex|Allanon|Arun Kumar|Francesco Bancalà|Fibia FBI|Fred Just|Gert E.R. Drapers|Gianluca Hotz|Jan Nielsen|Jordan Russel|Karl E. Peterson|KPDTeam@Allapi.net|Lars Broberg|Narayana Vyas Kondreddi|P. R.|Paolo Fisco|Peter Schmid|Peter Storz|Peter Swaniker|Planet Source Code|Rahul Sharma|Robert Vallee|Roberto Gismondi|Vincenzo Morgante|Yusuf Incekara"
Const kDelta = 1


Private m_sCredit() As String
Private m_lRowYpos As Long
Private m_lRowHeight As Long
Private m_lRegionHeight As Long
Private m_lRowWidth As Long
Private m_iUbound As Integer
Private m_iItemXform As Integer
Private m_iCurrent As Integer
Private m_iTik2Write As Integer

Private Sub Form_Load()

    m_sCredit = Split(kCredits, "|")
    m_iUbound = UBound(m_sCredit)
    
    m_lRowHeight = Me.TextHeight("")
    m_lRowWidth = Me.ScaleWidth
    
    m_iItemXform = ((Me.ScaleHeight \ m_lRowHeight) \ 2) - 1
    m_lRowYpos = Me.ScaleHeight - (2 * m_lRowHeight)
    m_lRegionHeight = Me.ScaleHeight
        
    CenterForm Me
    Me.Caption = MyLoadResString(k_Thanks)
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    TickScroll
End Sub

Private Sub TickScroll()

    Dim lRet As Long
        
    lRet = BitBlt(Me.hDC, 0, -kDelta, m_lRowWidth, m_lRegionHeight, Me.hDC, 0, 0, SRCCOPY)
        
    If m_iTik2Write = 0 Then PrintCurrent
    
    m_iTik2Write = m_iTik2Write + 1
    If m_iTik2Write = 25 Then m_iTik2Write = 0

End Sub
Private Sub PrintCurrent()

    Dim lWidth As Long
    
    If m_iCurrent >= 0 And m_iCurrent <= m_iUbound Then
        lWidth = Me.TextWidth(m_sCredit(m_iCurrent))
        Me.CurrentX = (m_lRowWidth - lWidth) \ 2
        Me.CurrentY = m_lRowYpos
        Me.Print m_sCredit(m_iCurrent)
    End If
    m_iCurrent = m_iCurrent + 1
    If m_iCurrent > m_iUbound Then m_iCurrent = -m_iItemXform

End Sub
