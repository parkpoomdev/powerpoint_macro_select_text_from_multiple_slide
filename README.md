# PowerPoint VBA – Copy Selected Slides as Plain Text

This repository contains a PowerPoint VBA macro that allows you to **select multiple slides from the left navigation pane** and **copy all text from those slides into the clipboard as plain text**.

This is useful when exporting content for documentation, AI processing, script extraction, versioning, or transferring slide content to external editors such as VSCode, Word, or Markdown.

---

## ✨ Features

## 📌 Installation

1. Open PowerPoint  
2. Press `ALT + F11` to open the **VBA editor**  
3. Go to **Insert → Module**  
4. Paste the entire macro code into the module  
5. Save your presentation as **.pptm** (macro-enabled)

---

<img width="1207" height="623" alt="image" src="https://github.com/user-attachments/assets/57be7eed-e2ea-454f-8c75-4fc83ee1cbdb" />

<img width="392" height="152" alt="image" src="https://github.com/user-attachments/assets/5ab10e8b-e6b9-42cf-968a-e381a4033f3f" />


===== Slide 25 =====

Finibus Bonorum
Sed ut perspiciatis unde omnis iste natus.
Nemo enim ipsam voluptatem quia voluptas.
Neque porro quisquam est, qui dolorem.
Quis autem vel eum iure reprehenderit.
At vero eos et accusamus et iusto odio.
Et harum quidem rerum facilis est et.

===== Slide 26 =====

Lorem Ipsum Dolor
Lorem ipsum dolor sit amet, consectetur.
Sed do eiusmod tempor incididunt ut labore.
Ut enim ad minim veniam, quis nostrud.
Duis aute irure dolor in reprehenderit.
Excepteur sint occaecat cupidatat non.
Sunt in culpa qui officia deserunt mollit.

===== Slide 27 =====

Why do we use it?
t represents a neutral visual placeholder.
Prevents distraction from meaningful content.
Showcases the font weight and kerning.
Simulates realistic paragraph distribution.
Various versions have evolved over the years.
Essential tool for designers and developers.

---

- Copy text from **one or multiple selected slides**
- Reads text from **all shapes** containing text frames
- Outputs **plain Unicode text** (no formatting)
- Works in **PowerPoint 32-bit and 64-bit**
- Uses **Windows API clipboard functions** (no MSForms dependency)
- VSCode-friendly (all comments in English, UTF-8 safe)

---


## 🚀 How to Use

1. In the left slide navigation pane (thumbnail view),  
   **Ctrl + click** to select multiple slides  
2. Go to **Developer → Macros**  
3. Choose:  
   `CopySelectedSlidesText_Plain`  
4. Click **Run**  
5. All extracted text is now in your **clipboard**, ready to paste anywhere.

---

## 🧩 Full VBA Code

```vb
Option Explicit

' ==== Windows API declarations for clipboard handling (compatible with 32/64-bit Office) ====
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, _
        ByVal Source As LongPtr, _
        ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As Long, _
        ByVal Source As Long, _
        ByVal Length As Long)
#End If

Const CF_UNICODETEXT As Long = 13
Const GMEM_MOVEABLE As Long = &H2

' ===========================================================
' Function: Copy a string (Unicode) to the Windows clipboard
' ===========================================================
Sub CopyToClipboard(ByVal s As String)
#If VBA7 Then
    Dim cbSize As LongPtr
    Dim hGlobal As LongPtr
    Dim pGlobal As LongPtr
#Else
    Dim cbSize As Long
    Dim hGlobal As Long
    Dim pGlobal As Long
#End If

    cbSize = (Len(s) + 1) * 2

    hGlobal = GlobalAlloc(GMEM_MOVEABLE, cbSize)
    If hGlobal = 0 Then Exit Sub

    pGlobal = GlobalLock(hGlobal)
    If pGlobal <> 0 Then
        CopyMemory pGlobal, StrPtr(s), cbSize
        GlobalUnlock hGlobal

        If OpenClipboard(0) <> 0 Then
            EmptyClipboard
            SetClipboardData CF_UNICODETEXT, hGlobal
            CloseClipboard
        End If
    End If
End Sub

' ===========================================================
' Macro: Copy plain text from ALL slides selected in the left
' navigation pane (multi-select supported)
' ===========================================================
Sub CopySelectedSlidesText_Plain()
    Dim sel As Selection
    Dim sld As Slide
    Dim shp As Shape
    Dim txt As String

    Set sel = ActiveWindow.Selection

    If sel Is Nothing Then
        MsgBox "No selection detected. Please select one or more slides from the left slide navigation pane.", vbExclamation
        Exit Sub
    End If

    If sel.Type = ppSelectionSlides Then
        For Each sld In sel.SlideRange
            txt = txt & "===== Slide " & sld.SlideIndex & " =====" & vbCrLf
            For Each shp In sld.Shapes
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        txt = txt & shp.TextFrame.TextRange.Text & vbCrLf
                    End If
                End If
            Next shp
            txt = txt & vbCrLf
        Next sld
    Else
        MsgBox "Selection is not slide-based. Please select slides from the left navigation pane.", vbExclamation
        Exit Sub
    End If

    If Len(Trim$(txt)) = 0 Then
        MsgBox "No text found in the selected slides.", vbInformation
        Exit Sub
    End If

    CopyToClipboard txt
    MsgBox "Plain text copied from the selected slides to the clipboard.", vbInformation
End Sub
