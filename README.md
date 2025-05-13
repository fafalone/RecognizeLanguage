# RecognizeLanguage
A short demo of the Windows ELS API in twinBASIC 

 ![image](https://github.com/user-attachments/assets/93ab4456-6e1e-4f5c-90fd-90387a0a0828) ![image](https://github.com/user-attachments/assets/d905204c-f091-424a-b94c-70ad318120d1)
 

This is just a small snippet of basic language recognition with the built in Windows API `MappingRecognizeText`.

It returns codes like `en` for English and `fr` for French.

**Requirements**\
Windows 7+\
(To compile from source) twinBASIC, Package: "Windows Development Library for twinBASIC" (WinDevLib) v8.12.530+\
(Project->References->Available Packages tab)


```vba
 
    Private Function RecognizeLanguage(ByVal sText As String, pMatches As String()) As Long
        Dim EnumOptions As MAPPING_ENUM_OPTIONS
        Dim prgServices As LongPtr 'PMAPPING_SERVICE_INFO   
        Dim dwServicesCount As Long
        Dim hr As Long
        Dim gSvc As UUID
        
        gSvc = ELS_GUID_LANGUAGE_DETECTION
        EnumOptions.Size = LenB(EnumOptions)
        EnumOptions.pGuid = VarPtr(gSvc)
        
        hr = MappingGetServices(EnumOptions, prgServices, dwServicesCount)
        
        If SUCCEEDED(hr) Then
            Dim bag As MAPPING_PROPERTY_BAG
            Dim pService As MAPPING_SERVICE_INFO = CType(Of MAPPING_SERVICE_INFO)(prgServices)
            
            bag.Size = LenB(bag)
            hr = MappingRecognizeText(pService, sText, Len(sText), 0, vbNullPtr, bag)
            If SUCCEEDED(hr) Then
                Dim pRange As MAPPING_DATA_RANGE = CType(Of MAPPING_DATA_RANGE)(bag.prgResultRanges)
                Dim cch As LongPtr
                Dim offset As LongPtr
                Dim sRes As String, nRes As Long
                Do
                    cch = wcslen(pRange.pData + offset)
                    If cch = 0 Then Exit Do
                    sRes = LPWSTRtoStr(pRange.pData + offset, False)
                    ReDim Preserve pMatches(nRes)
                    pMatches(nRes) = sRes
                    nRes += 1
                    offset += cch * 2 + 2
                Loop
                MappingFreePropertyBag(bag)
            Else
                Debug.Print "MappingRecognizeText error 0x" & Hex$(hr) & ", " & GetSystemErrorString(hr)
            End If
            MappingFreeServices(prgServices)
        Else
            Debug.Print "MappingGetServices error 0x" & Hex$(hr) & ", " & GetSystemErrorString(hr)
        End If
        Return nRes
    End Function
   ```

Usage example:
```vba
  
    Private Sub Command1_Click() Handles Command1.Click
        Dim sLang() As String
        Dim sOut As String
        If RecognizeLanguage(Text1.Text, sLang) Then
            For i As Long = 0 To UBound(sLang)
                If i = 0 Then
                    sOut = "Best match: " & sLang(0)
                Else
                    sOut &= vbCrLf & "Other result " & i & ": " & sLang(i)
                End If
            Next
        End If
    End Sub
```
