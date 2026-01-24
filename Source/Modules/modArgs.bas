Attribute VB_Name = "modArgs"
Option Explicit

Public argc&, argv() As String
Public Function argProcessCMDLine(Optional ByVal Args As String)
Dim I&

' This is now just a wrapper for GetArgs.  Call GetArgs to do the
' processing.
argGetArgs argv, argc, Args

' All done!  Now you can access the command-line arguments from
' argv() and argc

End Function
Public Function argGetArgs(ByRef argv() As String, ByRef argc As Long, _
 Optional ByVal Args As String)
Dim I&

' This is the temporary variable (duh).  We keep the 'processed'
' (ie mutilated) version of the command line in here.
Dim strArgTemp$

' This is used to store character positions gleaned from InStr() calls
Dim lngCharPos&

' Do we need to pull the arguments from Command$?
If Args <> "" Then
  strArgTemp = Trim$(Args)
Else
  strArgTemp = Trim$(command$)
End If

' Do we want to set the first argument to the EXE path?
If Args = "" Then
  
  ' Resize the array
  ReDim argv(0&)
  argc = 1&
  
  ' Save the value
  argv(0&) = App.Path
  If Right$(argv(0&), 1&) <> "\" Then argv(0&) = argv(0&) & "\"
  argv(0&) = argv(0&) & App.exeName & ".exe"
  
Else
  
  ' Nope.  Set argc to 0 so it works good-like
  argc = 0&
  
End If

'---------------------------------------------------------------------------
' Right, here's the main loop.  What we do, is every time we find an
' argument, we strip it from strArgTemp.  Ergo, when all arguments have been
' processed, the string is empty.  Simple, huh? :P
Do Until strArgTemp = ""
  
  ' First, we check to see if we're dealing with a quoted argument
  ' (ie: "this has three spaces!")
  If Left$(strArgTemp, 1&) = Chr$(34) Then
    
    ' Yup; increase the array by one
    argc = argc + 1&
    ReDim Preserve argv(argc - 1&)
    
    ' Find the ending quote
    lngCharPos = InStr(2&, strArgTemp, Chr$(34&))
    
    ' IS there an ending quote?  If not, use the rest of the string
    ' (The +2 is there to negate the -2 below which is designed to
    ' avoid catching that last quote... which we aren't worried
    ' about... ;)
    If lngCharPos = 0& Then lngCharPos = Len(strArgTemp) + 2&
    
    ' Strip out the argument
    argv(argc - 1&) = Mid$(strArgTemp, 2&, lngCharPos - 2&)
    
    ' Now remove that argument from the temp var
    strArgTemp = LTrim$(Mid$(strArgTemp, lngCharPos + 1&))
  Else
    
    ' No quotes; expand array
    argc = argc + 1&
    ReDim Preserve argv(argc - 1&)
    
    ' Now, are there actually any more spaces?
    If InStr(1, strArgTemp, " ") <> 0& Then
      
      ' Yes.  But first, check to see if there's a quote in this
      ' argument
      If InStr(1, strArgTemp, Chr$(34)) <> 0 And _
       InStr(1, strArgTemp, Chr$(34)) < InStr(1, strArgTemp, " ") Then
        
        ' Yes.  First, extract up to the first quote
        lngCharPos = InStr(1&, strArgTemp, Chr$(34))
        argv(argc - 1&) = Left$(strArgTemp, lngCharPos - 1&) & Chr$(34)
        strArgTemp = Mid$(strArgTemp, lngCharPos + 1&)
        
        ' Next, find the closing quote
        lngCharPos = InStr(1&, strArgTemp, Chr$(34))
        
        ' Does it exist?
        If lngCharPos <> 0& Then
          
          ' Yes, extract up till that point
          argv(argc - 1&) = argv(argc - 1&) & Left$(strArgTemp, lngCharPos - 1&) & Chr$(34)
          strArgTemp = Mid$(strArgTemp, lngCharPos + 1&)
          
        Else
          
          ' No... just extract the rest of the string
          argv(argc - 1&) = strArgTemp
          strArgTemp = ""
          
        End If
        
      Else
        
        ' Nope.  Just find and extract up till the next space
        lngCharPos = InStr(1&, strArgTemp, " ")
        
        ' Now strip out the argument, and remove it from strArgTemp
        argv(argc - 1&) = Left$(strArgTemp, lngCharPos - 1&)
        strArgTemp = Mid$(strArgTemp, lngCharPos + 1&)
        
      End If
      
    Else
      
      ' Nope.  The rest of the string IS the last argument
      argv(argc - 1&) = strArgTemp
      strArgTemp = ""
      
    End If
  End If
  
  ' Trim the command line
  strArgTemp = Trim$(strArgTemp)
  
Loop
End Function

'===========================================================================
' argSwitchPresent()
'   This is a little something I wrote because I am (like most programmers)
'   oh so very lazy.  Just feed it a switch (like /l, /W), and it will look
'   for it in your argv() array.  If it finds it, it returns True (guess
'   what it returns if it DOESN'T find it.... FALSE!  Bet you didn't expect
'   that :p).  What's more, it also returns the array index of that switch
'   (for extra processing) in Position.
'   If that wasn't enough, it also supports pattern matching, so you can
'   search for special switches like "/i:*".

Public Function argSwitchPresent(ByRef Switch As String, _
    Optional ByRef Position As Long = 0, _
    Optional ByVal UseWildcard As Boolean = False) As Boolean
Dim I&

' Do we want to use pattern matching?
If UseWildcard = True Then
  ' Yup; start searching
  For I = 0& To argc - 1&
    ' Compare using the Like operator
    If argv(I) Like Switch Then
      ' Return true, and the position
      argSwitchPresent = True
      Position = I
      Exit Function
    End If
  Next
Else
  ' Nup; start searching
  For I = 0& To argc - 1&
    ' Compare using the = operator (ohlike, wow...)
    If argv(I) = Switch Then
      ' Return true, and the position
      argSwitchPresent = True
      Position = I
      Exit Function
    End If
  Next
End If

' If it got here, it ain't there, so return false
argSwitchPresent = False

End Function

'===========================================================================
' argGetSwitchArg()
'  Returns the argument immediately after the specified switch.  Switch
'  finding is done in the same way that argSwitchPresent() does it.
Public Function argGetSwitchArg( _
  ByRef Switch As String, _
  Optional ByRef Position As Long = 0, _
  Optional ByVal UseWildcard As Boolean = False _
) As String
Dim I&

' Do we want to use pattern matching?
If UseWildcard = True Then
  ' Yup; start searching
  For I = 0& To argc - 1&
    ' Compare using the Like operator
    If argv(I) Like Switch Then
      ' Is there a next argument?
      If (I + 1&) < argc Then
        ' Yup, return it
        argGetSwitchArg = argv(I + 1&)
        Position = I + 1&
      Else
        ' Nope... return -1 and ""
        argGetSwitchArg = ""
        Position = -1&
      End If
      Exit Function
    End If
  Next
Else
  ' Nup; start searching
  For I = 0& To argc - 1&
    ' Compare using the = operator (ohlike, wow...)
    If argv(I) = Switch Then
      ' Is there a next argument?
      If (I + 1&) < argc Then
        ' Yup, return it
        argGetSwitchArg = argv(I + 1&)
        Position = I + 1&
      Else
        ' Nope... return -1 and ""
        argGetSwitchArg = ""
        Position = -1&
      End If
      Exit Function
    End If
  Next
End If

' If it got here, it ain't there, so return nothing
argGetSwitchArg = ""
Position = -1&

End Function

'===========================================================================
' argAdd()
'  This little puppy adds a new argument to the array.  Easy as.
Public Function argAdd(ByVal Argument As String)

' First, redimension the array
ReDim Preserve argv(argc)
argc = argc + 1&

' Now, append the argument
argv(argc - 1&) = Argument

' Done

End Function

'===========================================================================
' argRemove()
'  This method will remove a specified argument, and collapse the array.
Public Function argRemove(ByVal Index As Long)

' First up, do we need to redim the array, or erase it?
If argc = 1 Then
  
  Erase argv
  argc = 0&
  Exit Function
  
Else
  
  ' Loop through the elements, putting them back one index
  Dim I&
  For I = Index + 1& To argc - 1&
    argv(I - 1&) = argv(I)
  Next I
  
  ' Now, redim the array
  argc = argc - 1&
  ReDim Preserve argv(argc - 1&)
  
  ' Done
  Exit Function
  
End If

End Function

'===========================================================================
' argRebuildCmdLine()
'  Rebuilds the command line from the current array.
Public Function argRebuildCmdLine() As String

' Ok, here we are going to loop through argv[], appending the arguments to
' the string.
Dim m_strBuffer$, I&

If argc > 0& Then
  
  If InStr(argv(I), " ") > 0& Then
    m_strBuffer = Chr$(34) & argv(I) & Chr$(34)
  Else
    m_strBuffer = argv(I)
  End If
  
End If

For I = 1& To argc - 1&
  
  If InStr(argv(I), " ") <> 0& And InStr(argv(I), Chr$(34)) = 0& Then
    m_strBuffer = m_strBuffer & " " & Chr$(34) & argv(I) & Chr$(34)
  Else
    m_strBuffer = m_strBuffer & " " & argv(I)
  End If
  
Next I

' Return the command line
argRebuildCmdLine = m_strBuffer

End Function
