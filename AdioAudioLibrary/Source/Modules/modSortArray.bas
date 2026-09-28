Attribute VB_Name = "modSortArray"
Option Explicit
Const gcCallStackModName = "mSortArray" ' unique name different from functions elsewhere

'-------------------------------------------------------------------
'Sort a Two-Dimensional Array on Any Element
'
'Syntax:
'SortArray variablearray _
' , mDimensionToSort:=1, mPositionToSort:=1, mSecondaryPositionToSort:=2 _
' , mLo:=1, mHi:=100 _
' , isSortValue:=False
'
' (see acknowledgement.txt, section, sort array)
'-------------------------------------------------------------------

Public Function SortArray( _
   ByRef DArray _
 , Optional ByVal mDimensionToSort As Integer = 1 _
 , Optional ByVal mPositionToSort As Integer = 1 _
 , Optional ByVal mSecondaryPositionToSort As Integer = 0 _
 , Optional ByVal nLo As Long = -1, Optional ByVal nHi As Long = -1 _
 , Optional ByVal isSortValue As Boolean = False _
 , Optional ByVal mDimensionX As Integer = 0 _
 , Optional ByVal isDoEvents As Boolean = False _
 ) As Integer
    
    On Error GoTo ExitError
    
    '-------------------------------------------------------------------
    ' Alternative
    ' (see, RapidSort products,
    ' http://www.codebase.com/products/sorting
    ' (see, Opus6 Namespace,
    ' http://www.brpreiss.com/books/opus6/docs/Opus6.html
    
    ' Sorting guarantees that equal elements in the original
    ' sequence will retain their relative orderings in the
    ' final result.
    
    ' WARNING,
    ' Extra storage space was needed for sorting to apply the
    ' tertiary sort on the original message or element number
    ' so that equal elements in the original sequence will
    ' retain their relative orderings in the final result.
    ' Moving large groups of messages required it. Storage space
    ' was added to the main array for simplicity already so that
    ' a second array would not have to be created just for this,
    ' especially since the array size is dynamic.
    
    ' WARNING,
    ' Many elements in array will be affected consequtively
    ' between the specified lower and upper bound whether
    ' empty or not. Keeping the elements empty, that were
    ' originally empty before the sort, will not be
    ' possible here.
    ' (see 1.00.552)
    ' (see 1.00.563)
    
    ' WARNING,
    ' Quick/Shell/Bubble combination sorting without
    ' insertion/shifting. May be other faster options?
    ' (see, Sorting Algorithms,
    ' http://www.cs.ubc.ca/spider/harrison/Java/sorting-demo.html
    
    ' (see 1.00.563)
    '-------------------------------------------------------------------
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    'If IsMissing(mDimensionToSort) = True Then mDimensionToSort = 1
    ' =1 , sort data as row in first dimension, column in second dimension
    ' and sort rows in column, mPositionToSort
    ' =2 , sort data as row in second dimension, column in second dimension
    ' and sort rows in column, mPositionToSort
    If mDimensionToSort = 0 Then mDimensionToSort = 1

    'If IsMissing(mPositionToSort) = True Then mPositionToSort = 1
    If mPositionToSort = 0 Then mPositionToSort = 1
    
    'If IsMissing(isSortValue) = True Then isSortValue = False
    ' False=literal string, True=value&string
    ' Values and strings converted to strings to use primary+secondary+tertiary
    '
    ' when false, sort as a literal strings
    ' e.g. sort results would become,
    ' "1act"
    ' "1test"
    ' "11"
    ' "2"
    ' "toast"
    '
    ' when true, sort as values and strings
    ' e.g. sort results would become
    ' "toast" = "      0" + "toast"
    ' "1act" = "      1" + "1act"
    ' "1test" = "      1" + "1test"
    ' "2" = "      2" + "2"
    ' "11" = "     11" + "11"
    '
    ' e.g. sort results would become
    ' 0 = "      0" + ""
    ' 1 = "      1" + ""
    ' 2 = "      2" + ""
    ' 11 = "     11" + ""
    ' 110 = "    110" + ""
    
    If nLo = -1 Then nLo = LBound(DArray, mDimensionToSort)
    If nHi = -1 Then nHi = UBound(DArray, mDimensionToSort)
    'If IsMissing(nLo) = True Then nLo = LBound(DArray, mDimensionToSort)
    'If IsMissing(nHi) = True Then nHi = UBound(DArray, mDimensionToSort)
    ' Option to only sort a portion in the middle of the array
    ' by choosing the LBound and UBound of the array to search.
    ' Best used for arrays that are designed to accumulate data and expire
    ' old data and will not be redimmed or resorted to keep the LBound at 0 or 1.
    
    'If IsMissing(mSecondaryPositionToSort) = True Then mSecondaryPositionToSort = 0 ' no sort
    If mSecondaryPositionToSort < 0 Then mSecondaryPositionToSort = 0 ' no sort
    ' Option to use a secondary column to include in the sort.
    ' Indicate the location of the secondary column, otherwise remains primary sort only.
    
    'If IsMissing(mDimensionX) = True Then mDimensionX = 0 ' unknown dimension
    'If mDimensionX = 0 Then ' indicates to determine later but is slower
    
    ' (Function ismissing() is fictitious and only shown for reference)
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    Dim nGap As Long
    Dim nGapOriginal As Long
    Dim mDoneflag As Integer
    Dim isSwap As Boolean
    Dim tempvariable As Variant ' unknown type
    'Dim SwapArray() ' no longer implemented
    Dim nIndex As Long
    Dim mACol As Integer
    Dim nCountSwap As Long
    Dim nR As Long
    Dim nC As Long
    Dim temp As Integer
    
    Dim mStateSortPercent As Integer
    Dim nCurrent As Long
    Dim nMax As Long
    Dim dPercentStart As Double
    Dim dPercentEnd As Double
    Dim mPercentDiff As Integer
    Dim mAccel As Integer
    Dim mLeft As Integer
    Dim mWidth As Integer
    Dim mWidthMax As Integer
    
    Dim mTertiaryPositionToSort As Integer
    Dim DArray2() As String ' only string formats in conversions
    Dim mCommonLen As Integer
    Dim mCommon2Len As Integer
    
    Dim mErrorCode As Integer
    
    Const MB_INTEGERUBOUND = 32767
    Const MB_LONGUBOUND = &H7FFFFFFF
    
    mErrorCode = 1 ' assume it will fail
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ' Verify if one or two dimension with OnError during test
    ' if not specified from the parameters
    
    ' WARNING, it is better to specify the dimension as a parameter
    ' if known because OnError is not efficient and will slow down the
    ' program especially in fast loops.
    ' (see bug 105)
    ''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo ErrDimension ' temporary error catch
    If mDimensionX = 0 Then mDimensionX = 3: temp = UBound(DArray, 3)
    On Error GoTo ErrDimension
    If mDimensionX = 0 Then mDimensionX = 2: temp = UBound(DArray, 2)
    On Error GoTo ErrDimension
    If mDimensionX = 0 Then mDimensionX = 1: temp = UBound(DArray, 1)
    On Error GoTo ExitError ' restore error handling
    Select Case mDimensionX
     Case 1
        If mDimensionToSort <> 1 Or mPositionToSort > 1 Then
            ' can not sort a one-dimensional array with other positions
            Err.Raise 1, , "PROGRAM ERROR 32854, invalid array dimensions"
        End If
     Case 2
        If mDimensionToSort > 2 Then
            ' can not sort a one-dimensional array with other positions
            Err.Raise 1, , "PROGRAM ERROR 32854, invalid array dimensions"
        End If
     Case Else
        ' not yet implemented
        ' to sort a three-dimensional array
        Err.Raise 1, , "PROGRAM ERROR 32854, invalid array dimensions"
    End Select
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    Select Case mDimensionToSort
     Case 1
        ' E.g. sort (0) (1) (2) (3), first dimension
        ' E.g. sort (0,2) (1,2) (2,2) (3,2), first dimension
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Prepare to include additional column
        ' as a tertiary position to sort
        ' which represents the message or element number.
        mTertiaryPositionToSort = 1 ' in separate array
        ReDim DArray2(nLo To nHi, 1) ' prepare new column for tertiary position to sort
        For nIndex = LBound(DArray2, 1) To UBound(DArray2, 1)
            DArray2(nIndex, mTertiaryPositionToSort) = Right$(Space$(16) & Str$(nIndex), 16)
            'DArray2(nIndex, mTertiaryPositionToSort) = nIndex ' converts to string anyway
        Next nIndex
    
        Select Case mDimensionX
         Case 1
            ' Repeat checking with shorter gaps between two messages
            ' until the the last loop narrows the gap to one.
            ' The last loop is the same as using bubble sort
            ' which compares between two messages next to each other
            ' and luckily most messages are already sorted.
            nGap = Int((nLo + nHi - 1) / 2) ' Gap is half the records
            'nGap = Int(nHi / 2) ' not applicable anymore
            nGapOriginal = nGap
            Do While nGap > 0
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                
                ' Release resources
                ' (too slow, but may be useful)
                If isDoEvents = True Then DoEvents: If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                    
                ' Track status, scale 1 to 100%.
                ' However, only shows 1 to 50%
                ' which represents gaps greater than one.
                ' When gap is one, then no progress can
                ' be updated and have to assume it
                ' is between 50% to 100% meanwhile.
#If 1 = 0 Then ' comment out to enable test
                '{
                    mStateSortPercent = 0

                    ' Get max range of messages that would be processed.
                    ' Move range so progress bar is left justified.
                    nCurrent = nGap
                    'nCurrent = (nGapOriginal -nGap) ' not yet, invert, 0 to max
                    nMax = nGapOriginal
                    mWidthMax = 50 ' represents 50 to 100% percent

                    ' Convert results more linear or inverse logarithmic.
                    If nGap >= 5 Then
                        ' Move range to start
                        ' so progress bar is left justified
                        dPercentStart = 0
                        dPercentEnd = (CDbl(nCurrent) / CDbl(nMax)) * 100#
                        mPercentDiff = 0 ' calculate after any adjustments

                        If dPercentEnd < dPercentStart Then
                            ' start has to less or equal to end
                            ' usually occurs when array is almost empty where start points to
                            ' element after end to indicate it is empty
                            dPercentEnd = dPercentStart
                        End If

                        ' Show meter.
                        ' Indicating arrays are full when meter reaches all the way across textbox.
                        ' More sensitive initially and slows back down to normal at max.
                        ' Scale is logarithmic, e.g. 50% is actually ~40% capacity.
                        ' - accel=from 0 to 1000%, where 0 is no acceleration, 200% is twice as fast
                        ' - accel=0, normal and formula cancels out
                        ' - accel=50%, scale is exagerated somewhat
                        ' - accel=100%, scale is exagerated more
                        ' - accel=200%, scale is exagerated more and reach maximum twice as fast
                        ' (see notes - progressbar spreadsheet)
                        '{
                            mAccel = 5000 ' ' exagerate if array clears to zero often or too fast
                            'mAccel = 500 ' not exagerate much if array gets bigger normally
                            'mAccel = 0 ' not exagerate at all if gets very big and moves slow anyway
                            If dPercentEnd = dPercentStart Then
                                ' not sensitive because no difference
                            Else
                                dPercentEnd = dPercentEnd + dPercentEnd * (1 - dPercentEnd / 100) * (CDbl(mAccel) / 100)
                                If dPercentEnd > 100 Then dPercentEnd = 100
                            End If
                            mPercentDiff = Int(dPercentEnd - dPercentStart)
                        '}

                        mLeft = 0
                        mWidth = CDbl(mWidthMax) * dPercentEnd / 100
                        If mWidth < 1 Then
                            ' So it at least can see some of it if still processing
                            If nCurrent > 0 Then
                                mPercentDiff = 1 ' at least one, nonzero
                                mWidth = 1
                            End If
                        End If
                        mStateSortPercent = (mWidthMax - mWidth) ' invert, 0 to max
                        'mStateSortPercent = mWidth

                    Else
                        ' gaps are difficult to predict now since
                        ' excessive unknown number of loops
                        ' gap was 1-based scale, nonzero
                        mStateSortPercent = mWidthMax ' fixed for rest of scans
                    End If
                    ' ... = mStateSortPercent, global variable rest of program can monitor
                    Debug.Print mStateSortPercent, Rnd(1) ' see results
#End If
                
                    ' Alternative,
                    ' Results as is, logarithmic.
#If 1 = 0 Then ' comment out to enable test
                    mStateSortPercent = 0
                    If nGap >= 5 Then
                        mStateSortPercent = ((nGapOriginal - nGap) * mWidthMax) / nGapOriginal ' still small so not overflow
                        'mStateSortPercent = (nGapOriginal - nGap) * (CDbl(mWidthMax) / CDbl(nGapOriginal)) ' double too slow
                        If mStateSortPercent <= 0 Then mStateSortPercent = 1 ' nonzero to show started, and prevent overflow
                        If mStateSortPercent > mWidthMax Then mStateSortPercent = mWidthMax ' prevent overflow
    
                        ' WARNING,
                        ' Logarthmic progress, not linear.
                        ' Worse since slowing down especially when
                        ' reach 49% which is annoying. There is no other
                        ' reference to help compare with total loop
                        ' iterations or total messages scanned.
                    Else
                        ' gaps are difficult to predict now since
                        ' excessive unknown number of loops
                        ' gap was 1-based scale, nonzero
                        mStateSortPercent = mWidthMax ' fixed for rest of scans
                    End If
                    ' ... = mStateSortPercent, global variable rest of program can monitor
                    Debug.Print mStateSortPercent, Rnd(1) ' see results
#End If
                '}

                ' Repeat checking two messages at a distance of nGap from each other
                mDoneflag = 0
                Do ' alternative, Do While (mDoneflag <> 1)
                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown ' (too slow)
                    
                    ' WARNING,
                    ' Loop repeats unknown number of times.
                    ' Repeats until no more messages swap.
                    ' When the gap is one, then a typical
                    ' bubble sort occurs. Other gaps speed
                    ' up bubble sort.
                    
                    ' Check all messages
                    mDoneflag = 1
                    For nIndex = nLo To (nHi - nGap)
                        If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown ' (too slow)
                        
                        isSwap = False
                        ' Fixed width to maximum for comparisons
                        mCommonLen = -1 ' assume it will fail
                        If Len(DArray(nIndex)) >= Len(DArray(nIndex + nGap)) Then mCommonLen = Len(DArray(nIndex))
                        If Len(DArray(nIndex + nGap)) >= Len(DArray(nIndex)) Then mCommonLen = Len(DArray(nIndex + nGap))
                        
                        Select Case isSortValue
                         Case False
                            ' Data as string, compare 1st 1/2 to 2nd 1/2
                            ' (not optimized for integers (since using strings))
                            ' WARNING,
                            ' String manipulation in condition below slows down the sort.
                            ' Especially since it sorts when all elements are equal too.
                            If Left$(DArray(nIndex) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex, mTertiaryPositionToSort)) _
                             > Left$(DArray(nIndex + nGap) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex + nGap, mTertiaryPositionToSort)) _
                             Then
                                ' Swap if 1st > 2nd
                                tempvariable = DArray(nIndex)
                                DArray(nIndex) = DArray(nIndex + nGap)
                                DArray(nIndex + nGap) = tempvariable
                                tempvariable = DArray2(nIndex, mTertiaryPositionToSort)
                                DArray2(nIndex, mTertiaryPositionToSort) = DArray2(nIndex + nGap, mTertiaryPositionToSort)
                                DArray2(nIndex + nGap, mTertiaryPositionToSort) = tempvariable
                                nCountSwap = nCountSwap + 1
                                mDoneflag = 0
                                isSwap = True
                            End If
                         
                         Case True
                            ' Data as value and string, format as string
                            ' (not optimized for integers (since using strings))
                            ' WARNING,
                            ' String manipulation in condition below slows down the sort.
                            ' Especially since it sorts when all elements are equal too.
                            If Right$(Space$(mCommonLen) & Str$(Val(DArray(nIndex))), mCommonLen) _
                             & Left$(DArray(nIndex) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex, mTertiaryPositionToSort)) _
                             > Right$(Space$(mCommonLen) & Str$(Val(DArray(nIndex + nGap))), mCommonLen) _
                             & Left$(DArray(nIndex + nGap) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex + nGap, mTertiaryPositionToSort)) _
                             Then
                                ' Swap if 1st > 2nd
                                tempvariable = DArray(nIndex)
                                DArray(nIndex) = DArray(nIndex + nGap)
                                DArray(nIndex + nGap) = tempvariable
                                tempvariable = DArray2(nIndex, mTertiaryPositionToSort)
                                DArray2(nIndex, mTertiaryPositionToSort) = DArray2(nIndex + nGap, mTertiaryPositionToSort)
                                DArray2(nIndex + nGap, mTertiaryPositionToSort) = tempvariable
                                nCountSwap = nCountSwap + 1
                                mDoneflag = 0
                                isSwap = True
                            End If
                        End Select  ' isSortValue
                        
                        ' WARNING,
                        ' The count can be ten times higher than
                        ' the total messages depending on the
                        ' how much out of order it is.
                        ' The number messages actually scanned
                        ' is even higher.
                        ' Most of it is when the gap is near one
                        ' which is where the inefficiency of a
                        ' regular bubble sort takes place.
                        ' E.g. 5,000 messages can sort 50,000, scan 500,000
                        ' E.g. 35,000 messages can sort 500,000, scan 5,000,000
                        
#If 1 = 0 Then ' comment out to enable test
                        ' Verify sort after each time.
                        ' Verify sort equal times by original element number.
                        If isSwap = True Then
                        Debug.Print "------------------------------------------"
                        Debug.Print "Swapped "; nIndex; " to "; nIndex + nGap; ", range "; nLo; " to "; nHi
                        For nR = nLo To nHi
                         Debug.Print nR; Space$(5);
                         Debug.Print DArray(nR); Space$(5);
                         Debug.Print IIf(nR = nIndex, "<-----------", ""); IIf(nR = nIndex + nGap, "<-----------", "");
                         Debug.Print DArray2(nR, mTertiaryPositionToSort);
                         Debug.Print
                        Next
                        Stop
                        End If
#End If
                        
                        If nIndex = MB_LONGUBOUND Then Exit For ' (see 1.00.605)
                    Next nIndex
                Loop Until mDoneflag = 1
                nGap = Int(nGap / 2)
            Loop ' nGap
            mErrorCode = 0 ' okay
         
         Case 2 ' mDimensionX
            nGap = Int((nLo + nHi - 1) / 2) ' Gap is half the records
            nGapOriginal = nGap
            Do While nGap > 0
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                If isDoEvents = True Then DoEvents: If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                
                mDoneflag = 0
                Do ' alternative, Do While (mDoneflag <> 1)
                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown ' (too slow)
                    
                    mDoneflag = 1
                    For nIndex = nLo To (nHi - nGap)
                        If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown ' (too slow)
                        
                        isSwap = False
                        ' Fixed width to maximum for comparisons
                        mCommonLen = -1: mCommon2Len = -1 ' assume it will fail
                        If Len(DArray(nIndex, mPositionToSort)) >= Len(DArray(nIndex + nGap, mPositionToSort)) Then mCommonLen = Len(DArray(nIndex, mPositionToSort))
                        If Len(DArray(nIndex + nGap, mPositionToSort)) >= Len(DArray(nIndex, mPositionToSort)) Then mCommonLen = Len(DArray(nIndex + nGap, mPositionToSort))
                        If Len(DArray(nIndex, mSecondaryPositionToSort)) >= Len(DArray(nIndex + nGap, mSecondaryPositionToSort)) Then mCommon2Len = Len(DArray(nIndex, mSecondaryPositionToSort))
                        If Len(DArray(nIndex + nGap, mSecondaryPositionToSort)) >= Len(DArray(nIndex, mSecondaryPositionToSort)) Then mCommon2Len = Len(DArray(nIndex + nGap, mSecondaryPositionToSort))
                        
                        Select Case isSortValue
                         Case False
                            ' Data as string, compare 1st 1/2 to 2nd 1/2
                            ' (not optimized for integers (since using strings))
                            ' WARNING,
                            ' String manipulation in condition below slows down the sort.
                            ' Especially since it sorts when all elements are equal too.
                            If Left$(DArray(nIndex, mPositionToSort) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(nIndex, mSecondaryPositionToSort) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex, mTertiaryPositionToSort)) _
                             > Left$(DArray(nIndex + nGap, mPositionToSort) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(nIndex + nGap, mSecondaryPositionToSort) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex + nGap, mTertiaryPositionToSort)) _
                             Then
                            'If DArray(nIndex, mPositionToSort) > DArray(nIndex + nGap, mPositionToSort) Then ' not applicable
                                For mACol = nLo To (UBound(DArray, 2)) ' Move all related data together to temporary storage
                                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                                    ' Swap if 1st > 2nd
                                    tempvariable = DArray(nIndex, mACol)
                                    DArray(nIndex, mACol) = DArray(nIndex + nGap, mACol)
                                    DArray(nIndex + nGap, mACol) = tempvariable
                                Next mACol
                                tempvariable = DArray2(nIndex, mTertiaryPositionToSort)
                                DArray2(nIndex, mTertiaryPositionToSort) = DArray2(nIndex + nGap, mTertiaryPositionToSort)
                                DArray2(nIndex + nGap, mTertiaryPositionToSort) = tempvariable
                                nCountSwap = nCountSwap + 1
                                mDoneflag = 0
                                isSwap = True
                            End If
                         
                         Case True
                            ' Data as value and string, format as string
                            ' (not optimized for integers (since using strings))
                            ' WARNING,
                            ' String manipulation in condition below slows down the sort.
                            ' Especially since it sorts when all elements are equal too.
                            If Right$(Space$(mCommonLen) & Str$(Val(DArray(nIndex, mPositionToSort))), mCommonLen) _
                             & Left$(DArray(nIndex, mPositionToSort) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Right$(Space$(mCommon2Len) & Str$(Val(DArray(nIndex, mSecondaryPositionToSort))), mCommon2Len)) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(nIndex, mSecondaryPositionToSort) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex, mTertiaryPositionToSort)) _
                             > Right$(Space$(mCommonLen) & Str$(Val(DArray(nIndex + nGap, mPositionToSort))), mCommonLen) _
                             & Left$(DArray(nIndex + nGap, mPositionToSort) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Right$(Space$(mCommon2Len) & Str$(Val(DArray(nIndex + nGap, mSecondaryPositionToSort))), mCommon2Len)) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(nIndex + nGap, mSecondaryPositionToSort) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(nIndex + nGap, mTertiaryPositionToSort)) _
                             Then
                            'If DArray(nIndex, mPositionToSort) > DArray(nIndex + nGap, mPositionToSort) Then ' not applicable
                                For mACol = nLo To (UBound(DArray, 2)) ' Move all related data together to temporary storage
                                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                                    ' Swap if 1st > 2nd
                                    tempvariable = DArray(nIndex, mACol)
                                    DArray(nIndex, mACol) = DArray(nIndex + nGap, mACol)
                                    DArray(nIndex + nGap, mACol) = tempvariable
                                Next mACol
                                tempvariable = DArray2(nIndex, mTertiaryPositionToSort)
                                DArray2(nIndex, mTertiaryPositionToSort) = DArray2(nIndex + nGap, mTertiaryPositionToSort)
                                DArray2(nIndex + nGap, mTertiaryPositionToSort) = tempvariable
                                nCountSwap = nCountSwap + 1
                                mDoneflag = 0
                                isSwap = True
                            End If
                        End Select  ' isSortValue

#If 1 = 0 Then ' comment out to enable test
                        ' Verify sort after each time.
                        ' Verify sort equal times by original element number.
                        If isSwap = True Then ' by test primary, secondary
                        Debug.Print "------------------------------------------"
                        Debug.Print "Swapped "; nIndex; " to "; nIndex + nGap; ", range "; nLo; " to "; nHi
                        For nR = nLo To nHi
                         Debug.Print nR; Space$(5);
                         For nC = 1 To UBound(DArray, 2)
                            Debug.Print DArray(nR, nC) & Space$(5);
                         Next
                         Debug.Print DArray2(nR, 1);
                         Debug.Print IIf(nR = nIndex, "<-----------", ""); IIf(nR = nIndex + nGap, "<-----------", "");
                         Debug.Print
                        Next
                        Stop
                        End If
#End If
                        
                        If nIndex = MB_LONGUBOUND Then Exit For ' (see 1.00.605)
                    Next nIndex
                Loop Until mDoneflag = 1
                nGap = Int(nGap / 2)
            Loop ' nGap
            mErrorCode = 0 ' okay
        End Select
     
     Case 2 ' mDimensionToSort
        ' E.g. sort (2,0) (2,1) (2,2) (2,3), second dimension
        
        ''''''''''''''''''''''''''''''''''''''''''''''
        ' Prepare to include additional column
        ' as a tertiary position to sort
        ' which represents the message or element number.
        mTertiaryPositionToSort = 1 ' in separate array
        ReDim DArray2(1, nLo To nHi) ' prepare new column for tertiary position to sort
        For nIndex = LBound(DArray2, 2) To UBound(DArray2, 2)
            DArray2(mTertiaryPositionToSort, nIndex) = Right$(Space$(16) & Str$(nIndex), 16)
            'DArray2(mTertiaryPositionToSort, nIndex) = nIndex ' converts to string anyway
        Next nIndex
        
        Select Case mDimensionX
         Case 1
            ' not possible
        
         Case 2 ' mDimensionX
            nGap = Int((nLo + nHi - 1) / 2) ' Gap is half the records
            nGapOriginal = nGap
            Do While nGap > 0
                If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                If isDoEvents = True Then DoEvents: If gisEnd = True Then GoTo ExitEnd ' prevent multithreading issues caused by doevents
                
                mDoneflag = 0
                Do ' alternative, Do While (mDoneflag <> 1)
                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown ' (too slow)
                    
                    mDoneflag = 1
                    For nIndex = nLo To (nHi - nGap)
                        If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown ' (too slow)
                        
                        isSwap = False
                        ' Fixed width to maximum for comparisons
                        mCommonLen = -1: mCommon2Len = -1 ' assume it will fail
                        If Len(DArray(mPositionToSort, nIndex)) >= Len(DArray(mPositionToSort, nIndex + nGap)) Then mCommonLen = Len(DArray(mPositionToSort, nIndex))
                        If Len(DArray(mPositionToSort, nIndex + nGap)) >= Len(DArray(mPositionToSort, nIndex)) Then mCommonLen = Len(DArray(mPositionToSort, nIndex + nGap))
                        If Len(DArray(mSecondaryPositionToSort, nIndex)) >= Len(DArray(mSecondaryPositionToSort, nIndex + nGap)) Then mCommon2Len = Len(DArray(mSecondaryPositionToSort, nIndex))
                        If Len(DArray(mSecondaryPositionToSort, nIndex + nGap)) >= Len(DArray(mSecondaryPositionToSort, nIndex)) Then mCommon2Len = Len(DArray(mSecondaryPositionToSort, nIndex + nGap))
                        
                        Select Case isSortValue
                         Case False
                            ' Data as string, compare 1st 1/2 to 2nd 1/2
                            ' (not optimized for integers (since using strings))
                            ' WARNING,
                            ' String manipulation in condition below slows down the sort.
                            ' Especially since it sorts when all elements are equal too.
                            If Left$(DArray(mPositionToSort, nIndex) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(mSecondaryPositionToSort, nIndex) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(mTertiaryPositionToSort, nIndex)) _
                             > Left$(DArray(mPositionToSort, nIndex + nGap) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(mSecondaryPositionToSort, nIndex + nGap) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(mTertiaryPositionToSort, nIndex + nGap)) _
                             Then
                            'If DArray(mPositionToSort, nIndex) > DArray(mPositionToSort, nIndex + nGap) Then
                                For mACol = nLo To (UBound(DArray, 1)) ' Move all related data together to temporary storage
                                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                                    ' Swap if 1st > 2nd
                                    tempvariable = DArray(mACol, nIndex)
                                    DArray(mACol, nIndex) = DArray(mACol, nIndex + nGap)
                                    DArray(mACol, nIndex + nGap) = tempvariable
                                Next mACol
                                tempvariable = DArray2(nIndex, mTertiaryPositionToSort)
                                DArray2(nIndex, mTertiaryPositionToSort) = DArray2(nIndex + nGap, mTertiaryPositionToSort)
                                DArray2(nIndex + nGap, mTertiaryPositionToSort) = tempvariable
                                nCountSwap = nCountSwap + 1
                                mDoneflag = 0
                                isSwap = True
                            End If
                         
                         Case True
                            ' Data as value and string, format as string
                            ' (not optimized for integers (since using strings))
                            ' WARNING,
                            ' String manipulation in condition below slows down the sort.
                            ' Especially since it sorts when all elements are equal too.
                            If Right$(Space$(mCommonLen) & Str$(Val(DArray(mPositionToSort, nIndex))), mCommonLen) _
                             & Left$(DArray(mPositionToSort, nIndex) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Right$(Space$(mCommon2Len) & Str$(Val(DArray(mSecondaryPositionToSort, nIndex))), mCommon2Len)) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(mSecondaryPositionToSort, nIndex) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(mTertiaryPositionToSort, nIndex)) _
                             > Right$(Space$(mCommonLen) & Str$(Val(DArray(mPositionToSort, nIndex + nGap))), mCommonLen) _
                             & Left$(DArray(mPositionToSort, nIndex + nGap) & Space$(mCommonLen), mCommonLen) _
                             & IIf(mSecondaryPositionToSort = -1, "", Right$(Space$(mCommonLen) & Str$(Val(DArray(mSecondaryPositionToSort, nIndex + nGap))), mCommon2Len)) _
                             & IIf(mSecondaryPositionToSort = -1, "", Left$(DArray(mSecondaryPositionToSort, nIndex + nGap) & Space$(mCommon2Len), mCommon2Len)) _
                             & IIf(mTertiaryPositionToSort = -1, "", DArray2(mTertiaryPositionToSort, nIndex + nGap)) _
                             Then
                            'If DArray(mPositionToSort, nIndex) > DArray(mPositionToSort, nIndex + nGap) Then ' not applicable
                                For mACol = nLo To (UBound(DArray, 1)) ' Move all related data together to temporary storage
                                    If gisEnd = True Then GoTo ExitEnd ' not needed at shutdown
                                    ' Swap if 1st > 2nd
                                    tempvariable = DArray(mACol, nIndex)
                                    DArray(mACol, nIndex) = DArray(mACol, nIndex + nGap)
                                    DArray(mACol, nIndex + nGap) = tempvariable
                                Next mACol
                                tempvariable = DArray2(mTertiaryPositionToSort, nIndex)
                                DArray2(mTertiaryPositionToSort, nIndex) = DArray2(mTertiaryPositionToSort, nIndex + nGap)
                                DArray2(mTertiaryPositionToSort, nIndex + nGap) = tempvariable
                                nCountSwap = nCountSwap + 1
                                mDoneflag = 0
                                isSwap = True
                            End If
                        End Select  ' isSortValue
                        
#If 1 = 0 Then ' comment out to enable test
                        ' Verify sort after each time.
                        ' Verify sort equal times by original element number.
                        If isSwap = True Then ' by test primary, secondary
                        Debug.Print "------------------------------------------"
                        Debug.Print "Swapped "; nIndex; " to "; nIndex + nGap; ", range "; nLo; " to "; nHi
                        For nR = nLo To nHi
                         Debug.Print nR; Space$(5);
                         For nC = 1 To UBound(DArray, 1)
                            Debug.Print DArray(nC, nR) & Space$(5);
                         Next
                         Debug.Print DArray2(1, nR);
                         Debug.Print IIf(nR = nIndex, "<-----------", ""); IIf(nR = nIndex + nGap, "<-----------", "");
                         Debug.Print
                        Next
                        Stop
                        End If
#End If
                        
                        If nIndex = MB_LONGUBOUND Then Exit For ' (see 1.00.605)
                    Next nIndex
                Loop Until mDoneflag = 1
                nGap = Int(nGap / 2)
            Loop ' nGap
            mErrorCode = 0 ' okay
        End Select
    
    End Select
    
    ''''''''''''''''''''''''''''''''''''''''''''''
#If 1 = 0 Then ' comment out to enable test
    'Debug.Print "PROGRAM WARNING 3598, executing tested code." ' already using debug.print
    
    ' Show results
    For nR = nLo To nHi
     Debug.Print nR;
     For nC = 1 To UBound(DArray, 2)
        Debug.Print DArray(nR, nC) & Space$(5);
     Next nC
     Debug.Print DArray2(nR, 1);
     Debug.Print
    Next nR
    For nR = nLo To nHi
     Debug.Print nR;
     For nC = 1 To UBound(DArray, 1)
        Debug.Print DArray(nC, nR) & Space$(5);
     Next
     Debug.Print DArray2(1, nR);
     Debug.Print
    Next
#End If

    SortArray = mErrorCode

    Exit Function
ErrDimension:
    mDimensionX = 0
    ' Note that the temp variable will store inaccurate information when
    ' crash occurs which is okay since discarded. E.g. error message,
    ' "float inexact result", would be reported by some debuggers.
    Resume Next

    Exit Function
ExitEnd: ' prevent multithreading issues caused by doevents or background processes
    
    Exit Function
ExitError: ' help common error handler, if any
    Err.Raise Err.number, , Err.Description
    Resume
End Function

