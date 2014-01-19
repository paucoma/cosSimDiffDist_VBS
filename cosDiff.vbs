' \file 	cosDiff.vbs
' \author 	Pau Coma paucoma@paucoma.com
' \date:	2013-09-28	
' 
' \description : This is a class for finding String Similarities
'
' \usage :  Class needs to be instatiated as an object

'# Handy for Error Investigation #
'On Error Resume Next	'This will continue even if there is an error on code which is called after this statement
'	tmpStr=tmpStr*1 	'Code to investigate
'If Err.Number <> 0 Then
'  WScript.Echo Err.Description & " : " & myStrNum
'  Err.Clear
'End If
'On Error Goto 0	'This will end the Error continue block of code
' Answer found in stackoverflow by Dylan Beattie at http://stackoverflow.com/questions/157747/vbscript-using-error-handling
' VBScript has no notion of throwing or catching exceptions, but the runtime provides a global Err object that contains the reuslts of the last operation performed. You have to explicitly check whether the Err.Number property is non-zero after each operation.
' VBScript doesn't support On Error Goto [label], only VB and VBA do.
' For a list of Error Code see : [http://support.microsoft.com/kb/146864 Error Trapping with Visual Basic for Applications, Article ID: 146864 ] or [http://msdn.microsoft.com/en-us/library/ms234761(v=vs.90).aspx Trappable Errors in Visual Basic]
' Examples on How to Use Error can be found [http://support.microsoft.com/kb/141571/EN-US How to Use "On Error" to Handle Errors in a Macro, Article ID: 141571]

Option Explicit

'! String Character Frequency Counter Subroutine 
'!      : 'Split does not accept the "" as parameter to split a string to characters
'!          so this is the work around helper funciton
'! 
'! @return arrChars: Array of characters derived from the string
'! @param myStr String to be split
Function getChars(myStr)
    Dim i, myLen
    Dim arrChars
    
    myLen = Len(myStr)-1
    ReDim arrChars(myLen)

    For i = 0 to myLen
        arrChars(i) = Mid(myStr, i + 1,1)
    Next
    getChars = arrChars
End Function
'! String Character Frequency Counter Subroutine 
'!      : Fills dictionary with characters and their occurance frequency within the string
'! 
'! @return myDict modified as it is passed by Reference
'! @param myDict Dictionary to be filled: Keys = Characters, Items = # of occurances
'! @param myStr String to be analyzed
Sub DefStrFreq(ByRef myDict,myStr)
    Dim strList
    Dim myChar
    'myDict.RemoveAll
    strList = getChars(myStr)
    'Split does not accept the "" as parameter so we make a helper funciton
    For Each myChar in strList
        If myDict.Exists(myChar) Then
            myDict.Item(myChar) = _
                    myDict.Item(myChar) + 1
        Else
            myDict.Add myChar, 1
        End If
    Next
End Sub
'! String Character Frequency Counter Subroutine 
'!      : Fills dictionary with characters occurance frequencies 
'!      , each element(Column) is a character in one of both Arrays
'! 
'! @return myArr 2 dimensional Array (2, NumCharacters)
'!     Union of Charcters | --> Not Represented as not important
'!          Frequency in A| | |
'!          Frequency in B| | |
'! @param myDictA Dictionary to be Analyzed
'! @param myDictB Dictionary to be Analyzed
Function VectorUnionMatrix(ByRef dictA, ByRef dictB)
    Dim myArr
    Dim myIndex
    Dim myElem
    'Array is defined backwards because Preserve
    '   can only preserve while changing the last dimension.
    ReDim myArr(1,dictA.Count + dictB.Count - 1)
    myIndex = -1
    For Each myElem in dictA.Keys
        myIndex = myIndex + 1
        myArr(0,myIndex) = dictA.Item(myElem)
        If dictB.Exists(myElem) Then
            myArr(1,myIndex) = dictB.Item(myElem)
        Else
            myArr(1,myIndex) = 0
        End If
    Next
    For Each myElem in dictB.Keys
        If Not dictA.Exists(myElem) Then
            myIndex = myIndex + 1
            myArr(0,myIndex)=0
            myArr(1,myIndex)=dictB.Item(myElem)
        End If
    Next
    ReDim Preserve myArr(1,myIndex)
    ' We return a two dimensional array
    VectorUnionMatrix = myArr
End Function
'! Dot Product calculator of specific Matrix output from VectorUnionMatrix
'!
'! @param myArr 2 dimensional Array (2, NumCharacters)
'!     Union of Charcters | --> Not Represented as not important
'!          Frequency in A| | |
'!          Frequency in B| | | 
'! @return mySum : the result of the dot Product
Function DotProduct(myArr)
    Dim mySum, i
    mySum = 0
    For i = 0 to UBound(myArr,2)
        mySum = mySum + myArr(0,i)*myArr(1,i)
    Next
    DotProduct = mySum
End Function
'! Vector Magnitude calculator: self explanitory...
'!
'! @param myArr : one dimensional array representing a vector
'! @return vecMagnitude : the vectors magnitude
Function VectorMagnitude(myArr)
    Dim mySum, i,j
    Dim myElem
    mySum = 0
    For Each myElem in myArr
        mySum = mySum + myElem*myElem
    Next
    VectorMagnitude = Sqr(mySum)
End Function
'! Cosine Similarity Calculator
'!
'! @param dictA : Dictionary with Frequency Occurance of String A characters
'! @param dictB : Dictionary with Frequency Occurance of String B characters
'! @return similarityCoefficient : Value between 0 and 1 defining Similarity
Function stepCosDist(ByRef dictA,ByRef dictB)
    Dim myArr
    If dictA.Count = 0 Or dictB.Count = 0 Then
        stepCosDist = 0
    Else
        myArr = VectorUnionMatrix(dictA, dictB)
        stepCosDist = DotProduct(myArr) / _
                    (VectorMagnitude(dictA.Items) * _
                    VectorMagnitude(dictB.Items))
    End If
End Function
'! Facilitator Function: 
'!          Creates the necessary environment variables for the 
'!          Cosine Similarity calculator to operate.      
'!
'! @param myStrA : String for comparison
'! @param myStrB : String for comparison
'! @return cosineSimilarity : @see Cosine Similarity Calculator
Function getSimilarity(myStrA,myStrB)
    Dim myOut
    Dim myDictA
    Set myDictA = CreateObject("Scripting.Dictionary")
    Call DefStrFreq(myDictA, myStrA)
    Dim myDictB
    Set myDictB = CreateObject("Scripting.Dictionary")
    Call DefStrFreq(myDictB, myStrB)
    myOut = stepCosDist(myDictA, myDictB)
    Set myDictA = Nothing
    Set myDictB = Nothing
    getSimilarity = myOut
End Function
'! Interface Facilitator Function: prompts the user for input strings
'!
'! @param myStr : Additional string to differenciate multiple prompt boxes
'! @return editBoxContents : What the edit box contains when the user presses OK
Function getInput(myStr)
    'InputBox(prompt[, title][, default][, xpos][, ypos][, helpfile, context])
    getInput = InputBox("Input Comparison String " & myStr, "Cosine Similarity") 
End Function

MsgBox getSimilarity(getInput("A"),getInput("B"))