VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "stdRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Spec:
'This class has been designed to extend the feature set provided by VBScript.Regexp (or to eventually replace it entirely)
'Functions implemented on the class
'CONSTRUCTORS
'    [X] Create - With Cache
'    [X] init     #PROTECTED
'
'PROPERTIES
'    [X] Get/Let Pattern
'    [X] Get/Let Flags
'
'INSTANCE METHODS
'    [X] Test
'    [X] Match
'    [X] MatchAll
'    [X] Replace
'    [X] List
'    [X] ListArr
'    [ ] Compile - Can we compile for faster execution?
'    [?] NamedCaptures - See ruby
'    [?] Names - Get the names of the pattern
'
'REGEX ENGINE
'    [X] Implement Named Capturing Groups
'    [-] Flags Supported
'        [X] i - Case insensitive
'        [X] g - Global
'        [X] m - Multiline
'        [X] x - Ignore whitespace
'        [X] c - Ignore comment
'        [X] s - Singleline mode. `.` Matches all characters including whitepsace
'        [ ] n - Explicit capture. Do not capture unnamed groups
'    [ ] More symbols e.g. \A, \Z, ...
'    [ ] Rewrite the regex engine to implement more features, dependency reduction and Mac compatibility? https://swtch.com/~rsc/regexp/regexp1.html and https://deniskyashif.com/2019/02/17/implementing-a-regular-expression-engine/ 
'        [ ] Stream Support? - could be helpful with huge files
'    [ ] Back propogation
'    [ ] Positive and Negative lookahead/lookbehind
'
'OUT-OF-SCOPE
'    * Grep from file - unless there a smarter way than using rx.Match(stdShell.Read()) we won't be adding this to this class
'
'EXAMPLES
'# 1 - Printing codes from some haystack  
'
'   sHaystack = "12345-AA1, 12345-AB1,15233-AC3, 63234-ZD2"
'   set debugPrint = stdCallback.CreateFromModule("Main", "debugPrint")
'   stdEnumerator.Create(stdRegex.Create("(?<Site>\d{5})-(\w{2}\d)").MatchAll(sHaystack)).Map(stdLambda.Create("$1.item(""Site"")")).Unique.ForEach(debugPrint)
'
'   Sub debugPrint(ByVal s as string)
'       Debug.Print s
'   End Sub
'
'# 2 - Dumping array to excel:
'
'   sResult = "Here is some cool data:" & vbCrLf & _ 
'             "12345-STA1,10/02/2019,123" & vbCrLf & _ 
'             "12323-STB9,01/01/2005,2123" & vbCrLf & _ 
'             "and here is some more:" & vbCrLf & _ 
'             "23565-STC2,??/??/????,23" & vbCrLf & _ 
'             "62346-STZ9,01/05/1932,5" 
'   vResult = rx.ListArr(sHaystack, Array("$id-$date","$count"))
'   Range("A1:B1").value = Array("ID-Date","Count")
'   Range("A2").Resize(ubound(vResult,1),2).value = vResult


'Regex tokenise example
'   /ab[c-d](?<someName>efg)\s+
' -->
'   [a-a][b-b][c-d](?<someName>[e-e][f-f][g-g])[\t-\r ]
' -->
'   [
'     {type: 1, charRange: {min: "a", max:"a", negated: false}, capture:{}},
'     {type: 1, charRange: {min: "b", max:"b", negated: false}, capture:{}},
'     {type: 1, charRange: {min: "c", max:"d", negated: false}, capture:{}},
'     {type: 3, charRange: {}, capture:{name: "someName", index: ###}},
'     {type: 1, charRange: {min: "e", max:"e", negated: false}, capture:{}},
'     {type: 1, charRange: {min: "f", max:"f", negated: false}, capture:{}},
'     {type: 1, charRange: {min: "g", max:"g", negated: false}, capture:{}},
'     ...
'   ]

Private Enum ESymbolType
  iEmptyTransition = 0
  iCharacterRanges = 1
  iCaptureGroupCheck = 2 '/(a)(?0)/
  iCaptureStart = 3
  iCaptureEnd = 4
End Enum

Private Type CharacterRange
  CharMin as Integer  '2 Byte
  CharMax as Integer  '2 Byte
  Negated as boolean  '2 Byte
End Type

Private Type CaptureInfo
  CaptureName as string '4+1+n byte (pointer to string of n chars length. Ends in \0)
  iStart as long        '4 byte
End Type

Private Type Symbol
  iSymbolType as ESymbolType           '2 Byte
  CharacterRanges() as CharacterRange  '6*n Byte  where n>=1. Note uninitialised == 6 bytes
  Capture as CaptureInfo               '
End Type

'https://kean.blog/post/regex-parser
'https://kean.blog/post/regex-compiler
'https://kean.blog/post/regex-matcher
'https://deniskyashif.com/2019/02/17/implementing-a-regular-expression-engine/
Private Sub Compile(ByVal sPattern as string)
  Call parseTokenBlock(opStack, sPattern, )
End Sub


Private Sub parseTokenBlock(ByRef opStack, ByRef sPattern as string, ByRef iIndex as long)
    
End Sub









'Test if a character matches a symbol
'@param {Byte}   The character to be tested
'@param {Symbol} The symbol to be tested against
'@returns {Boolean} Whether the character matched or not
Private Function testSymbol(ByRef char as Byte, ByRef sym as Symbol) as Boolean
  select case sym.iSymbolType 
    case iCharacterRanges
      Dim b as boolean: b=false
      For i = 0 to ubound(sym.CharacterRanges)
        b = b or testCharRange(char,sym.CharacterRanges(i))
        if b then exit for
      next
      testSymbol = b
    case else
      '...
  end select
End Function

'Test that a character is in a given range
'@param {Byte} The character to be tested
'@param {CharacterRange} The character range to be tested against
'@returns {Boolean} Whether the character matched or not
Private Function testCharRange(ByRef char as Byte, ByRef Range as CharacterRange) as boolean
  With Range
    testCharRange = .Negated XOR .Min <= char AND char <= .max
  end with
End Function


'Creates a Symbol comprising of 2 character ranges resembling a whitespace character \s
'@returns {Symbol} A whitespace symbol
Private Function getWhiteSpaceSymbol() as Symbol
  getWhiteSpaceSymbol.iSymbolType = iCharacterRanges
  Redim getWhiteSpaceSymbol.CharacterRanges(0 to 1)
  getWhiteSpaceSymbol.CharacterRanges(0) = getCharRange(9,13)
  getWhiteSpaceSymbol.CharacterRanges(0) = getCharRange(32)
End Function

'Given a min and optional max and negated, creates a CharacterRange to match
'@param {Long} Min - The min (inclusive) ascii value to match.
'@param {Long} Max - The max (inclusive) ascii value to match. If max not provided, min is used as max.
'@param {Boolean} Negated - Negated character ranges will return false if theres a match, and true if there isn't.
'@return {CharacterRange} Character range struct for matching
Private Function getCharRange(ByVal Min as long, Optional ByVal Max as long = -1, Optional ByVal Negated as boolean = false) as CharacterRange
  if Max = -1 then Max = Min
  getCharRange.CharMin = Min
  getCharRange.ChatMax = Max
  getCharRange.Negated = Negated
End Function 







'MIGHT NOT NEED THESE as we use DFS instead of BFS
'==================================================
'    Private NFAMetadata
'      iStart as long
'      iEnd as long
'    End Type
'    Private Type NFAEdge
'      Rule as Symbol
'      iNext as long
'    End Type
'    Private Type NFANode
'      '
'      metadata as NFAMetadata
'    
'      '
'      edges() as NFAEdge
'    End type
'    Private Type NFAStatePath
'      'The node that we came from
'      node as long
'    
'      'Keep track of how often this path is referenced, for garbage collection?
'      dependencies as long
'    
'      'The index of the character that was consumed to get in this node
'      char as long
'      
'      'The inex within this array of the previous node in the path
'      prev as long
'    End Type
'    
'    Private NFANodes() as NFANode
'    
'    'All the nodes and paths to get there, that we are currently in (are currently active)
'    Private currentState() as long
'    
'    Private paths() as NFAStatePath
'==================================================












































































'Private p_initialised As Boolean
'Private p_pattern As String
'Private p_cPattern As String
'Private p_flags As String
'Private p_namesDict As Object
'Private p_regex As Object




'Creates a regex object given a pattern and flags.
'@constructor
'
'@param {string}  Pattern - The pattern to match
'@param {string}  Flags - Optional flags to apply
'@return {stdRegex} Regular expression object
'@example
'    stdRegex.Create("A\d+","i")
Public Function Create(ByVal pattern As String, Optional ByVal flags As String = "") As stdRegex
  If Not Me Is stdRegex Then
    Call stdError.Raise("Constructor called on object not class")
    Exit Function
  End If
  
  'Cache regex objects for faster speed
  static cache as object: if cache is nothing then set cache = CreateObject("Scripting.Dictionary")
  Dim sKey as string: sKey = pattern & "&" & flags
  if not cache.exists(sKey) then
    'Call to [friend] init requires type information, can't use dict directly
    Dim r As stdRegex
    Set r = New stdRegex
    Call r.init(pattern, flags)
    
    'Set cache
    Set cache(sKey) = r
  end if

  'Return cached object
  set Create = cache(sKey)
End Function

'Initialises the class from within the static superclass. This method is meant for internal use only. Use at your own risk.
'@protected
'
'@param {string}  Pattern - The pattern to match
'@param {string}  Flags - Optional flags to apply
'@example
'    obj.init("A\d+","i")
Friend Sub init(ByVal pattern As String, ByVal flags As String)
  'TODO:
End Sub


















'Get/Set pattern for regex
Public Property Get pattern() As String
  'TODO:
End Property
Public Property Let pattern(val As String)
  'TODO:
End Property

'Get/Set Flags for regex
Public Property Get flags() As String
  'TODO:
End Property
Public Property Let flags(val As String)
  'TODO:
End Property



'Given a text string, return whether the source regex is present
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {Boolean} True if the regex is present, false if the regex is not present
'@example
'Regex:
'    \d{5}-ST[A-Z]\d
'
'Data:
'    * The site with id 12345-STA1 is one hell of a cool site.
'Result:
'    True
Public Function Test(ByVal sHaystack As String) As Boolean
  'TODO:
End Function



'Given a text string, return the first match.
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {Dictionary} A dictionary containing the match, submatches, named matches, Count and Raw match object
'@example
'Regex:
'    (?<id>\d{5}-ST[A-Z]\d) - (?<desc>.*)
'
'Data:
'    Some sites were in critical condition
'    * 12345-STA1 - Large crack through pipe.
'    * 12323-STB9 - Acid leakage polluting watercourse.
'    and some others were largely ok:
'    * 23565-STC2
'    * 62346-STZ9
'Result:
'    {
'      0: "12345-STA1 - Large crack through pipe.",
'      "id":"12345-STA1",
'      1: "12345-STA1",
'      "desc": "Large crack through pipe."
'      2: "Large crack through pipe.",
'      "$COUNT":2,
'      "$RAW": {...}
'    }
Public Function Match(ByVal sHaystack As String) As Object
  TODO:
End Function



'Given a text string, return all strings which match the source regex pattern.
'
'@param {string}  sHaystack - Text to search for regex in.
'@return {Collection<Dictionary>} An array of strings which match the regex.
'@example
'Regex:
'    \d{5}-ST[A-Z]\d
'
'Data:
'    Some sites were in critical condition
'    * 12345-STA1 - Large crack through pipe.
'    * 12323-STB9 - Acid leakage polluting watercourse.
'    and some others were largely ok:
'    * 23565-STC2
'    * 62346-STZ9
'Result:
'    [{0: "12345-STA1", ...}, {0: "12323-STB9", ...}, {0: "23565-STC2", ...}, {0: "62346-STZ9", ...}]
Public Function MatchAll(ByVal sHaystack As String) As Collection
  'TODO:
End Function



'Given a Haystack and a Replacer, Replace all matches with the source regex with the format given in the Replacer.
'
'@param {string}  sHaystack - Text to search for regex in.
'@param {string}  sReplacer - Format of replacement text.
'@return {string} The haystack with replaced details.
'@example
'Regex:
'    (?<id>\d{5}-ST[A-Z]\d))\s+(?<count>\d+)\s+(?<date>../../....)
'
'Data:
'    Here is some cool data:
'    12345-STA1  123    10/02/2019
'    12323-STB9  2123   01/01/2005
'    and here is some more:
'    23565-STC2  23     ??/??/????
'    62346-STZ9  5      01/05/1932
'Replacer: (to csv with column re-order)
'    $id,$date,$count
'Result:
'    Here is some cool data:
'    12345-STA1,10/02/2019,123
'    12323-STB9,01/01/2005,2123
'    and here is some more:
'    23565-STC2,??/??/????,23
'    62346-STZ9,01/05/1932,5
Public Function Replace(ByVal sHaystack As String, ByVal sReplacer As String) As String
  'TODO:
End Function

'Given a Haystack and a Format, Generate a string containing all matches in the format provided.
'
'@param {string}  sHaystack - Text to search for regex in.
'@param {string}  sFormat   - Format of output list.
'@return {string} A list of strings matched.
'@example
'Regex:
'    (?<id>\d{5}-ST[A-Z]\d))\s+(?<count>\d+)\s+(?<date>../../....)
'
'Data:
'    12345-STA1  123    10/02/2019
'    12323-STB9  2123   01/01/2005
'    23565-STC2  23     ??/??/????
'    62346-STZ9  5      01/05/1932
'Format: (to csv with column re-order)
'    $id,$date,$count\n
'Result:
'    12345-STA1,10/02/2019,123
'    12323-STB9,01/01/2005,2123
'    23565-STC2,??/??/????,23
'    62346-STZ9,01/05/1932,5
Public Function List(ByVal sHaystack As String, ByVal sFormat As String) As String
  If Me Is stdRegex Then
    stdError.Raise ("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim matches As Collection
  Set matches = MatchAll(sHaystack)
  
  'Loop over all matches and list
  Dim match As variant, j As Long, sList As String, sTmpFormat As String
  For each match in matches
    'Concatenate to sList
    sList = sList & FormatFromMatch(match,sFormat)
  Next

  'Return sList
  List = sList
End Function



'Supplied with a haystack to search, ListArr will return a 2d array of data where the rows represent each match, and the columns represent the formats passed into vFormats parameter
'@param {ByVal String} Haystack to search
'@param {ByVal Array<String>} Array of formats to print into seperate columns. E.G. `Array("$id-$date","$detail")`
'@returns {Array[2d]<String>} All matches as a 2d array. Useful especially for dumping matches to Excel.
Public Function ListArr(ByVal sHaystack as string, ByVal vFormats as variant) as variant
  If Me Is stdRegex Then
    Call stdError.Raise("Method called on class not object")
    Exit Function
  End If

  'Get all matches
  Dim matches As Collection
  Set matches = MatchAll(sHaystack)
  
  Dim numCols as long: numCols = ubound(vFormats)-lbound(vFormats)+1
  
  Dim v() as variant
  Redim v(1 to matches.count, 1 to numCols)

  'Loop over all matches
  Dim match As variant, j As Long, iRow as long, iCol as long
  For each match in matches
    iRow = iRow + 1
    iCol = 0
    For j = lbound(vFormats) to ubound(vFormats)
      iCol = iCol + 1
      v(iRow,iCol) = FormatFromMatch(match,vFormats(j))
    Next
  Next

  'Return sList
  ListArr = v
End Function


'*******************
'* PRIVATE METHODS *
'*******************

'Supplied with a format string and match object, dump the required information from the match object to the string, and return it.
'@param {ByVal Dictionary} Match dictionary with numerical keys for each match, key 0 is whole match and also named matches.
'@param {ByVal string} Format to return data in. Indexed submatches to be extracted using $1, $2, ... . Named submatches to be extracted using $name syntax. \r\n will be 
'@returns {String} String formatted with data from match object.
Private Function FormatFromMatch(ByVal match as object, ByVal sFormat as string) as string
  'TODO: Make this more efficient, ideally we'd only loop through the format once

  'Get keys of each match
  Dim keys As Variant
  keys = match.keys()

  'Loop over default keys and make replacements
  sFormat = VBA.Replace(sFormat, "$&", match(0))
  sFormat = VBA.Replace(sFormat, "\\", "f934b47b-b799-4f52-9b4c-f170a82b52fe")
  sFormat = VBA.Replace(sFormat, "\$", "bfde0637-8207-418d-9d34-68ba974e8110")
  sFormat = VBA.Replace(sFormat, "\r", Chr(13))
  sFormat = VBA.Replace(sFormat, "\n", Chr(10))
  sFormat = VBA.Replace(sFormat, "f934b47b-b799-4f52-9b4c-f170a82b52fe", "\")

  'Loop over remaining keys and make replacements
  For j = 1 To UBound(keys)
    Select Case keys(j)
      Case "$RAW", "$COUNT"
      Case Else
        sFormat = VBA.Replace(sFormat, "$" & keys(j), match(keys(j)))
    End Select
  Next

  sFormat = VBA.Replace(sFormat, "bfde0637-8207-418d-9d34-68ba974e8110","$")
  FormatFromMatch = sFormat
End Function
