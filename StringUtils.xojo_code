#tag Module
Protected Module StringUtils
	#tag Method, Flags = &h1
		Protected Function Chop(s As String, charsToCut As Integer) As String
		  // Return s with the rightmost 'charsToCut' chars removed.
		  
		  Dim charsLeft As Integer = s.Len - charsToCut
		  if charsLeft <= 0 then return ""
		  return s.Left( s.Len - charsToCut )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Chop(s As String, stringToCut As String) As String
		  // Chops 'stringToCut' off of s, if stringToCut is found at the end.
		  // Useful for removing file extensions, trailing punctuation, etc.
		  
		  Dim cutLen As Integer = stringToCut.Len
		  if Right(s, cutLen) = stringToCut then
		    return s.Left( s.Len - cutLen )
		  else
		    return s
		  end if
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ChopB(s As String, bytesToCut As Integer) As String
		  // Return s with the rightmost 'bytesToCut' bytes removed.
		  
		  Dim bytesLeft As Integer = s.LenB - bytesToCut
		  if bytesLeft <= 0 then return ""
		  return s.LeftB( s.LenB - bytesToCut )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ChopB(s As String, stringToCut As String) As String
		  // Chops 'stringToCut' off of s, if stringToCut is found at the end.
		  // Useful for removing file extensions, trailing punctuation, etc.
		  
		  Dim cutLenB As Integer = stringToCut.LenB
		  if StrComp( RightB(s, cutLenB), stringToCut, 0 ) = 0 then
		    return s.LeftB( s.LenB - cutLenB )
		  else
		    return s
		  end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Contains(extends s As String, what As String) As Boolean
		  // Return true if 's' contains the substring 'what'.
		  // By "contains" we mean case-insensitive, encoding-savvy containment
		  // as with InStr.
		  
		  if what = "" then return true
		  return InStr( s, what ) > 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ContainsB(extends s As String, what As String) As Boolean
		  // Return true if 's' contains the substring 'what'.
		  // By "contains" we mean binary containment
		  // as with InStrB.
		  
		  if what = "" then return true
		  return InStrB( s, what ) > 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ControlCharacters() As String
		  // Return the control character region of the ASCII set,
		  // i.e., ASCII 0 through 31.
		  Dim i As Integer
		  if mControlChars = "" then
		    for i = 0 to 31
		      mControlChars = mControlChars + Encodings.ASCII.Chr(i)
		    next
		  end if
		  
		  return mControlChars
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Count(source As String, substr As String) As Integer
		  // Return how many non-overlapping occurrences of 'substr' there
		  // are in 'source'.
		  
		  dim theCount as Integer
		  dim substrLength as Integer
		  dim start as Integer
		  
		  substrLength = Len(substr)
		  if substrLength = 0 then return Len(source) + 1
		  
		  start = 1
		  Do
		    start= InStr(start, source, substr)
		    If start < 1 then return theCount
		    theCount = theCount + 1
		    start = start + substrLength
		  Loop
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CountB(source As String, substr As String) As Integer
		  // Return how many non-overlapping occurrences of 'substr' there
		  // are in 'source', doing binary comparison.
		  
		  dim theCount as Integer
		  dim substrLength as Integer
		  dim start as Integer
		  
		  substrLength = Len(substr)
		  if substrLength = 0 then return LenB(source) + 1
		  
		  start = 1
		  Do
		    start= InStrB(start, source, substr)
		    If start < 1 then return theCount
		    theCount = theCount + 1
		    start = start + substrLength
		  Loop
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CountFieldsQuoted(src as string, sep as string) As integer
		  // Equivalent to RB's CountFields() function, but respects quoted values
		  // Usage:
		  //    s = """Hello, Kitty"", ""One"", ""Two, Three"""
		  //    x = CountFieldsQuoted(s, ",")
		  // result: x=3
		  
		  #pragma DisableBoundsChecking
		  #pragma NilObjectChecking false
		  
		  if InStr( src, sep ) = 0 then return 1
		  if InStr(src,"""")=0 then return CountFields(src, sep)
		  
		  dim countParts, i, n, c as integer
		  dim sepLen as integer = len( sep )
		  dim parts( -1 ) as string
		  
		  parts = split( src, """" )
		  countParts = UBound( parts )
		  for i = 0 to countParts step 2
		    n = InStr( parts( i ), sep )
		    while n > 0
		      c = c + 1
		      n = InStr( n + sepLen, parts( i ), sep )
		    wend
		  next i
		  
		  return c + 1
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CountRegEx(s As String, pattern As String) As Integer
		  // Count the number of occurrences of a RegEx pattern within a string.
		  
		  Dim out As Integer
		  
		  Dim re As New RegEx
		  Dim rm As RegExMatch
		  
		  re.SearchPattern = pattern
		  rm = re.Search( s )
		  while rm <> nil
		    'System.DebugLog rm.SubExpressionString(0) + " at " + str(rm.SubExpressionStartB(0)) + " matches " + pattern + " in " + s
		    out = out + 1
		    rm = re.Search
		  wend
		  
		  return out
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DecimalSeparator() As String
		  // Return the decimal separator the user uses (either "." or ",").
		  if mDecimalSeparator = "" then
		    mDecimalSeparator = Format(1.2, "0.0")
		    mDecimalSeparator = Mid( mDecimalSeparator, 2, 1 )
		  end if
		  
		  return mDecimalSeparator
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function DecodeCase(s As String) As String
		  // This function undoes the encoding of case done by EncodeCase, resulting
		  // in the original string.
		  
		  Dim parts() As String = Split( s, "^" )
		  for i As Integer = 1 to UBound( parts ) step 2
		    if parts(i) = "" then
		      parts(i) = "^"
		    else
		      parts(i) = Uppercase( parts(i) )
		    end if
		  next
		  return Join( parts, "" )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function EditDistance(s1 As String, s2 As String) As Integer
		  // Return the Levenshtein distance, aka the edit distance,
		  // between the two StringUtils.  That's the number of insertions,
		  // deletions, or changes required to make one string match the other.
		  // A result of 0 means the strings
		  // are identical; higher values mean more different.
		  
		  // Note that this function is case-sensitive; if you want a case-
		  // insensitive measure, simply Uppercase or Lowercase both strings
		  // before calling.
		  
		  // Implementation adapted from <http://www.merriampark.com/ld.htm>,
		  // though we're using only a 1D array since the 2D array is wasteful.
		  
		  #pragma DisableBackgroundTasks
		  #pragma DisableBoundsChecking
		  #pragma DisableAutoWaitCursor
		  
		  Dim n, m As Integer
		  n = s1.Len
		  m = s2.Len
		  if n = 0 then return m
		  if m = 0 then return n
		  
		  Dim i, j, cost As Integer
		  Dim d(-1) As Integer
		  Redim d(m)
		  for j = 1 to m
		    d(j) = j
		  next
		  
		  Dim s1chars(-1), s2chars(-1) As String
		  s1chars = Split( s1, "" )
		  s2chars = Split( s2, "" )
		  
		  Dim s1char As String
		  Dim lastCost, nextCost As Integer
		  Dim a, b, c As Integer
		  Dim jMinus1 As Integer
		  
		  for i = 1 to n
		    s1char = s1chars(i-1)
		    lastCost = i
		    jMinus1 = 0
		    for j = 1 to m
		      if StrComp(s1char, s2chars(jMinus1),0) = 0 then cost = 0 else cost = 1
		      
		      // set nextCost to the minimum of the following three possibilities:
		      a = d(j) + 1
		      b = lastCost + 1
		      c = cost + d(jMinus1)
		      
		      if a < b then
		        if c < a then nextCost = c else nextCost = a
		      else
		        if c < b then nextCost = c else nextCost = b
		      end if
		      
		      d(jMinus1) = lastCost
		      lastCost = nextCost
		      jMinus1 = j
		    next
		    d(m) = lastCost
		  next
		  
		  return nextCost
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function EncodeCase(s As String) As String
		  // Return a version of the given string that encodes the uppercase/lowercase state
		  // of each letter as follows: every time we switch from uppercase to lowercase, we
		  // insert a "^".  A "^" in the original string will be represented as "^^".  Finally, we
		  // will assume lowercase at the beginning of the string (so an initial capital letter
		  // will be preceded by "^").
		  //
		  // This can be useful for doing things like case-sensitive lookup of strings in a
		  // Dictionary, which is normally case insensitive.
		  //
		  // For the inverse operation, see DecodeCase.
		  
		  Dim chars() As String = Split( s, "" )
		  Dim inUpperState As Boolean = false
		  Dim idx As Integer
		  Dim anyUpper As Boolean
		  for idx = 0 to UBound( chars )
		    Dim c As String = chars(idx)
		    if c = "^" then
		      // got a caret in the original string; double it
		      chars.Insert idx, "^"
		      idx = idx + 1
		    else
		      // determine whether the given character fails to match the current state
		      Dim codepoint As Integer = Asc(c)
		      Dim switchCase As Boolean
		      if inUpperState then
		        if codepoint < 128 then
		          // lowercase ASCII ranges from 97 to 122
		          switchCase = (codepoint >= 97 and codepoint <= 122)
		        else
		          // if not ASCII, use RB's routines to see if it'd be any different in upper case
		          switchCase = StrComp( c, Uppercase(c), 0 ) <> 0
		        end if
		      else
		        if codepoint < 128 then
		          // uppercase ASCII ranges from 65 to 90
		          switchCase = (codepoint >= 65 and codepoint <= 90)
		        else
		          // if not ASCII, use RB's routines to see if it'd be any different in lower case
		          switchCase = StrComp( c, Lowercase(c), 0 ) <> 0
		        end if
		        if switchCase then anyUpper = true
		      end if
		      if switchCase then
		        // yep, time to switch case
		        chars.Insert idx, "^"
		        idx = idx + 1
		        inUpperState = not inUpperState
		      end if
		    end if
		  next
		  
		  Dim result As String = Join( chars, "" )
		  if anyUpper then result = Lowercase( result )
		  return result
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function EndsWith(extends s As String, withWhat As String) As Boolean
		  // Return true if 's' ends with the string 'withWhat',
		  // doing a standard string comparison.
		  
		  return Right(s, withWhat.Len) = withWhat
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function EndsWithB(extends s As String, withWhat As String) As Boolean
		  // Return true if 's' ends with the string 'withWhat',
		  // doing a binary comparison.
		  
		  return StrComp( RightB(s, withWhat.LenB), withWhat, 0 ) = 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GuessEncoding(s As String) As TextEncoding
		  // Guess what text encoding the text in the given string is in.
		  //
		  // Note that it this version does not report whether the UTF-32 or
		  // UTF-16 that is in the wrong byte order.  If that's a possibility,
		  // you should use the other version of GuessEncoding instead.
		  
		  Dim whoCares As Boolean
		  return StringUtils.GuessEncoding( s, whoCares )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GuessEncoding(s As String, ByRef outWrongOrder As Boolean) As TextEncoding
		  // Guess what text encoding the text in the given string is in.
		  // This ignores the encoding set on the string, and guesses
		  // one of the following:
		  //
		  //   * UTF-32
		  //   * UTF-16
		  //   * UTF-8
		  //   * Encodings.SystemDefault
		  //
		  // If the UTF-32 or UTF-16 is in the wrong byte order for this platform,
		  // then outWrongOrder will be set to true.
		  
		  static isBigEndian, endianChecked As Boolean
		  if not endianChecked then
		    Dim temp As String = Encodings.UTF16.Chr( &hFEFF )
		    isBigEndian = (AscB( MidB( temp, 1, 1 ) ) = &hFE)
		    endianChecked = true
		  end if
		  
		  // check for a BOM
		  Dim b0 As Integer = AscB( s.MidB( 1, 1 ) )
		  Dim b1 As Integer = AscB( s.MidB( 2, 1 ) )
		  Dim b2 As Integer = AscB( s.MidB( 3, 1 ) )
		  Dim b3 As Integer = AscB( s.MidB( 4, 1 ) )
		  if b0=0 and b1=0 and b2=&hFE and b3=&hFF then
		    // UTF-32, big-endian
		    outWrongOrder = not isBigEndian
		    return Encodings.UTF32BE
		  elseif b0=&hFF and b1=&hFE and b2=0 and b3=0 and s.LenB >= 4 then
		    // UTF-32, little-endian
		    outWrongOrder = isBigEndian
		    return Encodings.UTF32LE
		  elseif b0=&hFE and b1=&hFF then
		    // UTF-16, big-endian
		    outWrongOrder = not isBigEndian
		    return Encodings.UTF16
		  elseif b0=&hFF and b1=&hFE then
		    // UTF-16, little-endian
		    outWrongOrder = isBigEndian
		    return Encodings.UTF16
		  elseif b0=&hEF and b1=&hBB and b1=&hBF then
		    // UTF-8 (ah, a sensible encoding where endianness doesn't matter!)
		    return Encodings.UTF8
		  end if
		  
		  // no BOM; see if it's entirely ASCII.
		  Dim m As MemoryBlock = s
		  Dim i, maxi As Integer = s.LenB - 1
		  for i = 0 to maxi
		    if m.Byte(i) > 127 then exit
		  next
		  if i > maxi then return Encodings.ASCII
		  
		  // Not ASCII; check for a high incidence of nulls every other byte,
		  // which suggests UTF-16 (at least in Roman text).
		  Dim nulls(1) As Integer  // null count in even (0) and odd (1) bytes
		  for i = 0 to maxi
		    if m.Byte(i) = 0 then
		      nulls(i mod 2) = nulls(i mod 2) + 1
		    end if
		  next
		  if nulls(0) > nulls(1)*2 and nulls(0) > maxi\2 then
		    // UTF-16, big-endian
		    outWrongOrder = not isBigEndian
		    return Encodings.UTF16
		  elseif nulls(1) > nulls(0)*2 and nulls(1) > maxi\2 then
		    // UTF-16, little-endian
		    outWrongOrder = isBigEndian
		    return Encodings.UTF16
		  end if
		  
		  // it's not ASCII; check for illegal UTF-8 characters.
		  // See Table 3.1B, "Legal UTF-8 Byte Sequences",
		  // at <http://unicode.org/versions/corrigendum1.html>
		  Dim b As Byte
		  for i = 0 to maxi
		    select case m.Byte(i)
		    case &h00 to &h7F
		      // single-byte character; just continue
		    case &hC2 to &hDF
		      // one additional byte
		      if i+1 > maxi then exit for
		      b = m.Byte(i+1)
		      if b < &h80 or b > &hBF then exit for
		      i = i+1
		    case &hE0
		      // two additional bytes
		      if i+2 > maxi then exit for
		      b = m.Byte(i+1)
		      if b < &hA0 or b > &hBF then exit for
		      b = m.Byte(i+2)
		      if b < &h80 or b > &hBF then exit for
		      i = i+2
		    case &hE1 to &hEF
		      // two additional bytes
		      if i+2 > maxi then exit for
		      b = m.Byte(i+1)
		      if b < &h80 or b > &hBF then exit for
		      b = m.Byte(i+2)
		      if b < &h80 or b > &hBF then exit for
		      i = i+2
		    case &hF0
		      // three additional bytes
		      if i+3 > maxi then exit for
		      b = m.Byte(i+1)
		      if b < &h90 or b > &hBF then exit for
		      b = m.Byte(i+2)
		      if b < &h80 or b > &hBF then exit for
		      b = m.Byte(i+3)
		      if b < &h80 or b > &hBF then exit for
		      i = i+3
		    case &hF1 to &hF3
		      // three additional bytes
		      if i+3 > maxi then exit for
		      b = m.Byte(i+1)
		      if b < &h80 or b > &hBF then exit for
		      b = m.Byte(i+2)
		      if b < &h80 or b > &hBF then exit for
		      b = m.Byte(i+3)
		      if b < &h80 or b > &hBF then exit for
		      i = i+3
		    case &hF4
		      // three additional bytes
		      if i+3 > maxi then exit for
		      b = m.Byte(i+1)
		      if b < &h80 or b > &h8F then exit for
		      b = m.Byte(i+2)
		      if b < &h80 or b > &hBF then exit for
		      b = m.Byte(i+3)
		      if b < &h80 or b > &hBF then exit for
		      i = i+3
		    else
		      exit for
		    end select
		  next i
		  if i > maxi then return Encodings.UTF8  // no illegal UTF-8 sequences, so that's probably what it is
		  
		  // If not valid UTF-8, then let's just guess the system default.
		  return Encodings.SystemDefault
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Hash(s As String) As Integer
		  // Return the hash value of the given string, as used by RB's
		  // Variant and Dictionary classes.
		  
		  Dim v As Variant
		  v = s
		  return v.Hash
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function HexB(s As String) As String
		  // Return a hex representation of each byte of s,
		  // i.e., each byte becomes a pair of hexadecimal digits,
		  // separated by spaces from the next byte.
		  
		  // Credit: Charles Yeomans.
		  
		  if s = "" then return ""
		  
		  #if RBVersion > 5.5
		    Static HexTable as MemoryBlock
		  #else
		    Dim HexTable as MemoryBlock
		  #endif
		  
		  if HexTable Is Nil then
		    HexTable = new MemoryBlock(512)
		    For i as Integer = 0 to 15
		      HexTable.StringValue(i + i , 2) = "0" + Hex(i)
		    Next
		    For i as Integer = 16 to 255
		      HexTable.StringValue(i + i, 2) = Hex(i)
		    Next
		  end if
		  
		  dim inData as MemoryBlock = new MemoryBlock(LenB(s))
		  inData.StringValue(0, inData.Size) = s
		  
		  dim outData as MemoryBlock = new MemoryBlock(3*inData.Size - 1)
		  outData.Short(0) = HexTable.Short(2*inData.Byte(0))
		  
		  dim lastByte as Integer = inData.Size - 1
		  For offset as Integer = 1 to lastByte
		    outData.Byte(3*offset - 1) = 32 // (space)
		    outData.Short(3*offset) = HexTable.Short(2*inData.Byte(offset))
		  Next
		  
		  Return outData.StringValue(0, outData.Size, Encodings.ASCII)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InStrQuoted(start As Integer = 1, source As String, find As String) As Integer
		  // This is the same as the built-in InStr function, except that it ignores
		  // any occurrence of "find" within double quotes.
		  
		  
		  do
		    // Find the next occurrence of the search string.  If none, then we're done.
		    Dim foundPos As Integer = InStr( start, source, find )
		    if foundPos < 1 then return foundPos
		    
		    // Now, also find the next quotation mark; if this comes before
		    // our foundPos, then we need to ignore that and resume searching
		    // after the closing quote.  Otherwise, we're done.
		    Dim quotePos As Integer = InStr( start, source, """" )
		    if quotePos = 0 or quotePos > foundPos then return foundPos
		    
		    Dim closeQuotePos As Integer = InStr( quotePos+1, source, """" )
		    if closeQuotePos = 0 then return 0 // no closing quote, treat whole end of string as quoted
		    
		    start = closeQuotePos + 1
		  loop
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InStrReverse(startPos As Integer = - 1, source As String, substr As String) As Integer
		  // Similar to InStr, but searches backwards from the given position
		  // (or if startPos = -1, then from the end of the string).
		  // If substr can't be found, returns 0.
		  
		  Dim srcLen As Integer = source.Len
		  if startPos = -1 then startPos = srcLen
		  
		  // Here's an easy way...
		  // There may be a faster implementation, but then again, there may not -- it probably
		  // depends on the particulars of the data.
		  Dim reversedSource As String = Reverse(source)
		  Dim reversedSubstr As String = Reverse(substr)
		  Dim reversedPos As Integer
		  reversedPos = InStr( srcLen - startPos + 1, reversedSource, reversedSubstr )
		  if reversedPos < 1 then return 0
		  return srcLen - reversedPos - substr.Len + 2
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InStrReverseB(startPosB As Integer = - 1, source As String, substr As String) As Integer
		  // Similar to InStrB, but searches backwards from the given position
		  // (or if startPosB = -1, then from the end of the string).
		  // If substr can't be found, returns 0.
		  
		  Dim srcLen As Integer = source.LenB
		  Dim subLen As Integer = substr.LenB
		  if startPosB = -1 then startPosB = srcLen
		  
		  // We'll do a simple sequential search.  A Boyer-Moore algorithm
		  // would work better in many cases, but we'd have to rewrite the
		  // whole algorithm to work backwards.  The sequential search will
		  // be good enough in most cases anyway.
		  Dim posB As Integer
		  for posB = Min( srcLen - subLen + 1, startPosB ) downTo 1
		    if StrComp( MidB( source, posB, subLen ), substr, 0 ) = 0 then return posB
		  next posB
		  
		  return 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function IsEmpty(extends s As String) As Boolean
		  // Return true if the string is empty.
		  
		  return s = ""
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function JoinQuoted(fields() as String, delimiter as String) As String
		  // Join the given strings with a delimiter, just like RB's intrinsic Join
		  // method, except that if any of the fields contains the delimiter,
		  // that item will be surrounded by quotes in the output.  See
		  // SplitQuoted for the inverse function.
		  
		  // Approach: copy the items into a second array, putting quotes
		  // around any that contain the delimiter, then Join them.  This
		  // way we don't mutate the array that's passed in.
		  
		  Dim quoted() As String
		  Dim ub As Integer = UBound( fields )
		  Redim quoted( ub )
		  for i As Integer = 0 to ub
		    Dim fld As String = fields(i)
		    if Instr( fld, Delimiter ) > 0 then
		      quoted(i) = """" + fld + """"
		    else
		      quoted(i) = fld
		    end if
		  next
		  
		  return Join( quoted, delimiter )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LineEnding(s As String) As String
		  // Return the first line ending (of any standard sort: Unix, classic Mac, or Windows)
		  // which is found in the given string.  If none is found, return standard EndOfLine
		  // for the platform we're running on.
		  
		  // The code below can't deal with UTF-16, but should deal fine with anything else.
		  // So, in those rare cases where we have UTF-16, let's convert.  This could be a
		  // performance bottleneck if you call this on large UTF-16 strings.
		  if s.Encoding = Encodings.UTF16 then s = s.ConvertEncoding( Encodings.UTF8 )
		  
		  // Loop over the bytes of the string until we find a 13 or 10.  A line ending should
		  // be either 13 alone (Mac), 10 alone (Unix), or 13+10 (Windows).
		  Dim posB, maxPosB As Integer
		  maxPosB = LenB( s )
		  for posB = 1 to maxPosB
		    Dim b As Integer = AscB( MidB( s, posB, 1 ) )
		    if b = 10 then
		      return Chr(10)
		    elseif b = 13 then
		      if AscB( MidB( s, posB+1, 1 ) ) = 10 then return Chr(13) + Chr(10)
		      return Chr(13)
		    end if
		  next
		  
		  return EndOfLine
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LTrim(source As String, charsToTrim As String) As String
		  // This is an extended version of RB's LTrim function that lets you specify
		  // a set of characters to trim.
		  
		  Dim srcLen As Integer = source.Len
		  Dim leftPos, i As Integer
		  for i = 1 to srcLen
		    if InStr( charsToTrim, Mid(source, i, 1) ) = 0 then exit
		  next
		  leftPos = i
		  if leftPos > srcLen then return ""
		  
		  return Mid( source, leftPos )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function MatchCase(textToChange As String, sampleText As String) As String
		  // Return a version of textToChange that matches the case style
		  // of sampleText: Lowercase, Uppercase, or Titlecase.
		  
		  // Examine the text until we find 2 characters that have case.
		  Dim pos As Integer = 1
		  Dim maxPos As Integer = sampleText.Len
		  Dim foundUpper() As Boolean
		  Dim c, upperC, lowerC As String
		  for pos = 1 to maxPos
		    c = MidB( sampleText, pos, 1 )
		    upperC = Uppercase(c)
		    lowerC = Lowercase(c)
		    if StrComp( upperC, lowerC, 0 ) <> 0 then
		      // found a character with case; remember which case it is
		      foundUpper.Append StrComp(c, upperC, 0 ) = 0
		      if UBound( foundUpper ) >= 1 then exit
		    else
		      // found a caseless character; reset our found list
		      Redim foundUpper(-1)
		    end if
		  next
		  
		  if UBound( foundUpper ) < 1 then
		    // didn't find enough characters with case; sample is no good
		    return textToChange
		  end if
		  
		  if foundUpper(0) then
		    if foundUpper(1) then return Uppercase(textToChange)
		    return Titlecase(textToChange)
		  end if
		  return Lowercase(textToChange)
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Metaphone(source As String, ByRef outPrimary As String, ByRef outAlternate As String)
		  // Compute the Double Metaphone of the source string.  This is an algorithm that
		  // finds one or two approximate phonetic representations of a string, useful in
		  // searching for almost-matches -- e.g., looking for names whose spelling may have
		  // varied, or correcting typos made by the user, and so on.
		  //
		  // The output is roughly human-readable, with the following conventions:
		  //   Vowels are omitted from the output, except for a vowel at the beginning
		  //      of a word, which is represented by an A (e.g. "ox" becomes "AKS")
		  //   X is used to represent a "ch" sound (e.g., "church" becomes "XRX")
		  //   0 (zero) is used to represent a "th" sound (e.g. "think" becomes "0NK")
		  //
		  // For more information about Double Metaphone, see:
		  //     http://aspell.sourceforge.net/metaphone/
		  //     http://www.cuj.com/articles/2000/0006/0006d/0006d.htm?topic=articles
		  //
		  // This implementation is based on the one at:
		  //     http://aspell.sourceforge.net/metaphone/dmetaph.cpp
		  
		  
		  Dim length As Integer
		  length = source.Len
		  if length < 1 then
		    outPrimary = ""
		    outAlternate = ""
		    return
		  end if
		  
		  source = Uppercase(source) + " "
		  Dim current As Integer = 1
		  
		  Dim charAt(-1) As String
		  charAt = source.Split("")
		  charAt.Insert 0, ""  // (make it 1-based, like Mid)
		  
		  Dim slavoGermanic As Boolean
		  if InStr(source, "W") > 0 or InStr(source, "K") > 0 _
		    or InStr(source, "CZ") > 0 or InStr(source, "WITZ") > 0 then
		    slavoGermanic = true
		  end if
		  
		  Dim out1, out2 As String
		  
		  // skip these when at start of word
		  if MStringAt(source, 1, 2, "GN", "KN", "PN", "WR", "PS") then current = current + 1
		  
		  // initial 'X' is pronounced 'Z' e.g. 'Xavier'
		  if charAt(1) = "X" then
		    out1 = out1 + "S"
		    out2 = out2 + "S"  // "Z" maps to "S"
		    current = current + 1
		  end if
		  
		  //---------- main loop ---------------
		  while current <= length
		    
		    select case charAt(current)
		      
		    case "A", "E", "I", "O", "U", "Y"
		      if current = 1 then
		        // all initial vowels map to "A"; elsewhere they're skipped
		        out1 = out1 + "A"
		        out2 = out2 + "A"
		      end if
		      current = current + 1
		      
		    case "B"
		      //"-mb", e.g", "dumb", already skipped over...
		      out1 = out1 + "P"
		      out2 = out2 + "P"
		      if charAt(current + 1) = "B" then
		        current = current + 2
		      else
		        current = current + 1
		      end if
		      
		    case "Ç"
		      out1 = out1 + "S"
		      out2 = out2 + "S"
		      current = current +  1
		      
		    case "C"
		      // various germanic
		      if current > 2 _
		        and not MIsVowel(source, current - 2) _
		        and MStringAt(source, (current - 1), 3, "ACH") _
		        and (charAt(current + 2) <> "I" and (charAt(current + 2) <> "E"_
		        or MStringAt(source, current - 2, 6, "BACHER", "MACHER")) ) then
		        out1 = out1 + "K"
		        out2 = out2 + "K"
		        current = current + 2
		        
		      elseif current = 1 AND MStringAt(source, current, 6, "CAESAR") then
		        // special case 'caesar' (why didn't this go at the top?)
		        out1 = out1 + "S"
		        out2 = out2 + "S"
		        current = current + 2
		        
		      elseif MStringAt(source, current, 4, "CHIA") then
		        // italian 'chianti'
		        out1 = out1 + "K"
		        out2 = out2 + "K"
		        current = current + 2
		        
		      elseif MStringAt(source, current, 2, "CH") then
		        // find 'michael'
		        if current > 0 AND MStringAt(source, current, 4, "CHAE") then
		          out1 = out1 + "K"
		          out2 = out2 + "X"
		          current = current + 2
		          break
		          
		        elseif current = 0 _
		          and (MStringAt(source, current + 1, 5, "HARAC", "HARIS") _
		          or MStringAt(source, current + 1, 3, "HOR", "HYM", "HIA", "HEM")) _
		          and not MStringAt(source, 0, 5, "CHORE") then
		          // greek roots e.g. 'chemistry', 'chorus'
		          out1 = out1 + "K"
		          out2 = out2 + "K"
		          current = current + 2
		          
		        else
		          //germanic, greek, or otherwise 'ch' for 'kh' sound
		          if((MStringAt(source, 0, 4, "VAN ", "VON ") or MStringAt(source, 0, 3, "SCH")) _
		            _ // 'architect but not 'arch', 'orchestra', 'orchid'
		            or MStringAt(source, current - 2, 6, "ORCHES", "ARCHIT", "ORCHID") _
		            or MStringAt(source, current + 2, 1, "T", "S") _
		            or ((MStringAt(source, current - 1, 1, "A", "O", "U", "E") OR current = 1) _
		            _ //e.g., 'wachtler', 'wechsler', but not 'tichner'
		            and MStringAt(source, current + 2, 1, "L", "R", "N", "M", "B", "H", "F", "V", "W", " "))) then
		            out1 = out1 + "K"
		            out2 = out2 + "K"
		          else
		            if current > 1 then
		              if MStringAt(source, 1, 2, "MC") then
		                //e.g., "McHugh"
		                out1 = out1 + "K"
		                out2 = out2 + "K"
		              else
		                out1 = out1 + "X"
		                out2 = out2 + "K"
		              end if
		            else
		              out1 = out1 + "X"
		              out2 = out2 + "X"
		            end if
		          end if
		          current = current + 2
		        end if
		        
		        // end of CH case
		        
		      elseif MStringAt(source, current, 2, "CZ") and not MStringAt(source, current - 2, 4, "WICZ") then
		        //e.g, 'czerny'
		        out1 = out1 + "S"
		        out2 = out2 + "X"
		        current = current +  2
		        
		      elseif MStringAt(source, current + 1, 3, "CIA") then
		        //e.g., 'focaccia'
		        out1 = out1 + "X"
		        out2 = out2 + "X"
		        current = current +  3
		        
		      elseif MStringAt(source, current, 2, "CC") and not (current = 2 AND charAt(1) = "M") then
		        // double "C", but not if e.g. 'McClellan'
		        //'bellocchio' but not 'bacchus'
		        if MStringAt(source, current + 2, 1, "I", "E", "H") and not MStringAt(source, current + 2, 2, "HU") then
		          //'accident', 'accede" "succeed'
		          if((current = 2 AND charAt(current - 1) = "A") _
		            OR MStringAt(source, current - 1, 5, "UCCEE", "UCCES")) then
		            out1 = out1 + "KS"
		            out2 = out2 + "KS"
		            //'bacci', 'bertucci', other italian
		          else
		            out1 = out1 + "X"
		            out2 = out2 + "X"
		          end if
		          current = current +  3
		        else // Pierce's rule
		          out1 = out1 + "K"
		          out2 = out2 + "K"
		          current = current +  2
		        end if
		        
		      elseif MStringAt(source, current, 2, "CK", "CG", "CQ") then
		        out1 = out1 + "K"
		        out2 = out2 + "K"
		        current = current +  2
		        
		      elseif MStringAt(source, current, 2, "CI", "CE", "CY") then
		        // italian vs. english
		        if MStringAt(source, current, 3, "CIO", "CIE", "CIA") then
		          out1 = out1 + "S"
		          out2 = out2 + "X"
		        else
		          out1 = out1 + "S"
		          out2 = out2 + "S"
		        end if
		        current = current +  2
		        
		      else
		        // all other C cases are considered a K:
		        out1 = out1 + "K"
		        out2 = out2 + "K"
		        
		        // name sent in 'mac caffrey', 'mac gregor'
		        if MStringAt(source, current + 1, 2, " C", " Q", " G" ) then
		          current = current +  3
		        else
		          if MStringAt(source, current + 1, 1, "C", "K", "Q") _
		            AND not MStringAt(source, current + 1, 2, "CE", "CI") then
		            current = current +  2
		          else
		            current = current +  1
		          end if
		        end if
		      end if
		      
		    case "D"
		      if MStringAt(source, current, 2, "DG") then
		        if MStringAt(source, current + 2, 1, "I", "E", "Y") then
		          //e.g. 'edge'
		          out1 = out1 + "J"
		          out2 = out2 + "J"
		          current = current +  3
		        else
		          //e.g. 'edgar'
		          out1 = out1 + "TK"
		          out2 = out2 + "TK"
		          current = current +  2
		        end if
		        
		      elseif MStringAt(source, current, 2, "DT", "DD") then
		        out1 = out1 + "T"
		        out2 = out2 + "T"
		        current = current +  2
		      else
		        out1 = out1 + "T"
		        out2 = out2 + "T"
		        current = current +  1
		      end if
		      
		    case "F"
		      out1 = out1 + "F"
		      out2 = out2 + "F"
		      if charAt(current + 1) = "F" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "G"
		      if charAt(current + 1) = "H"  then
		        // GH...
		        if current > 1 AND not MIsVowel(source, current - 1) then
		          out1 = out1 + "K"
		          out2 = out2 + "K"
		          current = current +  2
		          
		        elseif current = 1 then
		          //'ghislane', ghiradelli
		          if charAt(current + 2) = "I" then
		            out1 = out1 + "J"
		            out2 = out2 + "J"
		          else
		            out1 = out1 + "K"
		            out2 = out2 + "K"
		          end if
		          current = current +  2
		          
		        elseif((current > 2 AND MStringAt(source, current - 2, 1, "B", "H", "D") ) _
		          _ //e.g., 'bough'
		          OR (current > 3 AND MStringAt(source, current - 3, 1, "B", "H", "D") ) _
		          _ //e.g., 'broughton'
		          OR (current > 4 AND MStringAt(source, current - 4, 1, "B", "H") ) ) then
		          //Parker's rule (with some further refinements) - e.g., 'hugh'
		          current = current +  2
		          
		        else
		          //e.g., 'laugh', 'McLaughlin', 'cough', 'gough', 'rough', 'tough'
		          if current > 3 _
		            AND charAt(current - 1) = "U" _
		            AND MStringAt(source, current - 3, 1, "C", "G", "L", "R", "T") then
		            out1 = out1 + "F"
		            out2 = out2 + "F"
		          else
		            if((current > 0) AND charAt(current - 1) <> "I") then
		              out1 = out1 + "K"
		              out2 = out2 + "K"
		            end if
		          end if
		          current = current +  2
		        end if
		        
		      elseif charAt(current + 1) = "N" then
		        // GN...
		        if current = 1 AND MIsVowel(source, 0) AND not SlavoGermanic then
		          out1 = out1 + "KN"
		          out2 = out2 + "N"
		        else
		          //not e.g. 'cagney'
		          if not MStringAt(source, current + 2, 2, "EY") _
		            AND charAt(current + 1) <> "Y" AND not SlavoGermanic then
		            out1 = out1 + "N"
		            out2 = out2 + "KN"
		          else
		            out1 = out1 + "KN"
		            out2 = out2 + "KN"
		          end if
		        end if
		        current = current +  2
		        
		        
		      elseif MStringAt(source, current + 1, 2, "LI") AND not SlavoGermanic then
		        //'tagliaro'
		        out1 = out1 + "KL"
		        out2 = out2 + "L"
		        current = current +  2
		        
		        
		      elseif current = 1 _
		        AND (charAt(current + 1) = "Y"  _
		        OR MStringAt(source, current + 1, 2, "ES", "EP", "EB", "EL", "EY", "IB", "IL", "IN", "IE", "EI", "ER")) then
		        //ges-,gep-,gel-, gie- at beginning
		        out1 = out1 + "K"
		        out2 = out2 + "J"
		        current = current +  2
		        
		      elseif (MStringAt(source, current + 1, 2, "ER") OR charAt(current + 1) = "Y") _
		        AND not MStringAt(source, 1, 6, "DANGER", "RANGER", "MANGER") _
		        AND not MStringAt(source, current - 1, 1, "E", "I") _
		        AND not MStringAt(source, current - 1, 3, "RGY", "OGY") then
		        // -ger-,  -gy-
		        out1 = out1 + "K"
		        out2 = out2 + "J"
		        current = current +  2
		        
		      elseif MStringAt(source, current + 1, 1, "E", "I", "Y") OR MStringAt(source, current - 1, 4, "AGGI", "OGGI") then
		        // italian e.g, 'biaggi'
		        //obvious germanic
		        if MStringAt(source, 1, 4, "VAN ", "VON ") OR MStringAt(source, 1, 3, "SCH") _
		          OR MStringAt(source, current + 1, 2, "ET") then
		          out1 = out1 + "K"
		          out2 = out2 + "K"
		        else
		          //always soft if french ending
		          if MStringAt(source, current + 1, 4, "IER ") then
		            out1 = out1 + "J"
		            out2 = out2 + "J"
		          else
		            out1 = out1 + "J"
		            out2 = out2 + "K"
		          end if
		        end if
		        current = current +  2
		        
		      else
		        // any other G
		        out1 = out1 + "K"
		        out2 = out2 + "K"
		        if charAt(current + 1) = "G" then
		          current = current +  2
		        else
		          current = current +  1
		        end if
		      end if
		      
		    case "H"
		      //only keep if first & before vowel or btw. 2 vowels
		      if (current = 1 OR MIsVowel(source, current - 1)) AND MIsVowel(source, current + 1) then
		        out1 = out1 + "H"
		        out2 = out2 + "H"
		        current = current +  2
		      else//also takes care of 'HH'
		        current = current +  1
		      end if
		      
		    case "J"
		      //obvious spanish, 'jose', 'san jacinto'
		      if MStringAt(source, current, 4, "JOSE") OR MStringAt(source, 1, 4, "SAN ") then
		        if (current = 0 AND charAt(current + 4) = " ") OR MStringAt(source, 1, 4, "SAN ") then
		          out1 = out1 + "H"
		          out2 = out2 + "H"
		        else
		          out1 = out1 + "J"
		          out2 = out2 + "H"
		        end if
		        current = current + 1
		        
		      else
		        if current = 0 AND not MStringAt(source, current, 4, "JOSE") then
		          out1 = out1 + "J"
		          out2 = out2 + "A"//Yankelovich/Jankelowicz
		        else
		          //spanish pron. of e.g. 'bajador'
		          if MIsVowel(source, current - 1) _
		            AND not SlavoGermanic _
		            AND (charAt(current + 1) = "A" OR charAt(current + 1) = "O") then
		            out1 = out1 + "J"
		            out2 = out2 + "H"
		          else
		            if current = length then
		              out1 = out1 + "J"
		              out2 = out2 + ""
		            elseif not MStringAt(source, current + 1, 1, "L", "T", "K", "S", "N", "M", "B", "Z") _
		              AND not MStringAt(source, current - 1, 1, "S", "K", "L") then
		              out1 = out1 + "J"
		              out2 = out2 + "J"
		            end if
		          end if
		        end if
		        
		        if charAt(current + 1) = "J" then //it could happen!
		          current = current +  2
		        else
		          current = current +  1
		        end if
		      end if
		      
		    case "K"
		      out1 = out1 + "K"
		      out2 = out2 + "K"
		      if charAt(current + 1) = "K" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "L"
		      if charAt(current + 1) = "L" then
		        //spanish e.g. 'cabrillo', 'gallegos'
		        if (current = length - 2 AND MStringAt(source, current - 1, 4, "ILLO", "ILLA", "ALLE")) _
		          OR ((MStringAt(source, length - 1, 2, "AS", "OS") OR MStringAt(source, length, 1, "A", "O")) _
		          AND MStringAt(source, current - 1, 4, "ALLE")) then
		          out1 = out1 + "L"
		          out2 = out2 + ""
		          current = current +  2
		        else
		          out1 = out1 + "L"
		          out2 = out2 + "L"
		          current = current +  2
		        end if
		      else
		        out1 = out1 + "L"
		        out2 = out2 + "L"
		        current = current +  1
		      end if
		      
		    case "M"
		      out1 = out1 + "M"
		      out2 = out2 + "M"
		      if (MStringAt(source, current - 1, 3, "UMB") _
		        AND (current + 1 = length OR MStringAt(source, current + 2, 2, "ER"))) _
		        _ //'dumb","thumb'
		        OR charAt(current + 1) = "M" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "N"
		      out1 = out1 + "N"
		      out2 = out2 + "N"
		      if charAt(current + 1) = "N" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "Ñ"
		      out1 = out1 + "N"
		      out2 = out2 + "N"
		      current = current +  1
		      
		    case "P"
		      if charAt(current + 1) = "H" then  // PH sounds like F
		        out1 = out1 + "F"
		        out2 = out2 + "F"
		        current = current +  2
		        
		      else
		        out1 = out1 + "P"
		        out2 = out2 + "P"
		        // (also account for "campbell", "raspberry")
		        if MStringAt(source, current + 1, 1, "P", "B") then
		          current = current +  2
		        else
		          current = current +  1
		        end if
		      end if
		      
		    case "Q"
		      out1 = out1 + "K"
		      out2 = out2 + "K"
		      if charAt(current + 1) = "Q" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "R"
		      //french e.g. 'rogier', but exclude 'hochmeier'
		      if current = length AND not SlavoGermanic _
		        AND MStringAt(source, current - 2, 2, "IE") _
		        AND not MStringAt(source, current - 4, 2, "ME", "MA") then
		        out1 = out1 + ""
		        out2 = out2 + "R"
		      else
		        out1 = out1 + "R"
		        out2 = out2 + "R"
		      end if
		      
		      if charAt(current + 1) = "R" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "S"
		      if MStringAt(source, current - 1, 3, "ISL", "YSL") then
		        //special cases 'island', 'isle', 'carlisle', 'carlysle'
		        current = current +  1
		        
		      elseif current = 1 AND MStringAt(source, current, 5, "SUGAR") then
		        //special case 'sugar-'
		        out1 = out1 + "X"
		        out2 = out2 + "S"
		        current = current +  1
		        
		      elseif MStringAt(source, current, 2, "SH") then
		        //germanic
		        if MStringAt(source, current + 1, 4, "HEIM", "HOEK", "HOLM", "HOLZ") then
		          out1 = out1 + "S"
		          out2 = out2 + "S"
		        else
		          out1 = out1 + "X"
		          out2 = out2 + "X"
		        end if
		        current = current +  2
		        
		      elseif MStringAt(source, current, 3, "SIO", "SIA") OR MStringAt(source, current, 4, "SIAN") then
		        //italian & armenian
		        if not SlavoGermanic then
		          out1 = out1 + "S"
		          out2 = out2 + "X"
		        else
		          out1 = out1 + "S"
		          out2 = out2 + "S"
		        end if
		        current = current +  3
		        
		      elseif (current = 1 AND MStringAt(source, current + 1, 1, "M", "N", "L", "W")) _
		        OR MStringAt(source, current + 1, 1, "Z") then
		        //german & anglicisations, e.g. 'smith' match 'schmidt', 'snider' match 'schneider'
		        //also, -sz- in slavic language altho in hungarian it is pronounced "s"
		        out1 = out1 + "S"
		        out2 = out2 + "X"
		        if MStringAt(source, current + 1, 1, "Z") then
		          current = current +  2
		        else
		          current = current +  1
		        end if
		        
		      elseif MStringAt(source, current, 2, "SC") then
		        //Schlesinger's rule
		        if charAt(current + 2) = "H" then
		          //dutch origin, e.g. 'school', 'schooner'
		          if MStringAt(source, current + 3, 2, "OO", "ER", "EN", "UY", "ED", "EM") then
		            //'schermerhorn', 'schenker'
		            if MStringAt(source, current + 3, 2, "ER", "EN") then
		              out1 = out1 + "X"
		              out2 = out2 + "SK"
		            else
		              out1 = out1 + "SK"
		              out2 = out2 + "SK"
		            end if
		            current = current +  3
		            
		          else
		            if current = 1 AND not MIsVowel(source, 4) AND charAt(4) <> "W" then
		              out1 = out1 + "X"
		              out2 = out2 + "S"
		            else
		              out1 = out1 + "X"
		              out2 = out2 + "X"
		            end if
		            current = current +  3
		          end if
		          
		        elseif MStringAt(source, current + 2, 1, "I", "E", "Y") then
		          out1 = out1 + "S"
		          out2 = out2 + "S"
		          current = current +  3
		          
		        else
		          out1 = out1 + "SK"
		          out2 = out2 + "SK"
		          current = current +  3
		        end if
		        
		      else
		        //french e.g. 'resnais', 'artois'
		        if current = length AND MStringAt(source, current - 2, 2, "AI", "OI") then
		          out1 = out1 + ""
		          out2 = out2 + "S"
		        else
		          out1 = out1 + "S"
		          out2 = out2 + "S"
		        end if
		        if MStringAt(source, current + 1, 1, "S", "Z") then
		          current = current +  2
		        else
		          current = current +  1
		        end if
		      end if
		      
		    case "T"
		      if MStringAt(source, current, 4, "TION") then
		        out1 = out1 + "X"
		        out2 = out2 + "X"
		        current = current +  3
		        
		      elseif MStringAt(source, current, 3, "TIA", "TCH") then
		        out1 = out1 + "X"
		        out2 = out2 + "X"
		        current = current +  3
		        
		      elseif MStringAt(source, current, 2, "TH") OR MStringAt(source, current, 3, "TTH") then
		        //special case 'thomas', 'thames' or germanic
		        if MStringAt(source, current + 2, 2, "OM", "AM") _
		          OR MStringAt(source, 1, 4, "VAN ", "VON ") OR MStringAt(source, 1, 3, "SCH") then
		          out1 = out1 + "T"
		          out2 = out2 + "T"
		        else
		          out1 = out1 + "0"     // 0 represents "TH" sound in Metaphone
		          out2 = out2 + "T"     // (a bad choice -- # would have been better)
		        end if
		        current = current +  2
		        
		      else
		        out1 = out1 + "T"
		        out2 = out2 + "T"
		        if MStringAt(source, current + 1, 1, "T", "D") then
		          current = current +  2
		        else
		          current = current +  1
		        end if
		      end if
		      
		    case "V"
		      out1 = out1 + "F"
		      out2 = out2 + "F"
		      if charAt(current + 1) = "V" then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "W"
		      //can also be in middle of word
		      if MStringAt(source, current, 2, "WR") then
		        out1 = out1 + "R"
		        out2 = out2 + "R"
		        current = current +  2
		        
		      else
		        if current = 1 AND (MIsVowel(source, current + 1) OR MStringAt(source, current, 2, "WH")) then
		          //Wasserman should match Vasserman
		          if(MIsVowel(source, current + 1)) then
		            out1 = out1 + "A"
		            out2 = out2 + "F"
		          else
		            //need Uomo to match Womo
		            out1 = out1 + "A"
		            out2 = out2 + "A"
		          end if
		        end if
		        
		        if (current = length AND MIsVowel(source, current - 1)) _
		          OR MStringAt(source, current - 1, 5, "EWSKI", "EWSKY", "OWSKI", "OWSKY") _
		          OR MStringAt(source, 1, 3, "SCH") then
		          //Arnow should match Arnoff
		          out1 = out1 + ""
		          out2 = out2 + "F"
		          current = current + 1
		          
		        elseif MStringAt(source, current, 4, "WICZ", "WITZ") then
		          //polish e.g. 'filipowicz'
		          out1 = out1 + "TS"
		          out2 = out2 + "FX"
		          current = current + 4
		          
		        else
		          //else skip it
		          current = current + 1
		        end if
		      end if
		      
		    case "X"
		      //french e.g. breaux
		      if not (current = length AND _
		        (MStringAt(source, current - 3, 3, "IAU", "EAU") OR MStringAt(source, current - 2, 2, "AU", "OU"))) then
		        out1 = out1 + "KS"
		        out2 = out2 + "KS"
		      end if
		      
		      if MStringAt(source, current + 1, 1, "C", "X") then
		        current = current +  2
		      else
		        current = current +  1
		      end if
		      
		    case "Z"
		      //chinese pinyin e.g. 'zhao'
		      if charAt(current + 1) = "H" then
		        out1 = out1 + "J"
		        out2 = out2 + "J"
		        current = current +  2
		        
		      else
		        if MStringAt(source, current + 1, 2, "ZO", "ZI", "ZA") _
		          OR (SlavoGermanic AND current > 1 AND charAt(current - 1) <> "T") then
		          out1 = out1 + "S"
		          out2 = out2 + "TS"
		        else
		          out1 = out1 + "S"
		          out2 = out2 + "S"
		        end if
		        
		        if charAt(current + 1) = "Z" then
		          current = current +  2
		        else
		          current = current +  1
		        end if
		      end if
		      
		      // ----
		    else
		      // if none of the above cases, just skip this character
		      current = current + 1
		    end Select
		  wend
		  
		  outPrimary = out1
		  outAlternate = out2
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MIsVowel(source As String, atPos As Integer) As Boolean
		  // This is a private helper function for the Metaphone method.
		  
		  return InStr( "AEIOUY", Mid(source, atPos, 1) ) > 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function MStringAt(source As String, start As Integer, length As Integer, paramArray args As String) As Boolean
		  // This is a private helper function for the Metaphone method.
		  
		  if start < 1 then return false
		  Dim target As String
		  if start > source.Len then
		    target = " "
		  else
		    target = Mid(source, start, length)
		  end if
		  return (args.IndexOf(target) >= 0)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function NthFieldQuoted(src as string, sep as string, index as integer) As string
		  // Equivalent to RB's nthField() function, but respects quoted values
		  // Usage:
		  //    s = """Hello, Kitty"", ""One"", ""Two, Three"""
		  //    s1 = nthFieldQuoted(s, ",", 3)
		  // result: s1 = "Two, Three" (including the quotes!)
		  
		  // Easy cases: no occurrences of the separator, or no quotation marks
		  if InStr(src,sep)=0 then
		    if index=1 then
		      return src
		    else
		      return ""
		    end if
		  elseif InStr(src,"""")= 0 then
		    return NthField(src, sep, index)
		  end if
		  
		  // Harder case: both separator and quotation marks.
		  // We'll use InStrQuoted to help us out.
		  Dim sepLen As Integer = sep.Len
		  Dim startPos As Integer = 1
		  for i As Integer = 1 to index - 1
		    Dim sepPos As Integer = InStrQuoted( startPos, src, sep )
		    if sepPos = 0 then return ""
		    startPos = sepPos + sepLen
		  next
		  Dim sepPos As Integer = InStrQuoted( startPos, src, sep )
		  if sepPos = 0 then return Mid( src, startPos )
		  return Mid( src, startPos, sepPos - startPos )
		  
		  // Old code is below for posterity.
		  'dim c, n, startPos, endPos as integer
		  'dim inQuotes as boolean
		  'dim a as string
		  '
		  'dim sepLen as integer = sep.Len
		  'dim srcLen as integer = len( src )
		  'dim leftSep as string = left( sep, 1 )
		  'dim adjustedIndex as integer = (index -1)
		  '
		  
		  'endpos = srcLen + 1  // (accounts for final field, where no ending separator is found)
		  'for n=1 to srcLen
		  'a = Mid(src,n,1)
		  'if a= """" then
		  'inQuotes = not inQuotes
		  'elseif (a=leftSep) and not inQuotes then
		  'if mid(src, n, sepLen) = sep then
		  'c = c + 1
		  'if index = 1 then
		  '// First Field
		  'startPos = 1
		  'endPos = n-1
		  'exit
		  'else
		  '// Field 2..x
		  'if (c=adjustedIndex)  then
		  '// Leading Sep gefunden
		  'startPos = n+sepLen
		  'elseif (c = index) then
		  '// Trailing Sep found
		  'endPos = n-1
		  'exit
		  'end if
		  'end if
		  'end if
		  'end if
		  'next
		  '
		  'if startPos = 0 then
		  'if index=1 then
		  'return src
		  'else
		  'return ""
		  'end if
		  'else
		  'return mid(src,startPos,endPos-startPos+1)
		  'end if
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PadBoth(s as String, width as Integer, padding as String = " ") As String
		  // Pad a string to at least 'width' characters, by adding padding characters
		  // to the left and right sides of the string.
		  //
		  // If it is impossible to center the string, the string will be one character
		  // to the right more than it is to the left.
		  
		  dim length as Integer
		  
		  length = len(s)
		  if length >= width then return s
		  
		  dim mostToRepeat as Integer
		  mostToRepeat = ceil((width-length)/len(padding))
		  
		  dim repeated as String
		  repeated = Repeat(padding, ceil(mostToRepeat/2))
		  
		  return mid(repeated, 1, ceil((width-length)/2)) + s + mid(repeated,1,(width-length)\2)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PadLeft(s as String, width as Integer, padding as String = " ") As String
		  // Pad a string to at least 'width' characters, by adding padding characters
		  // to the left side of the string.
		  
		  dim length as Integer
		  length = len(s)
		  if length >= width then return s
		  
		  dim mostToRepeat as Integer
		  mostToRepeat = ceil((width-length)/len(padding))
		  return mid(Repeat(padding, mostToRepeat),1,width-length) + s
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PadRight(s as String, width as Integer, padding as String = " ") As String
		  // Pad a string to at least 'width' characters, by adding padding characters
		  // to the right side of the string.
		  
		  dim length as Integer
		  length = len(s)
		  if length >= width then return s
		  
		  dim mostToRepeat as Integer
		  mostToRepeat = ceil((width-length)/len(padding))
		  return s + mid(Repeat(padding, mostToRepeat),1,width-length)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Random(length As Integer = 8, charset As String = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ") As String
		  // Generate a random string of the specified length, by choosing characters
		  // at random from the given character set.  NOTE: we currently assume that
		  // all characters in charset are only 1 byte each.
		  
		  #if RBVersion > 5.5
		    static m As MemoryBlock
		  #else
		    Dim m As MemoryBlock
		  #endif
		  
		  if m = nil then
		    m = New MemoryBlock( length )
		  else
		    m.Size = length
		  end if
		  
		  Dim charsetSize As Integer = Len( charset )
		  
		  #if RBVersion > 5.5
		    static r As New Random
		    for i As Integer = 0 to length - 1
		      m.StringValue( i, 1 ) = MidB( charset, r.InRange( 1, charsetSize ), 1 )
		    next
		  #else
		    for i As Integer = 0 to length - 1
		      m.StringValue( i, 1 ) = MidB( charset, Rnd * charsetSize + 1, 1 )
		    next
		  #endif
		  
		  return DefineEncoding( m.StringValue( 0, length ), charset.Encoding )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Remove(s As String, charSet As String = " ") As String
		  // Delete all characters which are members of charSet. Example:
		  // Delete("wooow maaan", "aeiou") = "ww mn".
		  
		  Dim sLenB As Integer = s.LenB
		  if sLenB < 2 then return s
		  
		  Dim m As MemoryBlock
		  m = NewMemoryBlock( sLenB )
		  
		  charSet = ConvertEncoding( charSet, s.Encoding )
		  
		  Dim sLen As Integer = s.Len
		  
		  Dim char As String
		  Dim spos, mpos As Integer
		  for spos = 1 to sLen
		    char = Mid( s, spos, 1 )
		    if InStrB( charSet, char ) < 1 then
		      m.StringValue( mpos, char.LenB ) = char
		      mpos = mpos + char.LenB
		    end if
		  next
		  
		  return DefineEncoding( m.StringValue(0, mpos), s.Encoding )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Repeat(s as String, repeatCount as Integer) As String
		  // Concatenate a string to itself 'repeatCount' times.
		  // Example: Repeat("spam ", 5) = "spam spam spam spam spam ".
		  
		  #pragma disablebackgroundTasks
		  
		  if repeatCount <= 0 then return ""
		  if repeatCount = 1 then return s
		  
		  // Implementation note: normally, you don't want to use string concatenation
		  // for something like this, since that creates a new string on each operation.
		  // But in this case, we can double the size of the string on iteration, which
		  // quickly reduces the overhead of concatenation to insignificance.  This method
		  // is faster than any other we've found (short of declares, which were only
		  // about 2X faster and were quite platform-specific).
		  
		  Dim desiredLenB As Integer = LenB(s) * repeatCount
		  dim output as String = s
		  dim cutoff as Integer = (desiredLenB+1)\2
		  dim curLenB as Integer = LenB(output)
		  
		  while curLenB < cutoff
		    output = output + output
		    curLenB = curLenB + curLenB
		  wend
		  
		  output = output + LeftB(output, desiredLenB - curLenB)
		  return output
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ReplaceRange(s As String, start As Integer, length As Integer, newText As String) As String
		  // Replace a part of the given string with a new string.
		  
		  return Left(s, start-1) + newText + Mid(s, start + length)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ReplaceRangeB(s As String, startB As Integer, lengthB As Integer, newText As String) As String
		  // Replace a part of the given string with a new string
		  // (with offset and length in bytes rather than characters).
		  
		  return LeftB(s, startB-1) + newText + MidB(s, startB + lengthB)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Reverse(s As String) As String
		  // Return s with the characters in reverse order.
		  
		  if Len(s) < 2 then return s
		  
		  Dim characters() as String = Split( s, "" )
		  Dim leftIndex as Integer = 0
		  Dim rightIndex as Integer = UBound(characters)
		  #pragma BackgroundTasks False
		  While leftIndex < rightIndex
		    Dim temp as String = characters(leftIndex)
		    characters(leftIndex) = characters(rightIndex)
		    characters(rightIndex) = temp
		    leftIndex = leftIndex + 1
		    rightIndex = rightIndex - 1
		  Wend
		  Return Join( characters, "" )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ReverseB(s As String) As String
		  // Return s with the bytes in reverse order.
		  // Note that if s is text in any encoding that may have
		  // multi-byte characters, you should probably be using
		  // Reverse instead of ReverseB.
		  
		  if LenB(s) < 2 then return s
		  
		  #pragma BackgroundTasks False
		  Dim m as MemoryBlock = s
		  Dim leftIndex as Integer = 0
		  Dim rightIndex as Integer = m.Size - 4
		  While leftIndex + 4 < rightIndex
		    Dim tempL as Integer = m.Long(leftIndex)
		    Dim tempR as Integer = m.Long(rightIndex)
		    m.LittleEndian = not m.LittleEndian
		    m.Long(leftIndex) = tempR
		    m.Long(rightIndex) =tempL
		    m.LittleEndian = not m.LittleEndian
		    leftIndex = leftIndex + 4
		    rightIndex = rightIndex - 4
		  Wend
		  rightIndex = rightIndex + 3
		  While leftIndex < rightIndex
		    Dim temp as Byte = m.Byte(leftIndex)
		    m.Byte(leftIndex) = m.Byte(rightIndex)
		    m.Byte(rightIndex) = temp
		    leftIndex = leftIndex + 1
		    rightIndex = rightIndex - 1
		  Wend
		  Return DefineEncoding(m, Encoding(s))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function RTrim(source As String, charsToTrim As String) As String
		  // This is an extended version of RB's RTrim function that lets you specify
		  // a set of characters to trim.
		  
		  Dim srcLen As Integer = source.Len
		  Dim rightPos, i As Integer
		  for i = srcLen DownTo 1
		    if InStr( charsToTrim, Mid(source, i, 1) ) = 0 then exit
		  next
		  rightPos = i
		  
		  return Mid( source, 1, rightPos )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Soundex(s As String, stripPrefix As Boolean = true) As String
		  // Return the Soundex code for the given string.
		  // That's the first character, followed by numeric
		  // codes for the first several consonants.
		  // For more detail, see: <http://www.searchforancestors.com/soundex.html>
		  
		  Dim prefix, prefixes(-1) As String
		  Dim i, prefixLen As Integer
		  
		  s = Trim( s )
		  
		  if stripPrefix then
		    prefixes = Array("La ", "De ", "Van ")  // more to come?
		    for each prefix in prefixes
		      prefixLen = prefix.Len
		      if Left( s, prefixLen ) = prefix then
		        s = Mid( s, prefixLen+1 )
		        exit
		      end if
		    next
		  end if
		  
		  Dim c, out As String
		  out = Uppercase( Left(s, 1) )
		  Dim sLen, curCode, lastCode As Integer
		  sLen = s.Len
		  for i = 2 to sLen
		    c = Uppercase( Mid( s, i, 1 ) )
		    if InStrB( "BPFV", c ) > 0 then
		      curCode = 1
		    elseif InStrB( "CSKGJQXZ", c ) > 0 then
		      curCode = 2
		    elseif InStrB( "DT", c ) > 0 then
		      curCode = 3
		    elseif c = "L" then
		      curCode = 4
		    elseif InStrB( "MN", c ) > 0 then
		      curCode = 5
		    elseif c = "R" then
		      curCode = 6
		    else
		      curCode = 0
		    end if
		    if curCode > 0 and curCode <> lastCode then
		      out = out + str(curCode)
		      if Len(out) = 4 then return out
		    end if
		  next
		  
		  return Left( out + "000", 4 )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitByLength(s As String, fieldWidth As Integer) As String()
		  // Split a string into fields, each containing 'fieldWidth' characters
		  // (except for the last one, which may have fewer).
		  
		  if fieldWidth < 1 then   // fieldWidth must be >= 1
		    raise New OutOfBoundsException
		  end if
		  
		  Dim out(-1) As String
		  
		  Dim qty As Integer
		  qty = Ceil( Len(s) / fieldWidth )
		  Redim out( qty - 1 )
		  
		  Dim pos, i As Integer
		  pos = 1
		  for i = 0 to qty-1
		    out(i) = Mid( s, pos, fieldWidth )
		    pos = pos + fieldWidth
		  next
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitByLengthB(s As String, fieldWidth As Integer) As String()
		  // Split a string into fields, each containing 'fieldWidth' bytes
		  // (except for the last one, which may have fewer).
		  
		  if fieldWidth < 1 then   // fieldWidth must be >= 1
		    raise New OutOfBoundsException
		  end if
		  
		  Dim out(-1) As String
		  
		  Dim qty As Integer
		  qty = Ceil( LenB(s) / fieldWidth )
		  Redim out( qty - 1 )
		  
		  Dim pos, i As Integer
		  pos = 1
		  for i = 0 to qty-1
		    out(i) = MidB( s, pos, fieldWidth )
		    pos = pos + fieldWidth
		  next
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitByRegEx(source As String, delimPattern As String) As String()
		  // Split a string into fields delimited by a regular expression.
		  
		  Dim out(-1) As String
		  
		  Dim re As New RegEx
		  Dim rm As RegExMatch
		  Dim startPos As Integer
		  
		  re.SearchPattern = delimPattern
		  rm = re.Search( source )
		  while rm <> nil
		    'System.DebugLog rm.SubExpressionString(0) + " at " + str(rm.SubExpressionStartB(0)) + " matches " + pattern + " in " + s
		    out.Append MidB( source, startPos + 1, rm.SubExpressionStartB(0) - startPos )
		    startPos = re.SearchStartPosition
		    rm = re.Search
		  wend
		  
		  if startPos < source.LenB then
		    out.Append MidB( source, startPos + 1 )
		  end if
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitQuoted(source As String, delimiter As String, removeQuotes as Boolean) As String()
		  // Same as RB's Split, except that it respects quoted strings.
		  
		  Dim out() As String
		  
		  // Easy cases: no occurrences of the separator, or no quotation marks
		  if InStr( source, delimiter )=0 then
		    out = Array( source )
		  elseif InStr( source, """" )= 0 then
		    out = Split( source, delimiter )
		  else
		    // Harder case: both separator and quotation marks.
		    // We'll use InStrQuoted to help us out.
		    Dim delimiterLen As Integer = delimiter.Len
		    Dim startPos As Integer = 1
		    do
		      Dim delimiterPos As Integer = InStrQuoted( startPos, source, delimiter )
		      if delimiterPos = 0 then
		        // no more separators
		        out.Append Mid( source, startPos )
		        exit
		      end if
		      out.Append Mid( source, startPos, delimiterPos - startPos )
		      startPos = delimiterPos + delimiterLen
		    loop
		  end if
		  
		  // Finally, strip quotes if asked to do so.
		  // (Including the case of a start quote without an end quote, which
		  // could happen in the last field.)
		  if removeQuotes then
		    for i As Integer = 0 to UBound( out )
		      Dim s As String = out(i)
		      if s.LeftB( 1 ) = """" then
		        if s.RightB( 1 ) = """" then
		          out(i) = MidB( s, 2, s.LenB - 2 )
		        else
		          out(i) = MidB( s, 2 )
		        end if
		      end if
		    next
		  end if
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitToCDbl(source As String, delimiter As String = " ") As Double()
		  // Split a string into fields, then convert each field into a Double
		  // using the CDbl function.  This is appropriate for a set of numbers
		  // entered or readable by the end-user.
		  
		  Dim fields(-1) As String
		  fields = source.Split(delimiter)
		  
		  Dim out(-1) As Double
		  Redim out( UBound(fields) )
		  
		  Dim i As Integer
		  for i = UBound(fields) DownTo 0
		    out(i) = CDbl( fields(i) )
		  next
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitToInt(source As String, delimiter As String = " ") As Integer()
		  // Split a string into fields, then convert each field into an Integer
		  // using the Val function.
		  
		  Dim fields(-1) As String
		  fields = source.Split(delimiter)
		  
		  Dim out(-1) As Integer
		  Redim out( UBound(fields) )
		  
		  Dim i As Integer
		  for i = UBound(fields) DownTo 0
		    out(i) = Val( fields(i) )
		  next
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SplitToVal(source As String, delimiter As String = " ") As Double()
		  // Split a string into fields, then convert each field into a Double
		  // using the Val function.  This is appropriate for a set of numbers
		  // used only by the computer; for human-readable numbers, consider
		  // using SplitToCDbl instead.
		  
		  Dim fields(-1) As String
		  fields = source.Split(delimiter)
		  
		  Dim out(-1) As Double
		  Redim out( UBound(fields) )
		  
		  Dim i As Integer
		  for i = UBound(fields) DownTo 0
		    out(i) = Val( fields(i) )
		  next
		  
		  return out
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1, CompatibilityFlags = TargetHasGUI
		Protected Function Sprintf(src as string, ParamArray data as Variant) As string
		  // Returns a string produced according to the formatting string <src>.
		  // The format string <src> is composed of zero or more directives: ordinary
		  // characters (excluding %) that are
		  // copied directly to the result, and conversion
		  // specifications, each of which results in fetching its
		  // own parameter.
		  // For details, see http://de.php.net/manual/en/function.sprintf.php
		  
		  // Attention: This function differs from the PHP sprintf() function in that
		  // it formats floating numbers according to the locale settings.
		  // For example, in Germany,
		  //    sprintf("%04.2f", 123.45)
		  // will return "0123,45".
		  
		  // Written by Frank Bitterlich, bitterlich@gsco.de
		  // Additional work by Florent Pillet, florent@florentpillet.com
		  
		  // NOTE: This method is currently available only to GUI apps due
		  // to <http://www.realsoftware.com/feedback/viewreport.php?reportid=owsxeqnf>.
		  // Once that bug is fixed, we can make this available to console apps too.
		  
		  dim rex as new RegEx
		  dim match as RegExMatch
		  dim argtype, padding, alignment, precstr, replacement, frmstr, s as string
		  dim p, width, precision, index, start, length as integer
		  dim vf as double
		  dim datum As Variant
		  
		  rex.SearchPattern = "(%)(0|/s|'.)?(-)?(\d*)(\.\d+)?([%bcdeufosxX])"
		  rex.Options.Greedy = true
		  match = rex.Search(src)
		  index = -1
		  
		  do until match = nil
		    if match.SubExpressionCount = 7 then
		      padding = Right(" " + match.SubExpressionString(2), 1)
		      // if padding = "" then padding = " " // default: space
		      alignment = match.SubExpressionString(3)
		      width = Val(match.SubExpressionString(4))
		      precstr = Mid(match.SubExpressionString(5), 2)
		      precision = Val(precstr)
		      if precstr="" then precision = 6
		      
		      argtype = match.SubExpressionString(6)
		      if argtype <> "%" then
		        index = index + 1
		        if index > UBound(data) then
		          datum = 0
		        else
		          datum = data(index)
		        end if
		      end if
		      
		      select case argtype
		      case "%"
		        replacement = "%"
		        
		      case "b" // binary int
		        replacement = bin(datum)
		        
		      case "c" // character
		        replacement = Encodings.UTF8.Chr(datum)
		        width = 0
		        
		      case "d" // signed int
		        if padding = "0" then
		          frmstr = "-"+Repeat("0", width)
		          if datum<0 then frmstr = Left(frmstr, Len(frmstr)-1)
		        else
		          frmstr = "-#"
		        end if
		        replacement = Format(datum, frmstr)
		        
		      case "e" // scientific notation
		        vf = datum
		        frmstr = "-#."+Repeat("0", precision)+"e+"
		        Replacement = Format(vf, frmstr)
		        p = InStr(Replacement, "e")
		        // Make sure the part after the "e" has two digits
		        Replacement = Left(Replacement, p)+Format(Val(Mid(Replacement, p+1)), "+00")
		        
		      case "u" // unsigned int
		        replacement = Format(datum, "#")
		        
		      case "f" // signed float
		        if padding = "0" then
		          frmstr = "-"+Repeat("0", width)
		          if datum<0 then frmstr = Left(frmstr, Len(frmstr)-1)
		        else
		          frmstr = "-#"
		        end if
		        if precision > 0 then
		          frmstr = frmstr + "." + Repeat("0", precision)
		        end if
		        Replacement = Format(datum, frmstr)
		        if precision > 0 and padding<>"0" then width = width + precision + 1
		        
		      case "o" // octal int
		        replacement = Oct(datum)
		        
		      case "s" // string
		        replacement = datum
		        
		      case "x" // hex int; uppercase "X" means uppercase hex, "x" is lowercase hex
		        replacement = hex(datum)
		        if asc(argtype) = &h58 then // uppercase "X"
		          replacement = Uppercase(replacement)
		        else // lowercase "x"
		          replacement = lowercase(replacement)
		        end if
		      end select
		      
		      if width>Len(replacement) then
		        if alignment="-" then // align left
		          replacement=replacement+Repeat(padding, width-Len(replacement))
		        else // align right
		          replacement=Repeat(padding, width-Len(replacement))+replacement
		        end if
		      end if
		    end if
		    start = match.SubExpressionStartB(0)+1
		    length = LenB(match.SubExpressionString(0))
		    
		    s = LeftB(src, start-1) + replacement
		    src = s + MidB(src, start+length)
		    
		    match = rex.Search(src, Len(s))
		  loop
		  
		  return src
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SQLify(s As String) As String
		  // Return a version of s ready for use in an SQL statement.
		  
		  // In other words, we just need to double the apostrophes:
		  return ReplaceAll( s, "'", "''" )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Squeeze(s As String, charSet As String = " ") As String
		  // Find any repeating characters, where the character is a member of
		  // charSet, and replace the run with a single character.  Note that this
		  // uses standard RB text matching, and so is case-insensitive; all
		  // runs will be replaced with the case as given in charSet.  If you need
		  // case-sensitive squeezing, use SqueezeB.
		  // Example: Squeeze("woOow mAAAn", "aeiou") = "wow man".
		  
		  Dim sLenB As Integer = s.LenB
		  if sLenB < 2 then return s
		  
		  #pragma BackgroundTasks False
		  
		  charSet = ConvertEncoding( charSet, s.Encoding )
		  for each char as String in Split( charSet, "" )
		    Dim doubleChar As String = char + char
		    do
		      s = ReplaceAll( s, doubleChar, char )
		      Dim newLenB As Integer = s.LenB
		      if newLenB = sLenB then exit
		      sLenB = newLenB
		    loop
		  next char
		  
		  return s
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SqueezeB(s As String, charSet As String = " ") As String
		  // Find any repeating characters, where the character is a member of
		  // charSet, and replace the run with a single character.  Note that this
		  // uses binary matching (and so is case sensitive), but does convert
		  // charSet into the encoding of s, so you don't have to worry about that.
		  // Example: Squeeze("wooow maaan", "aeiou") = "wow man".
		  
		  Dim sLenB As Integer = s.LenB
		  if sLenB < 2 then return s
		  
		  #pragma BackgroundTasks False
		  
		  charSet = ConvertEncoding( charSet, s.Encoding )
		  for each char as String in Split( charSet, "" )
		    Dim doubleChar As String = char + char
		    do
		      s = ReplaceAllB( s, doubleChar, char )
		      Dim newLenB As Integer = s.LenB
		      if newLenB = sLenB then exit
		      sLenB = newLenB
		    loop
		  next char
		  
		  return s
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function StartsWith(extends s As String, withWhat As String) As Boolean
		  // Return true if 's' starts with the string 'withWhat',
		  // doing a standard string comparison.
		  
		  return Left(s, withWhat.Len) = withWhat
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function StartsWithB(extends s As String, withWhat As String) As Boolean
		  // Return true if 's' starts with the string 'withWhat',
		  // doing a binary comparison.
		  
		  return StrComp( LeftB(s, withWhat.Len), withWhat, 0 ) = 0
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SwapBytePairs(s As String) As String
		  // Swap every pair of bytes in the given string.  This can be useful
		  // when dealing with UTF-16 data with the wrong byte order.
		  
		  Dim inm As MemoryBlock = s
		  inm.LittleEndian = true
		  Dim outm As New MemoryBlock( inm.Size )
		  outm.LittleEndian = false
		  Dim maxi As Integer = inm.Size - 2
		  for i As Integer = 0 to maxi step 2
		    outm.UInt16Value(i) = inm.UInt16Value(i)
		  next
		  return outm.StringValue( 0, outm.Size, s.Encoding )
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SwapByteQuads(s As String) As String
		  // Swap the order of every 4 bytes in the given string.  This can be useful
		  // when dealing with UTF-32 data with the wrong byte order.
		  
		  Dim inm As MemoryBlock = s
		  inm.LittleEndian = true
		  Dim outm As New MemoryBlock( inm.Size )
		  outm.LittleEndian = false
		  Dim maxi As Integer = inm.Size - 4
		  for i As Integer = 0 to maxi step 4
		    outm.UInt32Value(i) = inm.UInt32Value(i)
		  next
		  return outm.StringValue( 0, outm.Size, s.Encoding )
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ThousandsSeparator() As String
		  // Return the thousands separator the user uses (either "." or ",").
		  if mThousandsSeparator = "" then
		    mThousandsSeparator = Format(1000, "#,#")
		    mThousandsSeparator = Mid( mThousandsSeparator, 2, 1 )
		  end if
		  
		  return mThousandsSeparator
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Trim(source As String, charsToTrim As String) As String
		  // This is an extended version of RB's Trim function that lets you specify
		  // a set of characters to trim.
		  
		  Dim srcLen As Integer = source.Len
		  Dim leftPos, i As Integer
		  for i = 1 to srcLen
		    if InStr( charsToTrim, Mid(source, i, 1) ) = 0 then exit
		  next
		  leftPos = i
		  if leftPos > srcLen then return ""
		  
		  Dim rightPos As Integer
		  for i = srcLen DownTo 1
		    if InStr( charsToTrim, Mid(source, i, 1) ) = 0 then exit
		  next
		  rightPos = i
		  
		  return Mid( source, leftPos, rightPos - leftPos + 1 )
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub WrapLines(lines() As String, charsPerLine As Integer = 72, paragraphFill As Boolean = true)
		  // Wrap the text so that no line is longer than charsPerLine.  If paragraphFill
		  // is true, then whenever one long line is followed by a line that does not
		  // start with whitespace, join them together into one continuous paragraph.
		  
		  if UBound( lines ) < 0 then return
		  
		  // Start by joining lines, if called for.
		  if paragraphFill then
		    Dim lineNum As Integer = 1
		    Dim lastLineShort As Boolean = ( lines(0).Len < charsPerLine-20 )
		    while lineNum <= UBound( lines )
		      Dim line As String = lines( lineNum )
		      Dim firstChar As String = Left( line, 1 )
		      if lastLineShort then
		        // last line was short, so don't join this one to it
		        lineNum = lineNum + 1
		      elseif line = "" or firstChar <= " " or firstChar = ">" or firstChar = "|" then
		        // this line is empty or starts with whitespace or other special char; don't join it
		        lineNum = lineNum + 1
		      else
		        // this line starts with a character; join it to the previous line
		        lines( lineNum - 1 ) = lines( lineNum - 1 ) + " " + line
		        lines.Remove lineNum
		      end if
		      lastLineShort = ( line.Len < charsPerLine-20 )
		    wend
		  end if
		  
		  // Then, go through and do the wrapping.
		  for lineNum As Integer = 0 to UBound( lines )
		    Dim line As String = RTrim( lines( lineNum ) )
		    if line.Len <= charsPerLine then
		      lines( lineNum ) = line
		    else
		      Dim breakPos As Integer
		      for breakPos = charsPerLine DownTo 1
		        Dim c As String = Mid( line, breakPos, 1 )
		        if c <= " " or c = "-" then exit
		      next
		      if breakPos < 2 then breakPos = charsPerLine + 1  // no point breaking before char 1
		      lines.Insert lineNum + 1, LTrim( Mid( line, breakPos ) )
		      lines( lineNum ) = LTrim( Left( line, breakPos - 1) )
		    end if
		  next
		  
		End Sub
	#tag EndMethod


	#tag Note, Name = Contributors
		
		Many people have contributed to the development of this module, including:
		
		Frank Bitterlich, bitterlich@gsco.de
		Ben Blake, benb@kagi.com
		Cortis Clark, cortis@realsoftware.com
		Carlos Martinho, rblists@rbtips.com
		Jon Johnson, jonj@realsoftware.com
		Joe Strout, joe@strout.net (*)
		Florent Pillet, florent@florentpillet.com
		Kem Tekinay, ktekinay@mactechnologies.com
		Charles Yeomans, yeomans@desuetude.com
		
		(*) To whom correspondence should be addressed.
	#tag EndNote

	#tag Note, Name = Home Page
		
		This StringUtils module is maintained by Verified Express, LLC.
		The current administrative contact for it is Joe Strout (joe@strout.net).
		You should be able to find the latest version via this URL:
		
		    http://www.verex.com/opensource/
	#tag EndNote

	#tag Note, Name = License
		
		This StringUtils module is in the public domain.  You may use it for any purpose
		whatsoever, but it comes with no express or implied warranty of correctness or
		fitness for any purpose.
		
		Share and enjoy!
	#tag EndNote

	#tag Note, Name = Most Wanted
		Features we'd really like to have, but haven't yet found the time to implement:
		
		...none at the moment!
	#tag EndNote

	#tag Note, Name = Version History
		
		2004-JUL-17: version 1.0
		- First public release.
		
		2004-JUL-22: version 1.1
		- Fixed some (harmless) warnings in CountRegEx and Repeat.
		- Added ControlCharacters.
		- Added DecimalSeparator.
		- Added ThousandsSeparator.
		- Added SplitByLength and SplitByLengthB.
		- Added Sprintf.
		- Added Trim, LTrim, and RTrim (with charsToTrim parameter).
		- Improved the speed of CountFieldsQuoted substantially.
		- Improved the speed of NthFieldQuoted and Squeeze slightly.
		
		2004-SEP-23: version 1.2
		- Fixed a bug in the TestSplitToVal unit test.
		- Fixed a bug in EditDistance causing incorrect results in some cases.
		- Improved the speed of EditDistance substantially.
		- Changed EditDistance to be case-sensitive.
		
		2005-AUG-05: version 1.3
		- Improved speed of InStrReverseB.
		- Fixed a bug in Sprintf that caused incorrect results when non-ASCII
		  characters were used in the format string.
		
		2006-MAR-13: version 1.4
		- Improved speed of HexB slightly.
		- Added MatchCase function.
		- Fixed a bug in NthFieldQuoted which would cause it to fail when 
		   the separator occurred as the first character of the string.
		- Added Random function.
		- Made all functions available to console apps except Sprintf.
		
		2006-APR-10: version 1.5
		- Improved speed of HexB substantially.
		- Adjusted Random so that it should work in 5.5 and earlier (untested).
		- Added SplitQuoted and JoinQuoted functions.
		- Added EncodeCase and DecodeCase functions.
		
		2006-JUN-29: version 1.6
		- Fixed buggy handling of "%%" in Sprintf, and made it so that unspecified
		   values are output as if you had specified 0; also changed %c to output
		   a Unicode character for the given code point.
		- Improved speed of Reverse and ReverseB (thanks to Charles Yeomans).
		   Note that I'm not sure whether the new code works in 5.5; somebody who
		   has 5.5 is going to have to start helping with that (if anyone cares).
		- Added new LineEnding function, that returns the first line ending (of
		   any standard sort) that is found in the string.
		
		2006-AUG-08: version 1.7
		- Added InStrQuoted function.
		- Rewrote SplitQuoted in terms of InStrQuoted.  This fixed several bugs
		   in the previous implementation, including an infinite loop and an
		   OutOfBoundsException that could occur on certain inputs.
		- Rewrote NthFieldQuoted in terms of InStrQuoted, just to make the code
		   simpler and more consistent.  This should cause no changes in behavior.
		
		2006-SEP-06: version 1.8
		- Fixed a bug in EndsWithB that would cause it to fail on multi-byte characters.
		- Added versions of Chop and ChopB that specify a string to chop off, rather
		   than a number of characters.
		- Dramatically improved the speed of Squeeze.  NOTE: Also changed the
		   behavior slightly.  Squeeze now uses standard, non-binary text matching.
		- A new function, SqueezeB, is now provided for binary squeezing.  Both
		   versions are substantially faster than the old one, which always did binary,
		   but existing code may want to switch to SqueezeB.
		
		2007-MAR-15: version 1.9
		- Made Chop and ChopB work in RB2006R4 as they did in previous versions,
		  returning an empty string when told to chop more than the total number
		  of chars or bytes.  (Previous code relied on an undefined behavior of Left
		  and LeftB.)
		- Added the WrapLines function.
		- Improved speed of the Repeat function, especially on large repeat counts.
		- Added the GuessEncoding functions.
		- Added the SwapBytePairs and SwapByteQuads functions.
	#tag EndNote


	#tag Property, Flags = &h21
		Private mControlChars As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mDecimalSeparator As String
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mThousandsSeparator As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
