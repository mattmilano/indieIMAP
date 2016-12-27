#tag Module
Protected Module modUtilities
	#tag Method, Flags = &h0
		Function generateMD5HASH(key As string, data As string) As string
		  Dim result, bytes, hex as String
		  Dim hexResult as String
		  Dim i as Integer
		  
		  
		  // We do this in ASCII because the well known test vectors come in ASCII
		  key = ConvertEncoding(key,Encodings.ASCII)
		  data = ConvertEncoding(data,Encodings.ASCII)
		  
		  bytes = key
		  hex = encodeHex(data)
		  
		  
		  // Convert to HEX
		  For i = 1 to 16
		    hexResult = hexResult + Right("0"+Hex(Asc(Mid(result,i,1))),2)
		  next
		  
		  return hexResult
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub writeIMAPLog(log As string)
		  dim f As FolderItem
		  dim output As TextOutputStream
		  dim theDate As new date
		  
		  theDate = new date
		  
		  f = GetFolderItem("Logs")
		  if f = nil or not f.Exists then
		    f.CreateAsFolder
		  end if
		  
		  f = f.Child("indieMAP Log.txt")
		  if f = nil or not f.Exists then
		    output = f.CreateTextFile
		    output.Close
		  else
		    output = f.AppendToTextFile
		    output.Write theDate.ShortDate+" "+theDate.ShortTime+" - "+log+EndOfLine+EndOfLine
		    output.Close
		  end if
		End Sub
	#tag EndMethod


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
