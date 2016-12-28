#tag Class
Protected Class IMAPSocket
Inherits SSLSocket
	#tag Event
		Sub DataAvailable()
		  dim s, response, lines() As string
		  
		  s = me.ReadAll
		  s = trim(s)
		  lines = split(s, EndOfLine.Windows)
		  
		  if myCommand = "CONNECT" then
		    myTag = left(s, 1)
		  end if
		  
		  
		  if instr(lines(0), ". NO") > 0 or instr(lines(0), "error") > 0 or instr(lines(0), ". BAD") > 0 then
		    handleError(me.ReadAll)
		    
		  else
		    
		    if myCommand = "UID" then
		      bufferString = bufferString + s
		    else
		      s = TrimResponse(s)
		      bufferData.Append s
		    end if
		    
		    dim i As integer
		    i = UBound(lines)
		    if left(lines(i), 5) = ". OK " or left(lines(i), 4) = "+ OK" or myCommand = "Connect" or myCommand = "LOGIN" or myCommand = "PLAIN" or myCommand = "AUTHENTICATE CRAM-MD5" then
		      
		      if myCommand = "UID" then
		        response = bufferString
		        bufferString = ""
		      else
		        response = join(bufferData, EndOfLine.Windows)
		        ReDim bufferData(-1)
		      end if
		      
		      handleResponse(response)
		    end if
		    
		  end if
		End Sub
	#tag EndEvent

	#tag Event
		Sub SendComplete(UserAborted As Boolean)
		  writeIMAPLog "Sending command: "+myCommand
		  
		  dim s As string
		  s = me.Lookahead
		End Sub
	#tag EndEvent


	#tag Method, Flags = &h0
		Sub AppendMessage(Mailbox As String, Index As Integer, EmailSource As String, EmailSize As Integer)
		  dim s As string
		  
		  
		  me.myCommand = "APPEND"
		  s = ". APPEND "+Mailbox+" (\Seen) {"+str(EmailSize)+"}"+EndOfLine.Windows+EmailSource
		  currentEmail = index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ChangeFlags(Index As Integer)
		  dim s As string
		  
		  
		  me.myCommand = "STORE"
		  s = ". UID STORE "+str(Index)+" +FLAGS ("+mySecondary+")"
		  currentEmail = Index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function CleanUpResponse(Response As String) As String
		  dim i As integer
		  dim lines() As String
		  
		  i = instr(Response, ")"+EndOfLine.Windows+". OK") - 1
		  
		  if i <> -1 then
		    Response = mid(Response, 1, i)
		    Response = Trim(Response)
		  end if
		  
		  
		  lines = split(Response, EndOfLine.Windows)
		  i = UBound(lines)
		  if instr(lines(i), ". OK") > 0 then
		    lines.Remove i
		    Response = join(lines, EndOfLine.Windows)
		  end if
		  
		  if myCommand = "FETCH BODY" or myCommand = "FETCH EMAIL" then
		    lines = split(Response, EndOfLine.Windows)
		    i = UBound(lines)
		    if instr(lines(i), "* ") > 0 and instr(lines(i), "FETCH") > 0 and instr(lines(i), "FLAGS") > 0 then
		      lines.Remove i
		      Response = join(lines, EndOfLine.Windows)
		    end if
		  end if
		  
		  return Response
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub CopyMessage(Index As Integer, Destination As String)
		  dim s As string
		  
		  if instr(destination, " ") > 0 then
		    destination = """"+destination+""""
		  end if
		  
		  me.myCommand = "COPY"
		  s = ". UID COPY "+str(Index)+" "+Destination
		  currentEmail = index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub DeleteMessage(Index As Integer)
		  dim s As string
		  
		  
		  me.myCommand = "DELETE"
		  s = ". UID STORE "+str(Index)+" +FLAGS (\Deleted)"
		  currentEmail = Index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub ExpungeMessages()
		  dim s As string
		  
		  
		  me.myCommand = "EXPUNGE"
		  s = ". EXPUNGE"
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FetchBody(Index As Integer)
		  dim s As string
		  
		  
		  me.myCommand = "FETCH BODY"
		  s = ". UID FETCH "+str(Index)+" BODY[]"
		  currentEmail = index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FetchHeaders(Index As integer)
		  dim s As string
		  
		  
		  me.myCommand = "FETCH HEADER"
		  's = ". UID FETCH "+str(Index)+" RFC822.HEADER"
		  s = ". UID FETCH "+str(Index)+" BODY.PEEK[HEADER]"
		  currentEmail = index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FetchMessage(Index As integer)
		  dim s As string
		  
		  
		  me.myCommand = "FETCH EMAIL"
		  s = ". UID FETCH "+str(Index)+" BODY.PEEK[]"
		  currentEmail = index
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub FetchUID(Range As String)
		  dim s As string
		  
		  myCommand = "UID"
		  s = ". UID Fetch "+Range+" FLAGS"
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub handleError(Error As string)
		  ProtocolError(Error)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub handleResponse(Response As string)
		  dim s As string
		  s = myCommand
		  
		  LoggedCommunication Response
		  
		  Select Case myCommand
		    
		  Case "Connect", "CAPABILITY"
		    
		    if instr(Response, "STARTTLS") > 0 then
		      writeIMAPLog "This server requires STARTTLS. Starting the TLS Negotiation."
		      ProgressChange("Sending STARTTLS")
		      me.myCommand = "STARTTLS"
		      SendCommand "STARTTLS"
		      Return
		    end if
		    
		    if instr(me.Address, "mac.com") > 0 or instr(me.Address, "icloud.com") > 0 or instr(me.Address, "me.com") > 0 then
		      
		      writeIMAPLog "Proceeding with AUTH LOGIN Authorization."
		      s = me.UserName+" "+me.Password
		      me.myCommand = "LOGIN"
		      SendCommand ". LOGIN "+s
		      Return
		      
		    elseif instr(me.Address, "gmail.com") > 0 and instr(Response, "xoauth") > 0 then
		      
		      writeIMAPLog "Proceeding with AUTH LOGIN Authorization."
		      s = me.UserName+" "+me.Password
		      me.myCommand = "LOGIN"
		      SendCommand ". LOGIN "+s
		      Return
		      
		    elseif instr(Response, "CRAM-MD5") > 0 then
		      
		      writeIMAPLog "Proceeding with CRAM-MD5 Authorization."
		      me.myCommand = "AUTHENTICATE CRAM-MD5"
		      SendCommand ". AUTHENTICATE CRAM-MD5"
		      Return
		      
		    elseif instr(Response, "LOGIN") > 0 then
		      
		      writeIMAPLog "Proceeding with AUTH LOGIN Authorization."
		      s = me.UserName+" "+me.Password
		      me.myCommand = "LOGIN"
		      SendCommand ". LOGIN "+s
		      Return
		      
		    elseif instr(Response, "PLAIN") > 0 then
		      
		      writeIMAPLog "Proceeding with AUTH PLAIN Authorization."
		      s = chr(0)+me.Username+chr(0)+me.Password
		      me.myCommand = "PLAIN"
		      SendCommand ". AUTHENTICATE PLAIN "+s
		      Return
		      
		      
		    else
		      
		      writeIMAPLog "No AUTH information. Proceed with CAPABILITY."
		      s = "CAPABILITY"
		      me.myCommand = "CAPABILITY"
		      SendCommand ". CAPABILITY"
		      Return
		      
		    end if
		    
		    
		  Case "PLAIN"
		    s = chr(0)+me.Username+chr(0)+me.Password
		    s = EncodeBase64(s)
		    me.myCommand = "LOGIN"
		    SendCommand s
		    return
		    
		    
		  Case "AUTHENTICATE CRAM-MD5"
		    dim i As integer
		    dim temp As string
		    i = instr(Response, " ")+1
		    s = mid(Response, i)
		    s = DecodeBase64(s)
		    
		    temp = Lowercase(generateMD5HASH(me.Password, s))
		    
		    s = EncodeBase64(me.Username+" "+temp)
		    
		    me.myCommand = "LOGIN"
		    SendCommand s
		    Return
		    
		    
		  Case "LOGIN"
		    writeIMAPLog "Retrieving Mailbox List."
		    me.myCommand = "LIST"
		    s = ". LIST "+""""" "+"""*"""
		    SendCommand s
		    Return
		    
		    
		  Case "LIST"
		    
		    Response = CleanUpResponse(Response)
		    processMailboxLIST(Response)
		    
		    
		  Case "SELECT"
		    
		    MailboxSelected(Response)
		    
		  Case "UID"
		    
		    Response = CleanUpResponse(Response)
		    UIDSReceived(Response)
		    
		  Case "FETCH HEADER"
		    
		    Response = CleanUpResponse(Response)
		    HeadersReceived(currentEmail, Response)
		    
		  Case "FETCH BODY"
		    
		    Response = CleanUpResponse(Response)
		    BodyReceived(currentEmail, Response)
		    
		  Case "FETCH EMAIL"
		    
		    Response = CleanUpResponse(Response)
		    MessageReceived(currentEmail, Response)
		    
		  CASE "STORE"
		    
		    FlagChanged(currentEmail, Response)
		    
		  CASE "COPY"
		    
		    MessageCopied(currentEmail, Response)
		    
		  Case "APPEND"
		    
		    MessageAppended(currentEmail, Response)
		    
		  Case "DELETE"
		    
		    MessageDeleted(currentEmail, Response)
		    
		  Case "EXPUNGE"
		    
		    MessagesExpunged(Response)
		    
		  Case "LOGOUT"
		    
		    me.Disconnect
		    
		  End Select
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Initialize()
		  myCRLF = EndOfLine.Windows
		  
		  if myCommand <> "" then
		    mySecondary = myCommand
		  end if
		  
		  myCommand = "CONNECT"
		  me.Connect
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub LogOut()
		  dim s As string
		  
		  
		  me.myCommand = "LOGOUT"
		  s = ". LOGOUT"
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub processMailboxLIST(List As String)
		  dim i, x As integer
		  dim rg As new RegEx
		  dim myMatch As RegExMatch
		  dim s, strMatch, data(), mailboxName, mailboxParent As string
		  dim mb As IMAPMailbox
		  dim mbRank() As string
		  
		   data = Split(List, EndOfLine.Windows)
		  
		  //CLEANUP THE OK RESPONSE
		  for i = UBound(data) DownTo 0
		    if left(data(i), 4) = ". OK" then
		      data.Remove i
		    end if
		  next
		  
		  rg.Options.Greedy = false
		  rg.Options.CaseSensitive = false
		  rg.Options.DotMatchAll = true
		  
		  for i = 0 to UBound(data)
		    mb = new IMAPMailbox
		    
		    s = data(i)
		    
		    //DETERMINE IF THIS MAILBOX HAS CHILDREN OR NOT
		    if instr(s, "haschildren") > 0 then
		      mb.hasChildren = true
		    else
		      mb.hasChildren = false
		    end if
		    
		    //DETERMINE IF THIS IS A FOLDER OF MAILBOX
		    rg.SearchPattern = "\(.*\)."
		    myMatch = rg.Search(s)
		    if myMatch <> nil then
		      strMatch = myMatch.SubExpressionString(0)
		      if instr(strMatch, "noselect") > 0 then
		        mb.isMailbox =false
		      else
		        mb.isMailbox = true
		      end if
		      
		      s = Replace(s, strMatch, "")
		    end if
		    
		    //FIND THE DELIMITER
		    rg.SearchPattern = """."""
		    myMatch = rg.Search(s)
		    if myMatch <> nil then
		      strMatch = myMatch.SubExpressionString(0)
		      's = Replace(s, strMatch, "")
		      
		      strMatch = ReplaceAll(strMatch, """", "")
		      myDelimiter = strMatch
		      mb.myDelimiter = myDelimiter
		    end if
		    
		    //STRIP EVERYTHING BUT WHAT SHOULD BE THE NAME
		    x = instr(s, myDelimiter+"""") + 2
		    s = mid(s, x)
		    s = trim(s)
		    
		    strMatch = s
		    
		    //FIND THE MAILBOX NAME
		    'rg.SearchPattern = """.*"""
		    'myMatch = rg.Search(s)
		    'if myMatch <> nil then
		    'strMatch = myMatch.SubExpressionString(0)
		    
		    x = StringUtils.InStrReverse(-1, strMatch, myDelimiter) + 1
		    mailboxName = mid(strMatch, x)
		    mailboxParent = strMatch
		    
		    //REMOVE QUOTATION MARKS FROM MAILBOXNAME
		    if Left(mailboxName, 1) = """" then
		      mailboxName = mid(mailboxName, 2)
		    end if
		    if Right(mailboxName, 1) = """" then
		      mailboxName = left(mailboxName, mailboxName.len-1)
		    end if
		    
		    //REMOVE QUOTATION MARKS FROM MAILBOXPARENT
		    if Left(mailboxParent, 1) = """" then
		      mailboxParent = mid(mailboxParent, 2)
		    end if
		    if Right(mailboxParent, 1) = """" then
		      mailboxParent = left(mailboxParent, mailboxParent.len-1)
		    end if
		    
		    mb.myName = mailboxName
		    mb.myParent = mailboxParent
		    mb.myHierarchy = split(mb.myParent, mb.myDelimiter)
		    'end if
		    
		    if mb.myName <> "" then
		      myMailboxes.Append mb
		      mbRank.Append mb.myParent
		    end if
		    
		  next
		  
		  mbRank.SortWith(myMailboxes)
		  
		  if UBound(myMailboxes) > -1 then
		    MailboxListReceived(myMailboxes)
		  end if
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub processUIDList(UIDList As string)
		  dim temp(0), strMatch As string
		  dim rg As new RegEx
		  dim myMatch As RegExMatch
		  dim i, x, uid(0) As integer
		  
		  rg.Options.Greedy = false
		  rg.Options.CaseSensitive = false
		  rg.Options.DotMatchAll = true
		  
		  temp = Split(UIDList, EndOfLine.Windows)
		  x = UBound(temp)
		  if instr(temp(x), ". OK") > 0 then
		    temp.Remove x
		  end if
		  
		  
		  for i = 1 to UBound(temp)
		    
		    rg.SearchPattern = "UID (.*) "
		    myMatch = rg.Search( temp(i) )
		    
		    if myMatch <> nil then
		      strMatch = myMatch.SubExpressionString(0)
		      
		      strMatch = Replace(strMatch, "UID ", "")
		      strMatch = trim(strMatch)
		      
		      uid.Append val(strMatch)
		      
		    end if
		    
		  next
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SelectMailbox(Mailbox As string)
		  dim s As string
		  
		  
		  me.myCommand = "SELECT"
		  s = ". SELECT """+Mailbox+""""
		  currentMailbox = Mailbox
		  
		  me.SendCommand s
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub SendCommand(myCommand As String)
		  
		  'if myCommand = "LOGIN" then
		  'LoggedCommunication "This is where the socket sends the password."
		  'else
		  'LoggedCommunication "Sending Command: "+myCommand
		  'end if
		  
		  Write myCommand + myCRLF
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Function TrimResponse(Response As string) As string
		  Response = trim(Response)
		  
		  return Response
		End Function
	#tag EndMethod


	#tag Hook, Flags = &h0
		Event BodyReceived(Index As Integer, EmailSource As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event FlagChanged(Index As Integer, Response As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event HeadersReceived(Index As Integer, EmailSource As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event LoggedCommunication(Response As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MailboxListReceived(Mailboxes() As IMAPMailbox)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MailboxSelected(Flags As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MessageAppended(Index As Integer, Response As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MessageCopied(Index As Integer, Response As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MessageDeleted(Index As Integer, Response As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MessageReceived(Index As Integer, EmailSource As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event MessagesExpunged(Response As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProgressChange(Status As string)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event ProtocolError(Error As String)
	#tag EndHook

	#tag Hook, Flags = &h0
		Event UIDSReceived(UIDList As String)
	#tag EndHook


	#tag Note, Name = RegEx
		rg.SearchPattern = "name=(.*)(:|$)"
		
	#tag EndNote


	#tag Property, Flags = &h0
		bufferData() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		bufferString As String
	#tag EndProperty

	#tag Property, Flags = &h0
		currentEmail As Integer
	#tag EndProperty

	#tag Property, Flags = &h0
		currentMailbox As String
	#tag EndProperty

	#tag Property, Flags = &h0
		firstConnect As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		imapPrefix As String
	#tag EndProperty

	#tag Property, Flags = &h0
		initialMailbox As string
	#tag EndProperty

	#tag Property, Flags = &h0
		myCommand As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private myCRLF As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private myDelimiter As String
	#tag EndProperty

	#tag Property, Flags = &h0
		myMailboxes() As IMAPMailbox
	#tag EndProperty

	#tag Property, Flags = &h0
		mySecondary As String
	#tag EndProperty

	#tag Property, Flags = &h0
		myState As String
	#tag EndProperty

	#tag Property, Flags = &h0
		myTag As string
	#tag EndProperty

	#tag Property, Flags = &h0
		Password As string
	#tag EndProperty

	#tag Property, Flags = &h0
		UserName As string
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="bufferString"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CertificateFile"
			Visible=true
			Group="Behavior"
			Type="FolderItem"
			EditorType="File"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CertificatePassword"
			Visible=true
			Group="Behavior"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="CertificateRejectionFile"
			Visible=true
			Group="Behavior"
			Type="FolderItem"
			EditorType="File"
		#tag EndViewProperty
		#tag ViewProperty
			Name="ConnectionType"
			Visible=true
			Group="Behavior"
			InitialValue="2"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="currentEmail"
			Group="Behavior"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="currentMailbox"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="firstConnect"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="imapPrefix"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			Type="Integer"
			EditorType="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="initialMailbox"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="myCommand"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="mySecondary"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="myState"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="myTag"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Password"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Secure"
			Visible=true
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
			EditorType="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="UserName"
			Group="Behavior"
			Type="string"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
