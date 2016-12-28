#tag Class
Protected Class IMAPMailbox
	#tag Property, Flags = &h0
		hasChildren As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		isMailbox As Boolean
	#tag EndProperty

	#tag Property, Flags = &h0
		myDelimiter As String
	#tag EndProperty

	#tag Property, Flags = &h0
		myHierarchy() As String
	#tag EndProperty

	#tag Property, Flags = &h0
		myName As String
	#tag EndProperty

	#tag Property, Flags = &h0
		myParent As String
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="hasChildren"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="isMailbox"
			Group="Behavior"
			Type="Boolean"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="myDelimiter"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="myName"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
		#tag EndViewProperty
		#tag ViewProperty
			Name="myParent"
			Group="Behavior"
			Type="String"
			EditorType="MultiLineEditor"
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
End Class
#tag EndClass
