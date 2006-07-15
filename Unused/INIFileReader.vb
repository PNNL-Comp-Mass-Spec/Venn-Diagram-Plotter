Option Strict Off

Imports System
Imports System.Collections
Imports System.Collections.Specialized
Imports System.IO
Imports System.Windows.Forms
Imports System.Xml
Imports System.Xml.Xsl
Imports System.Xml.XPath
Imports System.Text

Enum IniItemTypeEnum
	GetKeys = 0
	GetValues = 1
	GetKeysAndValues = 2
End Enum

Public Class IniFileReaderNotInitializedException
	Inherits System.ApplicationException
	Public Overrides ReadOnly Property Message() As String
		Get
			Return "The IniFileReader instance has not been properly initialized."

		End Get
	End Property
End Class

Public Class IniFileReader
	Private m_IniFilename As String
	Private m_XmlDoc As XmlDocument
	Private unattachedComments As ArrayList = New ArrayList()
	Private sections As StringCollection = New StringCollection()
	Private m_CaseSensitive As Boolean = False
	Private m_SaveFilename As String
	Private m_initialized As Boolean = False

	Public Sub New(ByVal IniFilename As String)
		InitIniFileReader(IniFilename, False)
	End Sub
	Public Sub New(ByVal IniFilename As String, ByVal IsCaseSensitive As Boolean)
		InitIniFileReader(IniFilename, IsCaseSensitive)
	End Sub

	Private Sub InitIniFileReader(ByVal IniFilename As String, ByVal IsCaseSensitive As Boolean)
		Dim fi As FileInfo
		Dim s As String
		Dim tr As TextReader
		m_CaseSensitive = IsCaseSensitive
		m_XmlDoc = New XmlDocument()

		If ((IniFilename Is Nothing) OrElse (IniFilename.Trim() = "")) Then
			Return
		End If
		' try to load the file as an XML file
		Try
			m_XmlDoc.Load(IniFilename)
			UpdateSections()
			m_IniFilename = IniFilename
			m_initialized = True

		Catch
			' load the default XML
			m_XmlDoc.LoadXml("<?xml version=""1.0"" encoding=""UTF-8""?><sections></sections>")
			Try
                fi = New FileInfo(IniFilename)
				If (fi.Exists) Then
					tr = fi.OpenText
                    s = tr.ReadLine()
					Do While Not s Is Nothing
                        ParseLineXml(s, m_XmlDoc)
                        s = tr.ReadLine()
                    Loop
                    m_IniFilename = IniFilename
                    m_initialized = True
                Else
					m_XmlDoc.Save(IniFilename)
					m_IniFilename = IniFilename
					m_initialized = True
				End If
			Catch e As Exception
				MessageBox.Show(e.Message)
			Finally
				If (Not tr Is Nothing) Then
					tr.Close()
				End If
			End Try
		End Try
	End Sub

	Public ReadOnly Property IniFilename() As String
		Get
			If Not Initialized Then Throw New IniFileReaderNotInitializedException()
			Return (m_IniFilename)
		End Get
	End Property

	Public ReadOnly Property Initialized() As Boolean
		Get
			Return m_initialized
		End Get
	End Property

	Public ReadOnly Property CaseSensitive() As Boolean
		Get
			Return m_CaseSensitive
		End Get
	End Property

	Private Function SetNameCase(ByVal aName As String) As String
		If (CaseSensitive) Then
			Return aName
		Else
			Return aName.ToLower()
		End If
	End Function

	Private Function GetRoot() As XmlElement
		Return m_XmlDoc.DocumentElement
	End Function

	Private Function GetLastSection() As XmlElement
		If sections.Count = 0 Then
			Return GetRoot()
		Else
			Return GetSection(sections(sections.Count - 1))
		End If
	End Function

	Private Function GetSection(ByVal sectionName As String) As XmlElement
		If (Not (sectionName = Nothing)) AndAlso (sectionName <> "") Then
			sectionName = SetNameCase(sectionName)
			Return CType(m_XmlDoc.SelectSingleNode("//section[@name='" & sectionName & "']"), XmlElement)
		End If
		Return Nothing
	End Function

	Private Function GetItem(ByVal sectionName As String, ByVal keyName As String) As XmlElement
		Dim section As XmlElement
		If (Not keyName Is Nothing) AndAlso (keyName <> "") Then
			keyName = SetNameCase(keyName)
			section = GetSection(sectionName)
			If (Not section Is Nothing) Then
				Return CType(section.SelectSingleNode("item[@key='" + keyName + "']"), XmlElement)
			End If
		End If
		Return Nothing
	End Function

	Public Function SetIniSection(ByVal oldSection As String, ByVal newSection As String) As Boolean
		Dim section As XmlElement
		If Not Initialized Then
			Throw New IniFileReaderNotInitializedException()
		End If
		If (Not newSection Is Nothing) AndAlso (newSection <> "") Then
			section = GetSection(oldSection)
			If (Not (section Is Nothing)) Then
				section.SetAttribute("name", SetNameCase(newSection))
				UpdateSections()
				Return True
			End If
		End If
		Return False
	End Function

	Public Function SetIniValue(ByVal sectionName As String, ByVal keyName As String, ByVal newValue As String) As Boolean
		Dim item As XmlElement
		Dim section As XmlElement
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		section = GetSection(sectionName)
		If section Is Nothing Then
			If CreateSection(sectionName) Then
				section = GetSection(sectionName)
				' exit if keyName is Nothing or blank
				If (keyName Is Nothing) OrElse (keyName = "") Then
					Return True
				End If
			Else
				' can't create section
				Return False
			End If
		End If
		If keyName Is Nothing Then
			' delete the section
			Return DeleteSection(sectionName)
		End If

		item = GetItem(sectionName, keyName)
		If Not item Is Nothing Then
			If newValue Is Nothing Then
				' delete this item
				Return DeleteItem(sectionName, keyName)
			Else
				' add or update the value attribute
				item.SetAttribute("value", newValue)
				Return True
			End If
		Else
			' try to create the item
			If (keyName <> "") AndAlso (Not newValue Is Nothing) Then
				' construct a new item (blank values are OK)
				item = m_XmlDoc.CreateElement("item")
				item.SetAttribute("key", SetNameCase(keyName))
				item.SetAttribute("value", newValue)
				section.AppendChild(item)
				Return True
			End If
		End If
		Return False
	End Function

	Private Function DeleteSection(ByVal sectionName As String) As Boolean
		Dim section As XmlElement = GetSection(sectionName)
		If Not section Is Nothing Then
			section.ParentNode.RemoveChild(section)
			UpdateSections()
			Return True
		End If
		Return False
	End Function

	Private Function DeleteItem(ByVal sectionName As String, ByVal keyName As String) As Boolean
		Dim item As XmlElement = GetItem(sectionName, keyName)
		If Not item Is Nothing Then
			item.ParentNode.RemoveChild(item)
			Return True
		End If
		Return False
	End Function

	Public Function SetIniKey(ByVal sectionName As String, ByVal keyName As String, ByVal newValue As String) As Boolean
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Dim item As XmlElement = GetItem(sectionName, keyName)
		If Not item Is Nothing Then
			item.SetAttribute("key", SetNameCase(newValue))
			Return True
		End If
		Return False
	End Function

	Public Function GetIniValue(ByVal sectionName As String, ByVal keyName As String) As String
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Dim N As XmlNode = GetItem(sectionName, keyName)
		If Not N Is Nothing Then
			Return (N.Attributes.GetNamedItem("value").Value)
		End If
		Return Nothing
	End Function

	Public Function GetIniComments(ByVal sectionName As String) As StringCollection
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Dim sc As StringCollection = New StringCollection()
		Dim target As XmlNode
		Dim nodes As XmlNodeList
		Dim N As XmlNode
		If sectionName Is Nothing Then
			target = m_XmlDoc.DocumentElement
		Else
			target = GetSection(sectionName)
		End If
		If Not target Is Nothing Then
			nodes = target.SelectNodes("comment")
			If nodes.Count > 0 Then
				For Each N In nodes
					sc.Add(N.InnerText)
				Next
			End If
		End If
		Return sc
	End Function

	Public Function SetIniComments(ByVal sectionName As String, ByVal comments As StringCollection) As Boolean
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Dim target As XmlNode
		Dim nodes As XmlNodeList
		Dim N As XmlNode
		Dim s As String
		Dim NLastComment As XmlElement
		If sectionName Is Nothing Then
			target = m_XmlDoc.DocumentElement
		Else
			target = GetSection(sectionName)
		End If
		If Not target Is Nothing Then
			nodes = target.SelectNodes("comment")
			For Each N In nodes
				target.RemoveChild(N)
			Next
			For Each s In comments
				N = m_XmlDoc.CreateElement("comment")
				N.InnerText = s
				NLastComment = CType(target.SelectSingleNode("comment[last()]"), XmlElement)
				If NLastComment Is Nothing Then
					target.PrependChild(N)
				Else
					target.InsertAfter(N, NLastComment)
				End If
			Next
			Return True
		End If
		Return False
	End Function

	Private Sub UpdateSections()
		sections = New StringCollection()
		Dim N As XmlElement
		For Each N In m_XmlDoc.SelectNodes("sections/section")
			sections.Add(N.GetAttribute("name"))
		Next
	End Sub

	Public ReadOnly Property AllSections() As StringCollection
		Get
			If Not Initialized Then
				Throw New IniFileReaderNotInitializedException()
			End If
			Return sections
		End Get
	End Property

	Private Function GetItemsInSection(ByVal sectionName As String, ByVal itemType As IniItemTypeEnum) As StringCollection
		Dim nodes As XmlNodeList
		Dim items As StringCollection = New StringCollection()
		Dim section As XmlNode = GetSection(sectionName)
		Dim N As XmlNode
		If section Is Nothing Then
			Return Nothing
		Else
			nodes = section.SelectNodes("item")
			If nodes.Count > 0 Then
				For Each N In nodes
					Select Case itemType
						Case IniItemTypeEnum.GetKeys
							items.Add(N.Attributes.GetNamedItem("key").Value)
						Case IniItemTypeEnum.GetValues
							items.Add(N.Attributes.GetNamedItem("value").Value)
						Case IniItemTypeEnum.GetKeysAndValues
							items.Add(N.Attributes.GetNamedItem("key").Value & "=" & _
							N.Attributes.GetNamedItem("value").Value)
					End Select
				Next
			End If
			Return items
		End If
	End Function

	Public Function AllKeysInSection(ByVal sectionName As String) As StringCollection
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Return GetItemsInSection(sectionName, IniItemTypeEnum.GetKeys)
	End Function

	Public Function AllValuesInSection(ByVal sectionName As String) As StringCollection
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Return GetItemsInSection(sectionName, IniItemTypeEnum.GetValues)
	End Function

	Public Function AllItemsInSection(ByVal sectionName As String) As StringCollection
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Return (GetItemsInSection(sectionName, IniItemTypeEnum.GetKeysAndValues))
	End Function

	Public Function GetCustomIniAttribute(ByVal sectionName As String, ByVal keyName As String, ByVal attributeName As String) As String
		Dim N As XmlElement
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		If (Not attributeName Is Nothing) AndAlso (attributeName <> "") Then
			N = GetItem(sectionName, keyName)
			If Not N Is Nothing Then
				attributeName = SetNameCase(attributeName)
				Return N.GetAttribute(attributeName)
			End If
		End If
		Return Nothing
	End Function

	Public Function SetCustomIniAttribute(ByVal sectionName As String, ByVal keyName As String, ByVal attributeName As String, ByVal attributeValue As String) As Boolean
		Dim N As XmlElement
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		If attributeName <> "" Then
			N = GetItem(sectionName, keyName)
			If Not N Is Nothing Then
				Try
					If attributeValue Is Nothing Then
						' delete the attribute
						N.RemoveAttribute(attributeName)
						Return True
					Else
						attributeName = SetNameCase(attributeName)
						N.SetAttribute(attributeName, attributeValue)
						Return True
					End If

				Catch e As Exception
					MessageBox.Show(e.Message)
				End Try
			End If
			Return False
		End If
	End Function

	Private Function CreateSection(ByVal sectionName As String) As Boolean
		Dim N As XmlElement
		Dim Natt As XmlAttribute
		If (Not sectionName Is Nothing) AndAlso (sectionName <> "") Then
			sectionName = SetNameCase(sectionName)
			Try
				N = m_XmlDoc.CreateElement("section")
				Natt = m_XmlDoc.CreateAttribute("name")
				Natt.Value = SetNameCase(sectionName)
				N.Attributes.SetNamedItem(Natt)
				m_XmlDoc.DocumentElement.AppendChild(N)
				sections.Add(Natt.Value)
				Return True
			Catch e As Exception
				MessageBox.Show(e.Message)
				Return False
			End Try
		End If
		Return False
	End Function

	Private Function CreateItem(ByVal sectionName As String, ByVal keyName As String, ByVal newValue As String) As Boolean
		Dim item As XmlElement
		Dim section As XmlElement
		Try
			section = GetSection(sectionName)
			If Not section Is Nothing Then
				item = m_XmlDoc.CreateElement("item")
				item.SetAttribute("key", keyName)
				item.SetAttribute("newValue", newValue)
				section.AppendChild(item)
				Return True
			End If
			Return False
		Catch e As Exception
			MessageBox.Show(e.Message)
			Return False
		End Try
	End Function

	Private Sub ParseLineXml(ByVal s As String, ByVal doc As XmlDocument)
		Dim key As String
		Dim value As String
		Dim N As XmlElement
		Dim Natt As XmlAttribute
		Dim parts() As String
		s.TrimStart()

		If s.Length = 0 Then
			Return
		End If
		Select Case (s.Substring(0, 1))
			Case "["
				' this is a section
				' trim the first and last characters
				s = s.TrimStart("[")
				s = s.TrimEnd("]")
				' create a new section element
				CreateSection(s)
			Case ";"
				' new comment
				N = doc.CreateElement("comment")
				N.InnerText = s.Substring(1)
				GetLastSection().AppendChild(N)
			Case Else
				' split the string on the "=" sign, if present
				If (s.IndexOf("=") > 0) Then
					parts = s.Split("=")
					key = parts(0).Trim()
					value = parts(1).Trim()
				Else
					key = s
					value = ""
				End If
				N = doc.CreateElement("item")
				Natt = doc.CreateAttribute("key")
				Natt.Value = SetNameCase(key)
				N.Attributes.SetNamedItem(Natt)
				Natt = doc.CreateAttribute("value")
				Natt.Value = value
				N.Attributes.SetNamedItem(Natt)
				GetLastSection().AppendChild(N)
		End Select

	End Sub

	Public Property OutputFilename() As String
		Get
			If Not Initialized Then Throw New IniFileReaderNotInitializedException()
			Return m_SaveFilename
		End Get
		Set(ByVal Value As String)
			Dim fi As FileInfo
			If Not Initialized Then Throw New IniFileReaderNotInitializedException()
			fi = New FileInfo(Value)
			If Not fi.Directory.Exists Then
				MessageBox.Show("Invalid path.")
			Else
				m_SaveFilename = Value
			End If
		End Set
	End Property

	Public Sub Save()
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		If Not OutputFilename Is Nothing AndAlso Not m_XmlDoc Is Nothing Then
			Dim fi As FileInfo = New FileInfo(OutputFilename)
			If Not fi.Directory.Exists Then
				MessageBox.Show("Invalid path.")
				Return
			End If
			If fi.Exists Then
				fi.Delete()
				m_XmlDoc.Save(OutputFilename)
			Else
				m_XmlDoc.Save(OutputFilename)
			End If
		End If
	End Sub

	Public Function AsIniFile() As String
		If Not Initialized Then Throw New IniFileReaderNotInitializedException()
		Try
			Dim xsl As XslTransform = New XslTransform()
			xsl.Load("c:\\XMLToIni.xslt")
			Dim sb As StringBuilder = New StringBuilder()
			Dim sw As StringWriter = New StringWriter(sb)
			xsl.Transform(m_XmlDoc, Nothing, sw, Nothing)
			sw.Close()
			Return sb.ToString
		Catch e As Exception
			MessageBox.Show(e.Message)
			Return Nothing
		End Try
	End Function

	Public ReadOnly Property XmlDoc() As XmlDocument
		Get
			If Not Initialized Then Throw New IniFileReaderNotInitializedException()
			Return m_XmlDoc
		End Get
	End Property

	Public ReadOnly Property XML() As String
		Get
			If Not Initialized Then Throw New IniFileReaderNotInitializedException()
			Dim sb As StringBuilder = New StringBuilder()
			Dim sw As StringWriter = New StringWriter(sb)
			Dim xw As XmlTextWriter = New XmlTextWriter(sw)
			xw.Indentation = 3
			xw.Formatting = Formatting.Indented
			m_XmlDoc.WriteContentTo(xw)
			xw.Close()
			sw.Close()
			Return sb.ToString()
		End Get
	End Property
End Class