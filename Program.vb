Imports System.Net.NetworkInformation
Imports System.Management
Imports System.Xml
Imports RestSharp
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json
Public Class Var
	'Allgemeine Infos
	Public Shared logpath As String
	Public Shared curbenutzer As String = Environment.UserDomainName & Chr(92) & Environment.UserName
	Public Shared time As String = Date.Now.Hour.ToString & "." & Date.Now.Minute.ToString
	Public Shared computerid As Integer
	Public Shared computername As String = Environment.MachineName
	Public Shared noLogs As Boolean = False
	Public Shared noXML As Boolean = True
	Public Shared pxe As Boolean = False
	Public Shared streamdirectory As String
	Public Shared macadressen(,) As String
	''API Inos
	Public Shared noAPI As Boolean = False
	Public Shared APIPath As String
	Public Shared APIKey As String
	Public Shared APIcomponents As JArray
	Public Shared APICheckedcomponents(0) As Integer
	Public Shared APIlocation_id As Integer
	Public Shared APIcompany_id As Integer
	''API categories ids
	Public Shared APIcategorieid(5) As Integer ''CPU,GPU,HDD/SSD,RAM,Mainboard,Computer/Notebook
	Public Shared APIfieldsetid As Integer
	'OS Informationen
	Public Shared osinstalldatum As String
	Public Shared osversion As String = Environment.OSVersion.ToString
	'Bios Informationen
	Public Shared biosManufacturer(0) As String
	Public Shared biosName(0) As String
	Public Shared biosReleaseDate(0) As String
	Public Shared biosSerialNumber(0) As String
	'Mainboard Inforamtionen
	Public Shared mainboardManufacturer(0) As String
	Public Shared mainboardProduct(0) As String
	Public Shared mainboardVersion(0) As String
	'Festplatten Inforamtionen
	Public Shared festplatteFirmwareRevision() As String
	Public Shared festplatteModel() As String
	Public Shared festplatteSerialNumber() As String
	Public Shared festplatteSize() As String
	Public Shared festplatteanzahl() As Integer
	'RAM Inforamtionen
	Public Shared ramBankLabel() As String
	Public Shared ramCapacity() As String
	Public Shared ramManufacturer() As String
	Public Shared ramPartNumber() As String
	Public Shared ramseriennummer() As String
	Public Shared ramanzahl() As Integer
	'GPU Inforamtionen
	Public Shared gpuName() As String
	Public Shared gpuVideoProcessor() As String
	'CPU Inforamtionen
	Public Shared cpuManufacturer() As String
	Public Shared cpuName() As String
	Public Shared cpuDeviceID() As String
	'Allgemeine Computer Inforamtionen
	Public Shared computerManufacturer(0) As String
	Public Shared computerModel(0) As String
	Public Shared computerSystemFamily(0) As String
	Public Shared computerSystemSKUNumber(0) As String
	'Lizenz Inforamtionen
	Public Shared softwareliceOA3xOriginalProductKey(0) As String
	Public Shared softwareliceOA3xOriginalProductKeyDescription(0) As String
	'Software
	Public Shared antivirus() As String
	''Software Regedit
	Public Shared SoftwareRegEdit(5, 1) As String
End Class
Module Program
	Sub Main()
		Logschreiber(Log:="Startzeit -> " & DateTime.Now.ToString, art:="log", section:="Main")
		Console.WriteLine("ReadConf")
		Readconf()
		Console.WriteLine("Argumentcheck")
		Argumentscheck()
		Console.WriteLine("Logschreiber")
		Logschreiber(Log:="Startzeit " & DateTime.Now.ToString, art:="log", section:="Main")
		Console.WriteLine("WMIAbfragenAllgemein")
		WMIAbfragenAllgemein()
		Console.WriteLine("WMIMainboard")
		WMIMainboard()
		Console.WriteLine("WMIBIOS")
		WMIBIOS()
		Console.WriteLine("WMIFestplatte")
		WMIFestplatte()
		Console.WriteLine("WMIRAM")
		WMIRAM()
		Console.WriteLine("WMIGPU")
		WMIGPU()
		Console.WriteLine("WMICPU")
		WMICPU()
		Console.WriteLine("macadresse")
		Macadresse()
		If Var.noXML = False Then
			Xmlauswertung()
		End If
		If Var.noAPI = False Then
			Console.WriteLine("APITest")
			If APITest() = Net.HttpStatusCode.OK Then
				Console.WriteLine("computerfinden")
				APIGETAssetID() 'Sucht die Computer ID aus SnipIT
				If Var.computerid <> 0 Then
					Console.WriteLine("APICPU")
					APICPU()
					Console.WriteLine("APIGPU")
					APIGPU()
					Console.WriteLine("APIRAM")
					APIRAM()
					Console.WriteLine("APIHDD_SSD")
					APIHDD_SSD()
					Console.WriteLine("ReversCheck")
					ReversCheck()
				Else
					Logschreiber(Log:="Fehler keine Computer ID", art:="fehler", section:="Main")
				End If
			Else
				Logschreiber(Log:="Fehler API", art:="fehler", section:="APITest")
			End If
		Else
			Logschreiber(Log:="Keine API Eintragungen", art:="log", section:="Main")
		End If
		Logschreiber(Log:="Endzeit " & DateTime.Now.ToString, art:="log", section:="Main")
	End Sub
	'''Function
	Function Logschreiber(Log As String, art As String, section As String)
		Console.WriteLine(Log & "    Sektion -> " & section)
		If Var.noLogs = False Then
			Try
				System.IO.Directory.CreateDirectory(Var.logpath & Environment.MachineName & "\" & DateAndTime.Now.Day.ToString & "_" & DateAndTime.Now.Month.ToString & "_" & DateAndTime.Now.Year.ToString)
				Var.streamdirectory = Var.logpath & Environment.MachineName & "\" & DateAndTime.Now.Day.ToString & "_" & DateAndTime.Now.Month.ToString & "_" & DateAndTime.Now.Year.ToString
			Catch ex As Exception
				Try
					System.IO.Directory.CreateDirectory(".\" & Environment.MachineName)
					Var.streamdirectory = ".\" & Environment.MachineName
				Catch ex1 As Exception
					Var.noLogs = True
				End Try
			End Try
			If Var.noLogs = False Then
				Try
					Dim LogStreamWriter As New IO.StreamWriter(Var.streamdirectory & "\" & art & Var.time & "Uhr.txt", True)
					LogStreamWriter.WriteLine(section & " = " & Log)
					LogStreamWriter.Close()
				Catch ex As System.IO.IOException
				End Try
				Return 0
			Else
			End If
			Return 0
		End If
	End Function
	<CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Plattformkompatibilität überprüfen", Justification:="<Ausstehend>")>
	Function WMIAbfrage(WMISlect As String, WMIFrom As String, Optional WMIScope As String = "root\CIMV2", Optional nolog As Boolean = False)
		Dim wert(0) As String
		Dim i As Integer
		Try
			Dim ms As New ManagementScope(WMIScope)
			Dim mq As New SelectQuery("Select " & WMISlect & " FROM " & WMIFrom)
			Dim mos As New ManagementObjectSearcher(ms, mq)
			Dim moc As ManagementObjectCollection = mos.Get()
			If moc.Count - 1 <> 0 Then
				ReDim wert(moc.Count - 1)
			End If
			Dim mo As ManagementObject
			For Each mo In moc
				Dim pd As PropertyData
				For Each pd In mo.Properties
					If IsNothing(pd.Value) Then
						wert(i) = "Keine Daten"
						Console.WriteLine("Value " & wert(i))
						If nolog = False Then
							Logschreiber(Log:=WMISlect & " = " & wert(i), art:="log", section:="WMIAbfrage")
						End If
					Else
						If pd.Value.GetType.Name.ToString = "UInt16[]" Then
							Dim UIntWert() As UInt16 = pd.Value
							For j As Integer = 0 To UIntWert.Length - 1 Step 1
								If UIntWert(j) <> 0 Then
									wert(i) = wert(i) & Chr(UIntWert(j))
									Console.WriteLine("Value " & wert(i))
								End If
							Next
						Else
							wert(i) = Trim(pd.Value.ToString())
							Console.WriteLine("Value " & wert(i))
						End If
						Console.WriteLine("VALUE = " & wert(i))
						If nolog = False Then
							Logschreiber(Log:=WMISlect & " = " & wert(i), art:="log", section:="WMIAbfrage")
						End If
					End If
					i += 1
				Next
			Next
			Return wert
		Catch ex As Management.ManagementException
			Logschreiber(Log:=ex.ToString & vbCrLf & "WMI Befehl -> Select " & WMISlect & " FROM " & WMIFrom, art:="fehler", section:="WMIAbfrage")
			wert(0) = "Keine Daten"
			Return wert
		End Try
	End Function
	''API
	Function APITest() As Net.HttpStatusCode
		Try
			Dim RestSharpclient As New RestClient(Var.APIPath & "components")
			Dim RestSharpRequest As New RestRequest
			RestSharpRequest.AddHeader("Authorization", "Bearer " & Var.APIKey)
			RestSharpRequest.AddHeader("Accept", "application/json")
			Dim RestSharpResponse As Task(Of RestResponse)
			RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequest)
			Return RestSharpResponse.Result.StatusCode
		Catch ex As Exception
			Logschreiber(Log:="Fehler API" & vbCrLf & ex.ToString, art:="fehler", section:="APITest")
			Return Net.HttpStatusCode.InternalServerError
		End Try
	End Function
	Function APIGetModel(model As String, hersteller As String, fieldset_id As Integer, category_id As Integer, Optional computermodell As Boolean = False, Optional systemfamily As String = "") As Integer
		Try
			Dim RestSharpclient As New RestClient(Var.APIPath)
			Dim RestSharpRequest As New RestRequest("models?search=" & model)
			RestSharpRequest.AddHeader("Authorization", "Bearer " & Var.APIKey)
			RestSharpRequest.AddHeader("Accept", "application/json")
			Dim RestSharpResponse As Task(Of RestResponse)
			RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequest)
			Dim Data As JObject = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
			If Convert.ToInt32(Data("total")) = 0 Then
				''Get HerstellerID
				Dim herstellerid As Integer = 0
				Dim RestSharpRequestGetManufacturers As New RestRequest("manufacturers?search=" & hersteller)
				RestSharpRequestGetManufacturers.AddHeader("Authorization", "Bearer " & Var.APIKey)
				RestSharpRequestGetManufacturers.AddHeader("Accept", "application/json")
				RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestGetManufacturers)
				Data = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
				If Convert.ToInt32(Data("total")) = 0 Then
					Dim RestSharpRequestAddManufacturers As New RestRequest("manufacturers", Method.Post)
					RestSharpRequestAddManufacturers.AddHeader("Authorization", "Bearer " & Var.APIKey)
					RestSharpRequestAddManufacturers.AddHeader("Accept", "application/json")
					RestSharpRequestAddManufacturers.AddJsonBody(New With {Key .name = hersteller})
					RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestAddManufacturers)
					Data = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
					If Data("status").ToString = "success" Then
						Dim Data1 As JObject = Data("payload")
						herstellerid = Data1("id")
					Else
						Logschreiber(Log:="Fehler APIGetModell" & vbCrLf & Data.ToString, art:="fehler", section:="APIGetModell_ADD_Manufactor")
						Environment.Exit(1)
					End If
				Else
					Dim Data1 As JArray = Data("rows")
					For Each JObject As JObject In Data1
						If hersteller = JObject.Item("name").ToString Then
							herstellerid = Convert.ToInt32(JObject.Item("id"))
						End If
					Next
				End If
				''End_Get HerstellerID
				''Add Model
				Dim RestSharpRequestAddModell As New RestRequest("models", Method.Post)
				RestSharpRequestAddModell.AddHeader("Authorization", "Bearer " & Var.APIKey)
				RestSharpRequestAddModell.AddHeader("Accept", "application/json")
				If computermodell Then
					RestSharpRequestAddModell.AddJsonBody(New With {Key .name = systemfamily, Key .category_id = category_id, Key .manufacturer_id = herstellerid, Key .fieldset_id = fieldset_id, Key .model_number = model})
				Else
					RestSharpRequestAddModell.AddJsonBody(New With {Key .name = model, Key .category_id = category_id, Key .manufacturer_id = herstellerid, Key .fieldset_id = fieldset_id})
				End If
				RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestAddModell)
				Data = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
				If Data("status").ToString = "success" Then
					Dim Data1 As JObject = Data("payload")
					Return Convert.ToInt32(Data1("id"))
				Else
					Logschreiber(Log:="Fehler APIGetModell" & vbCrLf & Data.ToString, art:="fehler", section:="APIGetModell_ADD_Model")
					Environment.Exit(1)
				End If
				''End_Add Model
			Else
				Dim ThisData1 As JArray = Data("rows")
				For Each JObject As JObject In ThisData1
					Return Convert.ToInt32(JObject.Item("id"))
				Next
			End If
		Catch ex As Exception
			Logschreiber(Log:="Fehler APIGetModell" & vbCrLf & ex.ToString, art:="fehler", section:="APIGetModell")
			Environment.Exit(1)
		End Try
	End Function
	Function APIcomponentsAdd(name As String, category_id As Integer, Optional serial As String = "") As Boolean
		Dim RestSharpclient As New RestClient(Var.APIPath)
		Dim RestSharpRequest As New RestRequest("components?search=" & name)
		RestSharpRequest.AddHeader("Authorization", "Bearer " & Var.APIKey)
		RestSharpRequest.AddHeader("Accept", "application/json")
		Dim RestSharpResponse As Task(Of RestResponse)
		RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequest)
		Dim Data As JObject = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
		Select Case Convert.ToInt32(Data.Item("total"))
			Case 0
				Dim RestSharpRequestAddcomponents As New RestRequest("components", Method.Post)
				RestSharpRequestAddcomponents.AddHeader("Authorization", "Bearer " & Var.APIKey)
				RestSharpRequestAddcomponents.AddHeader("Accept", "application/json")
				RestSharpRequestAddcomponents.AddJsonBody(New With {Key .name = name, Key .qty = 1, Key .category_id = category_id, Key .location_id = Var.APIlocation_id, Key .company_id = Var.APIcompany_id})
				RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestAddcomponents)
				Data = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
				If Data("status").ToString = "success" Then
					Dim Data2 As JObject = Data("payload")
					If APICheckout_in(Convert.ToInt32(Data2("id")), True) = False Then
						Return False
					Else
						Return True
					End If
				Else
					Logschreiber(Log:="Fehler components ADD" & vbCrLf & Data.ToString, art:="fehler", section:="APIcomponentsAdd")
					Return False
				End If
			Case 1
				Dim DataJarray As JArray = Data.Item("rows")
				For Each JObject As JObject In DataJarray
					If APICheckout_in(Convert.ToInt32(JObject.Item("id")), True) = False Then
						Return False
					End If
				Next
			Case > 1
				If serial <> "" Then
					Dim DataJarray As JArray = Data.Item("rows")
					For Each JObject As JObject In DataJarray
						If JObject.Item("serial").ToString = serial Then
							If APICheckout_in(Convert.ToInt32(DataJarray.Item("id")), True) = False Then
								Return False
							Else
								Return True
							End If
							Exit For
						End If
					Next
				Else
					Logschreiber(Log:="Fehler components" & vbCrLf & "Componente kann nicht eindeutig identifiziert werden, Serienummer fehlt", art:="fehler", section:="APIcomponentsAdd")
					Return False
				End If
		End Select
	End Function
	Function APICheckout_in(id As Integer, checkout As Boolean) As Boolean
		''Anzahl verfügbarer kompmenten
		Dim remainingcomponents, qty As Integer
		Dim RestSharpclient As New RestClient(Var.APIPath)
		Dim RestSharpRequestGetRemaining As New RestRequest("components/1")
		RestSharpRequestGetRemaining.AddHeader("Authorization", "Bearer " & Var.APIKey)
		RestSharpRequestGetRemaining.AddHeader("Accept", "application/json")
		Dim RestSharpResponseGetRemaining As Task(Of RestResponse)
		RestSharpResponseGetRemaining = RestSharpclient.ExecuteAsync(RestSharpRequestGetRemaining)
		Dim Data3 As JObject = JsonConvert.DeserializeObject(RestSharpResponseGetRemaining.Result.Content.ToString)
		remainingcomponents = Convert.ToInt32(Data3.Item("remaining"))
		qty = Convert.ToInt32(Data3.Item("qty"))
		Dim RestSharpRequestCheckout As New RestRequest("")
		If checkout = True Then ''vor checkout anzhal falls nötig erhöhen
			RestSharpRequestCheckout.Resource = "components/" & id.ToString & "/checkout"
			RestSharpRequestCheckout.AddJsonBody(New With {Key .assigned_to = Var.computerid, Key .assigned_qty = 1})
			If remainingcomponents = 0 Or remainingcomponents > 1 Then
				Dim RestSharpRequestUpdateQTY As New RestRequest("components/" & id.ToString, Method.Patch)
				Dim RestSharpResponseUpdateQTY As Task(Of RestResponse)
				RestSharpRequestUpdateQTY.AddHeader("Authorization", "Bearer " & Var.APIKey)
				RestSharpRequestUpdateQTY.AddHeader("Accept", "application/json")
				If remainingcomponents = 0 Then
					RestSharpRequestUpdateQTY.AddJsonBody(New With {Key .qty = qty + 1})
				Else
					RestSharpRequestUpdateQTY.AddJsonBody(New With {Key .qty = qty - (remainingcomponents - 1)})
				End If
				RestSharpResponseUpdateQTY = RestSharpclient.ExecuteAsync(RestSharpRequestUpdateQTY)
				Dim Data4 As JObject = JsonConvert.DeserializeObject(RestSharpResponseUpdateQTY.Result.Content.ToString)
				If Data4("status").ToString = "success" Then
				Else
					Logschreiber(Log:="Patch QTY components" & vbCrLf & Data4.ToString, art:="fehler", section:="APICheckout_in-Patch_QTY_components")
					Return False
				End If
			End If
		Else ''CheckIn
			Dim RestSharpRequestCheckInID As New RestRequest("components/" & id.ToString & "/assets")
			RestSharpRequestCheckInID.AddHeader("Authorization", "Bearer " & Var.APIKey)
			RestSharpRequestCheckInID.AddHeader("Accept", "application/json")
			Dim RestSharpResponseCheckInID As Task(Of RestResponse)
			RestSharpResponseCheckInID = RestSharpclient.ExecuteAsync(RestSharpRequestCheckInID)
			Dim Data1 As JObject = JsonConvert.DeserializeObject(RestSharpResponseCheckInID.Result.Content.ToString)
			Dim assigned_pivot_id As Integer
			If Convert.ToInt32(Data1("total")) > 0 Then
				Dim JArray As JArray = Data1("rows")
				For Each JObject As JObject In JArray
					If Convert.ToInt32(JObject.Item("id")) = Var.computerid Then
						assigned_pivot_id = Convert.ToInt32(JObject.Item("assigned_pivot_id"))
						Exit For
					End If
				Next
			Else
			End If
			RestSharpRequestCheckout.Resource = "components/" & assigned_pivot_id.ToString & "/checkin"
			RestSharpRequestCheckout.AddJsonBody(New With {Key .assigned_qty = 1})
		End If
		RestSharpRequestCheckout.Method = Method.Post
		RestSharpRequestCheckout.AddHeader("Authorization", "Bearer " & Var.APIKey)
		RestSharpRequestCheckout.AddHeader("Accept", "application/json")
		Dim RestSharpResponse As Task(Of RestResponse)
		RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestCheckout)
		Dim Data As JObject = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
		If Data("status").ToString = "success" Then
			If checkout = False Then ''Nach Checkin Anzahl veringern
				If remainingcomponents = 0 Or remainingcomponents > 1 Then
					Dim RestSharpRequestUpdateQTY As New RestRequest("components/" & id.ToString, Method.Patch)
					Dim RestSharpResponseUpdateQTY As Task(Of RestResponse)
					RestSharpRequestUpdateQTY.AddHeader("Authorization", "Bearer " & Var.APIKey)
					RestSharpRequestUpdateQTY.AddHeader("Accept", "application/json")
					If remainingcomponents = 0 Then
						RestSharpRequestUpdateQTY.AddJsonBody(New With {Key .qty = qty + 1})
					Else
						RestSharpRequestUpdateQTY.AddJsonBody(New With {Key .qty = qty - (remainingcomponents - 1)})
					End If
					RestSharpResponseUpdateQTY = RestSharpclient.ExecuteAsync(RestSharpRequestUpdateQTY)
					Dim Data4 As JObject = JsonConvert.DeserializeObject(RestSharpResponseUpdateQTY.Result.Content.ToString)
					If Data4("status").ToString = "success" Then
					Else
						Logschreiber(Log:="Patch QTY components" & vbCrLf & Data4.ToString, art:="fehler", section:="APICheckout_in-Patch_QTY_components")
						Return False
					End If
				End If
			End If
			Return True
		Else
			Logschreiber(Log:="Fehler APICheckout_in" & vbCrLf & Data.ToString, art:="fehler", section:="APIcomponentsAdd_Checkoutcomponent")
			Return False
		End If
	End Function
	'''Sub
	Sub Argumentscheck()
		Dim i As Integer = 0
		For Each s As String In Environment.GetCommandLineArgs
			If i > 0 Then
				Console.WriteLine("Parameter -> " & s)
				Select Case s
					Case "-noAPI"
						Console.WriteLine("Kein SQL")
						Var.noAPI = True
					Case "-XML"
						Console.WriteLine("XML Ausgabe")
						Var.noXML = False
					Case "-noLogs"
						Console.WriteLine("Keine Logs")
						Var.noLogs = True
					Case "-Help"
						Console.WriteLine("-noAPI -> Nur auslesen keine Anpaasung über die API")
						Console.WriteLine("-XML -> es erfolgt eine zusätzliche Ausgabe einer XML zusammenfassung der Computer Informationen ins Log Verzeichniss")
						Console.WriteLine("-noLogs -> Es werden keine Logs lokal, auf dem  SMB share oder in die DB geschrieben")
						Console.WriteLine("-Help -> zeigt das Menu an")
						End
					Case Else
						Logschreiber(Log:="Fehler Argument " & s & " konnte nicht erkannt werden.", art:="fehler", section:="Main")
				End Select
			End If
			i += 1
		Next
	End Sub
	Sub Readconf()
		If (IO.File.Exists(".\Config.xml")) Then
			Dim XMLReader As New XmlTextReader(".\Config.xml")
			Try
				While XMLReader.Read
					If XMLReader.Name = "logpath" Then
						Var.logpath = XMLReader.ReadElementString()
					ElseIf XMLReader.Name = "APIPath" Then
						Var.APIPath = XMLReader.ReadElementString()
					ElseIf XMLReader.Name = "APIKey" Then
						Var.APIKey = XMLReader.ReadElementString()
					ElseIf XMLReader.Name = "APIlocation_id" Then
						Var.APIlocation_id = XMLReader.ReadElementContentAsInt()
					ElseIf XMLReader.Name = "APIcompany_id" Then
						Var.APIcompany_id = XMLReader.ReadElementContentAsInt()
					ElseIf XMLReader.Name = "APIcategorieid" Then
						Dim s() As String = Split(XMLReader.ReadElementString(), ",")
						Dim i As Integer
						For Each str As String In s
							Var.APIcategorieid(i) = Convert.ToInt32(str)
							i += 1
						Next
					ElseIf XMLReader.Name = "APIfieldsetid" Then
						Var.APIfieldsetid = XMLReader.ReadElementContentAsInt()
					End If
				End While
				XMLReader.Close()
			Catch ex As Exception
				Console.WriteLine(ex.ToString)
				Environment.Exit(1)
			End Try
		Else
			Console.WriteLine(".\Config.xml wurde nicht gefunden!")
			Environment.Exit(1)
		End If
	End Sub
	Sub Macadresse()
		Dim j As Integer
		Try
			Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
			If NetworkInterface.GetAllNetworkInterfaces.Length = 0 Then
				If nics(0).GetPhysicalAddress.ToString() <> "" Then
					If nics(0).GetPhysicalAddress.ToString().Length = 12 Then
						ReDim Preserve Var.macadressen(1, 2)
						Var.macadressen(0, 0) &= nics(0).GetPhysicalAddress.ToString()
						Var.macadressen(0, 1) &= nics(0).Name.ToString
						Var.macadressen(0, 2) &= nics(0).Description.ToString
						Logschreiber(Log:="Macadresse = " & nics(0).GetPhysicalAddress.ToString() & " Name=  " & nics(0).Name.ToString & " Beschreibung " & nics(0).Description.ToString, art:="log", section:="macadresse")
					End If
				End If
			Else
				For k As Integer = 0 To NetworkInterface.GetAllNetworkInterfaces.Length - 1 Step 1
					If nics(k).GetPhysicalAddress.ToString() <> "" Then
						If nics(k).GetPhysicalAddress.ToString().Length = 12 Then
							j += 1
						End If
					End If
				Next
				ReDim Preserve Var.macadressen(j - 1, 2)
				For l As Integer = 0 To j - 1 Step 1
					If nics(l).GetPhysicalAddress.ToString() <> "" Then
						If nics(l).GetPhysicalAddress.ToString().Length = 12 Then
							Var.macadressen(l, 0) &= nics(l).GetPhysicalAddress.ToString()
							Var.macadressen(l, 1) &= nics(l).Name.ToString
							Var.macadressen(l, 2) &= nics(l).Description.ToString
							Logschreiber(Log:=nics(l).GetPhysicalAddress.ToString() & " Name =" & nics(l).Name.ToString & " Beschreibung =" & nics(l).Description.ToString, art:="log", section:="macadresse")
						End If
					End If
				Next
			End If
		Catch ex As Exception
			Logschreiber(Log:=ex.ToString, art:="fehler", section:="macadresse")
		End Try
	End Sub
	Sub Xmlauswertung()
		Dim XmlDoc As New XmlDocument
		Dim XmlDeclaration As XmlDeclaration = XmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
		'Create the root element
		Dim RootNode As XmlElement = XmlDoc.CreateElement("Computer_Infos")
		XmlDoc.InsertBefore(XmlDeclaration, XmlDoc.DocumentElement)
		XmlDoc.AppendChild(RootNode)
		Try
			''AllgemeineInfos
			Dim AllgemeineInfos As XmlElement = XmlDoc.CreateElement("AllgemeineInfos")
			XmlDoc.DocumentElement.PrependChild(AllgemeineInfos)
			'Computername
			Dim ComputernameElement As XmlElement = XmlDoc.CreateElement("Computername")
			Dim ComputernameText As XmlText = XmlDoc.CreateTextNode(Var.computername)
			AllgemeineInfos.AppendChild(ComputernameElement)
			ComputernameElement.AppendChild(ComputernameText)
			'ComputerHersteller
			Dim ComputerHerstellerElement As XmlElement = XmlDoc.CreateElement("ComputerHersteller")
			Dim ComputerHerstellerText As XmlText = XmlDoc.CreateTextNode(Var.computerManufacturer(0))
			AllgemeineInfos.AppendChild(ComputerHerstellerElement)
			ComputerHerstellerElement.AppendChild(ComputerHerstellerText)
			'ComputerModell
			Dim ComputerModellElement As XmlElement = XmlDoc.CreateElement("ComputerModell")
			Dim ComputerModellText As XmlText = XmlDoc.CreateTextNode(Var.computerModel(0))
			AllgemeineInfos.AppendChild(ComputerModellElement)
			ComputerModellElement.AppendChild(ComputerModellText)
			'COmputerSeriennummer
			Dim COmputerSeriennummerElement As XmlElement = XmlDoc.CreateElement("COmputerSeriennummer")
			Dim COmputerSeriennummerText As XmlText = XmlDoc.CreateTextNode(Var.biosSerialNumber(0))
			AllgemeineInfos.AppendChild(COmputerSeriennummerElement)
			COmputerSeriennummerElement.AppendChild(COmputerSeriennummerText)
			'Windowslizenz
			Dim WindowslizenzElement As XmlElement = XmlDoc.CreateElement("Windows_Lizenz")
			Dim WindowslizenzText As XmlText = XmlDoc.CreateTextNode(Var.softwareliceOA3xOriginalProductKey(0))
			AllgemeineInfos.AppendChild(WindowslizenzElement)
			WindowslizenzElement.AppendChild(WindowslizenzText)
			Try
				'Virenschutz
				For i As Integer = 0 To Var.antivirus.GetLength(0) - 1
					Dim VirenschutzElement As XmlElement = XmlDoc.CreateElement("Virenschutz" & i.ToString)
					Dim VirenschutzText As XmlText = XmlDoc.CreateTextNode(Var.antivirus(i))
					AllgemeineInfos.AppendChild(VirenschutzElement)
					VirenschutzElement.AppendChild(VirenschutzText)
				Next
			Catch ex As Exception
				Console.WriteLine(ex.ToString)
			End Try
			'OSversion
			Dim OSversionElement As XmlElement = XmlDoc.CreateElement("OS_Version")
			Dim OSversionText As XmlText = XmlDoc.CreateTextNode(Var.osversion)
			AllgemeineInfos.AppendChild(OSversionElement)
			OSversionElement.AppendChild(OSversionText)
			'IntallDatum
			Dim IntallDatumElement As XmlElement = XmlDoc.CreateElement("Installations_Datum")
			Dim IntallDatumText As XmlText = XmlDoc.CreateTextNode(Var.osinstalldatum)
			AllgemeineInfos.AppendChild(IntallDatumElement)
			IntallDatumElement.AppendChild(IntallDatumText)
			'LetzterBenutzer
			Dim LetzterBenutzeElement As XmlElement = XmlDoc.CreateElement("letzter_Benutzer")
			Dim LetzterBenutzeText As XmlText = XmlDoc.CreateTextNode(Var.curbenutzer)
			AllgemeineInfos.AppendChild(LetzterBenutzeElement)
			LetzterBenutzeElement.AppendChild(LetzterBenutzeText)
		Catch ex As Exception
			Console.WriteLine(ex.ToString)
		End Try
		Try
			''Mainbaord
			Dim Mainboard As XmlElement = XmlDoc.CreateElement("Mainboard")
			XmlDoc.DocumentElement.PrependChild(Mainboard)
			Try
				'MainboardHersteller
				Dim MainboardHerstellerElement As XmlElement = XmlDoc.CreateElement("Mainboard_Hersteller")
				Dim MainboardHerstellerText As XmlText = XmlDoc.CreateTextNode(Var.mainboardManufacturer(0))
				Mainboard.AppendChild(MainboardHerstellerElement)
				MainboardHerstellerElement.AppendChild(MainboardHerstellerText)
			Catch ex As Exception
			End Try
			Try
				'MainboardModell
				Dim MainboardModellElement As XmlElement = XmlDoc.CreateElement("Mainboard_Modell")
				Dim MainboardModellText As XmlText = XmlDoc.CreateTextNode(Var.mainboardProduct(0))
				Mainboard.AppendChild(MainboardModellElement)
				MainboardModellElement.AppendChild(MainboardModellText)
			Catch ex As Exception
			End Try
			Try
				'MainboardRev
				Dim MainboardRevElement As XmlElement = XmlDoc.CreateElement("Mainboard_Rev")
				Dim MainboardRevText As XmlText = XmlDoc.CreateTextNode(Var.mainboardVersion(0))
				Mainboard.AppendChild(MainboardRevElement)
				MainboardRevElement.AppendChild(MainboardRevText)
			Catch ex As Exception
			End Try
		Catch ex As Exception
			Console.WriteLine("Mainboard " & ex.ToString)
		End Try
		Try
			''MacAdresse
			Dim MACAdressAll As XmlElement = XmlDoc.CreateElement("MacAdressen")
			XmlDoc.DocumentElement.PrependChild(MACAdressAll)
			For i As Integer = 0 To Var.macadressen.GetLength(0) - 1
				Dim MACAdress As XmlElement = XmlDoc.CreateElement("MacAdresse" & i.ToString)
				XmlDoc.DocumentElement.PrependChild(MACAdress)
				'MACAdresse
				Dim MacAdresseElement As XmlElement = XmlDoc.CreateElement("MacAdresse")
				Dim MacAdresseText As XmlText = XmlDoc.CreateTextNode(Var.macadressen(i, 0))
				MACAdress.AppendChild(MacAdresseElement)
				MacAdresseElement.AppendChild(MacAdresseText)
				'MACAdressBeschreibung
				Dim MACAdressBeschreibungElement As XmlElement = XmlDoc.CreateElement("MACAdress_Beschreibung")
				Dim MACAdressBeschreibungText As XmlText = XmlDoc.CreateTextNode(Var.macadressen(i, 2))
				MACAdress.AppendChild(MACAdressBeschreibungElement)
				MACAdressBeschreibungElement.AppendChild(MACAdressBeschreibungText)
				'MACAdressName
				Dim MACAdressNameElement As XmlElement = XmlDoc.CreateElement("MacAdress_Name")
				Dim MACAdressNameText As XmlText = XmlDoc.CreateTextNode(Var.macadressen(i, 1))
				MACAdress.AppendChild(MACAdressNameElement)
				MACAdressNameElement.AppendChild(MACAdressNameText)
				MACAdressAll.AppendChild(MACAdress)
			Next
		Catch ex As Exception
		End Try
		Try
			''CPU
			For i As Integer = 0 To Var.cpuName.Length - 1 Step 1
				Dim CPU As XmlElement = XmlDoc.CreateElement("CPU" & i.ToString)
				XmlDoc.DocumentElement.PrependChild(CPU)
				'CPUModell
				Dim CPUModellElement As XmlElement = XmlDoc.CreateElement("CPU_Modell")
				Dim CPUModellText As XmlText = XmlDoc.CreateTextNode(Var.cpuName(i))
				CPU.AppendChild(CPUModellElement)
				CPUModellElement.AppendChild(CPUModellText)
				'CPUHersteller
				Dim CPUHerstellerElement As XmlElement = XmlDoc.CreateElement("CPU_Hersteller")
				Dim CPUHerstellerText As XmlText = XmlDoc.CreateTextNode(Var.cpuManufacturer(i))
				CPU.AppendChild(CPUHerstellerElement)
				CPUHerstellerElement.AppendChild(CPUHerstellerText)
				'CPUDeviceID
				Dim CPUDeviceIDElement As XmlElement = XmlDoc.CreateElement("CPU_Sockel")
				Dim CPUDeviceIDText As XmlText = XmlDoc.CreateTextNode(Var.cpuDeviceID(i))
				CPU.AppendChild(CPUDeviceIDElement)
				CPUDeviceIDElement.AppendChild(CPUDeviceIDText)
			Next
		Catch ex As Exception
		End Try
		Try
			''GPU
			For i As Integer = 0 To Var.gpuName.Length - 1 Step 1
				Dim GPU As XmlElement = XmlDoc.CreateElement("GPU" & i.ToString)
				XmlDoc.DocumentElement.PrependChild(GPU)
				'GPUName
				Dim GPUNameElement As XmlElement = XmlDoc.CreateElement("GPU_Modell")
				Dim GPUNameText As XmlText = XmlDoc.CreateTextNode(Var.gpuName(i))
				GPU.AppendChild(GPUNameElement)
				GPUNameElement.AppendChild(GPUNameText)
				'GPUHersteller
				Dim GPUHerstellerElement As XmlElement = XmlDoc.CreateElement("GPU_Hersteller")
				Dim GPUHerstellerText As XmlText = XmlDoc.CreateTextNode(Var.gpuVideoProcessor(i))
				GPU.AppendChild(GPUHerstellerElement)
				GPUHerstellerElement.AppendChild(GPUHerstellerText)
			Next
		Catch ex As Exception
		End Try
		Try
			''RAM
			For i As Integer = 0 To Var.ramPartNumber.Length - 1 Step 1
				Dim RAM As XmlElement = XmlDoc.CreateElement("RAM" & i.ToString)
				XmlDoc.DocumentElement.PrependChild(RAM)
				'RAMBankLabel
				Dim RAMBankLabelElement As XmlElement = XmlDoc.CreateElement("RAM_Slot")
				Dim RAMBankLabelText As XmlText = XmlDoc.CreateTextNode(Var.ramBankLabel(i))
				RAM.AppendChild(RAMBankLabelElement)
				RAMBankLabelElement.AppendChild(RAMBankLabelText)
				'RAMCapacity
				Dim RAMCapacityElement As XmlElement = XmlDoc.CreateElement("RAM_Größe")
				Dim RAMCapacityText As XmlText = XmlDoc.CreateTextNode(Var.ramCapacity(i))
				RAM.AppendChild(RAMCapacityElement)
				RAMCapacityElement.AppendChild(RAMCapacityText)
				'RAMManufacturer
				Dim RAMManufacturerElement As XmlElement = XmlDoc.CreateElement("RAM_Hersteller")
				Dim RAMManufacturerText As XmlText = XmlDoc.CreateTextNode(Var.ramManufacturer(i))
				RAM.AppendChild(RAMManufacturerElement)
				RAMManufacturerElement.AppendChild(RAMManufacturerText)
				'RAMPartnumber
				Dim RAMPartnumberElement As XmlElement = XmlDoc.CreateElement("RAM_Partnummer")
				Dim RAMPartnumberText As XmlText = XmlDoc.CreateTextNode(Var.ramPartNumber(i))
				RAM.AppendChild(RAMPartnumberElement)
				RAMPartnumberElement.AppendChild(RAMPartnumberText)
				'RAMSeriennummer
				Dim RAMSeriennummerElement As XmlElement = XmlDoc.CreateElement("RAM_Seriennummer")
				Dim RAMSeriennummerText As XmlText = XmlDoc.CreateTextNode(Var.ramseriennummer(i))
				RAM.AppendChild(RAMSeriennummerElement)
				RAMSeriennummerElement.AppendChild(RAMSeriennummerText)
			Next
		Catch ex As Exception
		End Try
		Try
			''Festplatte
			For i As Integer = 0 To Var.festplatteModel.Length - 1 Step 1
				Dim Festplatte As XmlElement = XmlDoc.CreateElement("Festplatte" & i.ToString)
				XmlDoc.DocumentElement.PrependChild(Festplatte)
				'FestplatteFirmware
				Dim FestplatteFirmwareElement As XmlElement = XmlDoc.CreateElement("Festplatte_Firmware")
				Dim FestplatteFirmwareText As XmlText = XmlDoc.CreateTextNode(Var.festplatteFirmwareRevision(i))
				Festplatte.AppendChild(FestplatteFirmwareElement)
				FestplatteFirmwareElement.AppendChild(FestplatteFirmwareText)
				'FestplatteModell
				Dim FestplatteModellElement As XmlElement = XmlDoc.CreateElement("Festplatte_Modell")
				Dim FestplatteModellText As XmlText = XmlDoc.CreateTextNode(Var.festplatteModel(i))
				Festplatte.AppendChild(FestplatteModellElement)
				FestplatteModellElement.AppendChild(FestplatteModellText)
				'FestplatteSerialNumber
				Dim FestplatteSerialNumberElement As XmlElement = XmlDoc.CreateElement("Festplatte_Seriennummer")
				Dim FestplatteSerialNumberText As XmlText = XmlDoc.CreateTextNode(Var.festplatteSerialNumber(i))
				Festplatte.AppendChild(FestplatteSerialNumberElement)
				FestplatteSerialNumberElement.AppendChild(FestplatteSerialNumberText)
				'FestplaltteSize
				Dim FestplaltteSizeElement As XmlElement = XmlDoc.CreateElement("Festplatte_Größe")
				Dim FestplaltteSizeText As XmlText = XmlDoc.CreateTextNode(Var.festplatteSize(i))
				Festplatte.AppendChild(FestplaltteSizeElement)
				FestplaltteSizeElement.AppendChild(FestplaltteSizeText)
			Next
		Catch ex As Exception
		End Try
		'Save to the XML file
		Try
			XmlDoc.Save(Var.logpath & Environment.MachineName & "\" & DateAndTime.Now.Day.ToString & "_" & DateAndTime.Now.Month.ToString & "_" & DateAndTime.Now.Year.ToString & "\Auswertung.xml")
		Catch ex As Exception
			XmlDoc.Save(".\Auswertung.xml")
		End Try
	End Sub
	''API
	Sub APIGETAssetID()
		Dim idmodell As Integer = APIGetModel(Var.computerModel(0), Var.computerManufacturer(0), Var.APIfieldsetid, Var.APIcategorieid(5), True, Var.computerSystemFamily(0))
		Dim Virenschutz As String = ""
		Dim i As Integer
		Do While i < Var.antivirus.Length
			Virenschutz = Virenschutz & Var.antivirus(i) & ","
			i += 1
		Loop
		Virenschutz = Virenschutz.Remove(Virenschutz.Length - 1)
		Dim macadresse As String = ""
		For j As Integer = 0 To Var.macadressen.GetLength(0) - 1
			If Var.macadressen(j, 0) <> "" Then
				macadresse = macadresse + Var.macadressen(j, 0) & ","
			End If
		Next
		Dim RestSharpclient As New RestClient(Var.APIPath)
		Dim RestSharpRequest As New RestRequest("hardware?search=" & Var.biosSerialNumber(0) & "&model_id=" & idmodell & "&components=true")
		RestSharpRequest.AddHeader("Authorization", "Bearer " & Var.APIKey)
		RestSharpRequest.AddHeader("Accept", "application/json")
		Dim RestSharpResponse As Task(Of RestResponse)
		RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequest)
		Dim Data As JObject = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
		Select Case Convert.ToInt32(Data("total"))
			Case 0
				Dim RestSharpRequestAddAsset As New RestRequest("hardware", Method.Post)
				RestSharpRequestAddAsset.AddHeader("Authorization", "Bearer " & Var.APIKey)
				RestSharpRequestAddAsset.AddHeader("Accept", "application/json")
				RestSharpRequestAddAsset.AddJsonBody(New With {Key .asset_tag = Var.computername, Key .status_id = 2, Key .model_id = idmodell, Key .name = Var.computername, Key ._snipeit_computername_6 = Var.computername, Key .serial = Var.biosSerialNumber(0), Key ._snipeit_virenschutz_5 = Virenschutz, Key ._snipeit_letzter_benutzer_7 = Var.curbenutzer, Key ._snipeit_bios_version_9 = Var.biosName(0), Key ._snipeit_mac_adresse_computer_10 = macadresse})
				RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestAddAsset)
				Data = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
				If Data("status").ToString = "success" Then
					Dim Data1 As JObject = Data("payload")
					Var.computerid = Convert.ToInt32(Data1("id"))
				Else
					Logschreiber(Log:="Fehler APIAddAsset" & vbCrLf & Data.ToString, art:="fehler", section:="APIAddAsset")
					Environment.Exit(1)
				End If
			Case 1
				Dim Data1 As JArray = Data("rows")
				For Each JObject As JObject In Data1
					Var.computerid = Convert.ToInt32(JObject.Item("id"))
					Var.APIcomponents = JObject.Item("components")
				Next
				Dim RestSharpRequestUpdateAsset As New RestRequest("hardware/" & Var.computerid.ToString, Method.Patch)
				RestSharpRequestUpdateAsset.AddHeader("Authorization", "Bearer " & Var.APIKey)
				RestSharpRequestUpdateAsset.AddHeader("Accept", "application/json")
				RestSharpRequestUpdateAsset.AddJsonBody(New With {Key .asset_tag = Var.computername, Key .name = Var.computername, Key ._snipeit_computername_6 = Var.computername, Key ._snipeit_virenschutz_5 = Virenschutz, Key ._snipeit_letzter_benutzer_7 = Var.curbenutzer, Key ._snipeit_letzter_durchlauf_8 = Now.ToString, Key ._snipeit_bios_version_9 = Var.biosName(0), Key ._snipeit_mac_adresse_computer_10 = macadresse})
				RestSharpResponse = RestSharpclient.ExecuteAsync(RestSharpRequestUpdateAsset)
				Data = JsonConvert.DeserializeObject(RestSharpResponse.Result.Content.ToString)
				If Data("status").ToString = "success" Then
					Logschreiber(Log:="Info APIUpdateAsset" & vbCrLf & "Asset erfolgreich aktualisiert", art:="Info", section:="APIUpdateAsset")
				Else
					Logschreiber(Log:="Fehler APIUpdateAsset" & vbCrLf & "Asset nicht erfolgreich aktualisiert", art:="fehler", section:="APIUpdateAsset")
				End If
			Case Else
				Logschreiber(Log:="Fehler APIGETAssetID" & vbCrLf & "Asset kann nicht Eindeutig bestimmt werden", art:="fehler", section:="APIGETAssetID")
				Environment.Exit(1)
		End Select
	End Sub
	Sub APICPU()
		Select Case Var.cpuName.Length
			Case 1
				Dim found As Boolean = False
				If Var.APIcomponents.Count <> 0 Then
					For Each JObject As JObject In Var.APIcomponents
						If JObject.Item("name").ToString = (Var.cpuName(0) & " | " & Var.cpuManufacturer(0)) Then
							found = True
							If Var.APICheckedcomponents.Length = 1 Then
								Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
							Else
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
								Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
							End If
						End If
					Next
				End If
				If found = False Then
					If APIcomponentsAdd((Var.cpuName(0) & " | " & Var.cpuManufacturer(0)), Var.APIcategorieid(0)) = False Then
						Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APICPU")
					End If
				End If
			Case > 1
				For i As Integer = 0 To Var.cpuName.Length - 1
					Dim found As Boolean = False
					If Var.APIcomponents.Count <> 0 Then
						For Each JObject As JObject In Var.APIcomponents
							If JObject.Item("name").ToString = (Var.cpuName(0) & " | " & Var.cpuManufacturer(0)) Then
								Dim found2 As Boolean = False
								For Each j As Integer In Var.APICheckedcomponents
									If j = Convert.ToInt32(JObject.Item("id")) Then
										found2 = True
									End If
								Next
								If found2 = False Then
									If Var.APICheckedcomponents.Length = 1 Then
										Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
									Else
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
										Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
									End If
									found = True
								End If
							End If
						Next
					End If
					If found = False Then
						If APIcomponentsAdd((Var.cpuName(0) & " | " & Var.cpuManufacturer(0)), Var.APIcategorieid(0)) = False Then
							Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APICPU")
						End If
					End If
				Next
		End Select
	End Sub
	Sub APIGPU()
		Select Case Var.gpuName.Length
			Case 1
				Dim found As Boolean = False
				If Var.APIcomponents.Count <> 0 Then
					For Each JObject As JObject In Var.APIcomponents
						If JObject.Item("name").ToString = (Var.gpuName(0) & " | " & Var.gpuVideoProcessor(0)) Then
							found = True
							If Var.APICheckedcomponents.Length = 1 Then
								Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
							Else
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
								Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
							End If
						End If
					Next
				End If
				If found = False Then
					If APIcomponentsAdd((Var.gpuName(0) & " | " & Var.gpuVideoProcessor(0)), Var.APIcategorieid(1)) = False Then
						Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APIGPU")
					End If
				End If
			Case > 1
				For i As Integer = 0 To Var.gpuName.Length - 1
					Dim found As Boolean = False
					If Var.APIcomponents.Count <> 0 Then
						For Each JObject As JObject In Var.APIcomponents
							If JObject.Item("name").ToString = (Var.gpuName(i) & " | " & Var.gpuVideoProcessor(i)) Then
								Dim found2 As Boolean = False
								For Each j As Integer In Var.APICheckedcomponents
									If j = Convert.ToInt32(JObject.Item("id")) Then
										found2 = True
									End If
								Next
								If found2 = False Then
									If Var.APICheckedcomponents.Length = 1 Then
										Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
									Else
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
										Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
									End If
									found = True
								End If
							End If
						Next
					End If
					If found = False Then
						If APIcomponentsAdd((Var.gpuName(i) & " | " & Var.gpuVideoProcessor(i)), Var.APIcategorieid(1)) = False Then
							Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APIGPU")
						End If
					End If
				Next
		End Select
	End Sub
	Sub APIRAM()
		Select Case Var.ramseriennummer.Length
			Case 1
				Dim found As Boolean = False
				If Var.APIcomponents.Count <> 0 Then
					For Each JObject As JObject In Var.APIcomponents
						If JObject.Item("name").ToString = (Var.ramPartNumber(0) & " | " & Var.ramManufacturer(0)) Then
							found = True
							If Var.APICheckedcomponents.Length = 1 Then
								Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
							Else
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
								Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
							End If
						End If
					Next
				End If
				If found = False Then
					If APIcomponentsAdd((Var.ramPartNumber(0) & " | " & Var.ramManufacturer(0)), Var.APIcategorieid(3), Var.ramseriennummer(0)) = False Then
						Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APIRAM")
					End If
				End If
			Case > 1
				For i As Integer = 0 To Var.ramPartNumber.Length - 1
					Dim found As Boolean = False
					If Var.APIcomponents.Count <> 0 Then
						For Each JObject As JObject In Var.APIcomponents
							If JObject.Item("name").ToString = (Var.ramPartNumber(i) & " | " & Var.ramManufacturer(i)) Then
								Dim found2 As Boolean = False
								For Each j As Integer In Var.APICheckedcomponents
									If j = Convert.ToInt32(JObject.Item("id")) Then
										found2 = True
									End If
								Next
								If found2 = False Then
									If Var.APICheckedcomponents.Length = 1 Then
										Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
									Else
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
										Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
									End If
									found = True
								End If
							End If
						Next
					End If
					If found = False Then
						If APIcomponentsAdd((Var.ramPartNumber(i) & " | " & Var.ramManufacturer(i)), Var.APIcategorieid(3), Var.ramseriennummer(i)) = False Then
							Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APIRAM")
						End If
					End If
				Next
		End Select
	End Sub
	Sub APIHDD_SSD()
		Select Case Var.festplatteSerialNumber.Length
			Case 1
				Dim found As Boolean = False
				If Var.APIcomponents.Count <> 0 Then
					For Each JObject As JObject In Var.APIcomponents
						If JObject.Item("name").ToString = Var.festplatteModel(0) Then
							found = True
							If Var.APICheckedcomponents.Length = 1 Then
								Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
							Else
								ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
								Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
							End If
						End If
					Next
				End If
				If found = False Then
					If APIcomponentsAdd(Var.festplatteModel(0), Var.APIcategorieid(2), Var.festplatteSerialNumber(0)) = False Then
						Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APIRAM")
					End If
				End If
			Case > 1
				For i As Integer = 0 To Var.festplatteSerialNumber.Length - 1
					Dim found As Boolean = False
					If Var.APIcomponents.Count <> 0 Then
						For Each JObject As JObject In Var.APIcomponents
							If JObject.Item("name").ToString = Var.festplatteModel(i) Then
								Dim found2 As Boolean = False
								For Each j As Integer In Var.APICheckedcomponents
									If j = Convert.ToInt32(JObject.Item("id")) Then
										found2 = True
									End If
								Next
								If found2 = False Then
									If Var.APICheckedcomponents.Length = 1 Then
										Var.APICheckedcomponents(0) = Convert.ToInt32(JObject.Item("id"))
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
									Else
										ReDim Preserve Var.APICheckedcomponents(Var.APICheckedcomponents.Length)
										Var.APICheckedcomponents(Var.APICheckedcomponents.Length - 1) = Convert.ToInt32(JObject.Item("id"))
									End If
									found = True
								End If
							End If
						Next
					End If
					If found = False Then
						If APIcomponentsAdd(Var.festplatteModel(i), Var.APIcategorieid(2), Var.festplatteSerialNumber(i)) = False Then
							Logschreiber(Log:="APIcomponentsAdd", art:="fehler", section:="APIRAM")
						End If
					End If
				Next
		End Select
	End Sub
	Sub ReversCheck()
		If Var.APIcomponents.Count <> 0 Then
			For Each JObject As JObject In Var.APIcomponents
				Dim found As Boolean = False
				For Each Int As Integer In Var.APICheckedcomponents
					If Convert.ToInt32(JObject.Item("id")) <> Int Then
						found = True
						Exit For
					End If
				Next
				If found = False Then
					If APICheckout_in(Convert.ToInt32(JObject.Item("id")), False) = False Then
						Logschreiber(Log:="ReversCheck", art:="fehler", section:="ReversCheck")
					End If
				End If
			Next
		End If
	End Sub
	''WMI
	Sub WMIAbfragenAllgemein()
		Var.computerManufacturer = WMIAbfrage(WMISlect:="Manufacturer", WMIFrom:="Win32_ComputerSystem")
		Var.computerModel = WMIAbfrage(WMISlect:="Model", WMIFrom:="Win32_ComputerSystem")
		Var.computerSystemFamily = WMIAbfrage(WMISlect:="SystemFamily", WMIFrom:="Win32_ComputerSystem")
		Var.computerSystemSKUNumber = WMIAbfrage(WMISlect:="SystemSKUNumber", WMIFrom:="Win32_ComputerSystem")
		Var.osinstalldatum = ManagementDateTimeConverter.ToDateTime(WMIAbfrage(WMISlect:="InstallDate", WMIFrom:="Win32_OperatingSystem")(0)).ToString("yyyy-MM-dd HH: mm:ss")
		Var.softwareliceOA3xOriginalProductKey = WMIAbfrage(WMISlect:="OA3xOriginalProductKey", WMIFrom:="SoftwareLicensingService")
		Var.softwareliceOA3xOriginalProductKeyDescription = WMIAbfrage(WMISlect:="OA3xOriginalProductKeyDescription", WMIFrom:="SoftwareLicensingService")
		Var.antivirus = WMIAbfrage(WMISlect:="displayName", WMIFrom:="antivirusproduct", WMIScope:="root\securitycenter2")
	End Sub
	Sub WMIMainboard()
		Var.mainboardManufacturer = WMIAbfrage(WMISlect:="Manufacturer", WMIFrom:="Win32_BaseBoard")
		Var.mainboardProduct = WMIAbfrage(WMISlect:="Product", WMIFrom:="Win32_BaseBoard")
		Var.mainboardVersion = WMIAbfrage(WMISlect:="Version", WMIFrom:="Win32_BaseBoard")
	End Sub
	Sub WMIBIOS()
		Var.biosManufacturer = WMIAbfrage(WMISlect:="Manufacturer", WMIFrom:="Win32_BIOS")
		Var.biosName = WMIAbfrage(WMISlect:="Name", WMIFrom:="Win32_BIOS")
		Var.biosReleaseDate = WMIAbfrage(WMISlect:="ReleaseDate", WMIFrom:="Win32_BIOS")
		Var.biosSerialNumber = WMIAbfrage(WMISlect:="SerialNumber", WMIFrom:="Win32_BIOS") 'Seriennummer Gerät
		Select Case Var.biosReleaseDate(0)
			Case ""
			Case "Keine Daten"
			Case Else
				Var.biosReleaseDate(0) = ManagementDateTimeConverter.ToDateTime(Var.biosReleaseDate(0)).ToString("yyyy-MM-dd HH:mm:ss")
		End Select
	End Sub
	Sub WMIFestplatte()
		Var.festplatteFirmwareRevision = WMIAbfrage(WMISlect:="FirmwareRevision", WMIFrom:="Win32_DiskDrive WHERE MediaType Like 'Fixed hard disk media'")
		Var.festplatteModel = WMIAbfrage(WMISlect:="Model", WMIFrom:="Win32_DiskDrive WHERE MediaType Like 'Fixed hard disk media'")
		Var.festplatteSerialNumber = WMIAbfrage(WMISlect:="SerialNumber", WMIFrom:="Win32_DiskDrive WHERE MediaType Like 'Fixed hard disk media'")
		Var.festplatteSize = WMIAbfrage(WMISlect:="Size", WMIFrom:="Win32_DiskDrive WHERE MediaType Like 'Fixed hard disk media'")
	End Sub
	Sub WMIRAM()
		Var.ramBankLabel = WMIAbfrage(WMISlect:="BankLabel", WMIFrom:="Win32_PhysicalMemory")
		Var.ramCapacity = WMIAbfrage(WMISlect:="Capacity", WMIFrom:="Win32_PhysicalMemory")
		Var.ramManufacturer = WMIAbfrage(WMISlect:="Manufacturer", WMIFrom:="Win32_PhysicalMemory")
		Var.ramPartNumber = WMIAbfrage(WMISlect:="PartNumber", WMIFrom:="Win32_PhysicalMemory")
		Var.ramseriennummer = WMIAbfrage(WMISlect:="SerialNumber", WMIFrom:="Win32_PhysicalMemory")
	End Sub
	Sub WMIGPU()
		Var.gpuName = WMIAbfrage(WMISlect:="Name", WMIFrom:="Win32_VideoController")
		Var.gpuVideoProcessor = WMIAbfrage(WMISlect:="VideoProcessor", WMIFrom:="Win32_VideoController")
	End Sub
	Sub WMICPU()
		Var.cpuName = WMIAbfrage(WMISlect:="Name", WMIFrom:="Win32_Processor")
		Var.cpuManufacturer = WMIAbfrage(WMISlect:="Manufacturer", WMIFrom:="Win32_Processor")
		Var.cpuDeviceID = WMIAbfrage(WMISlect:="DeviceID", WMIFrom:="Win32_Processor")
	End Sub
End Module
