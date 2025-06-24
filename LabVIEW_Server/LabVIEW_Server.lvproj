<?xml version='1.0' encoding='UTF-8'?>
<Project Type="Project" LVVersion="25008000">
	<Property Name="NI.LV.All.SaveVersion" Type="Str">25.0</Property>
	<Property Name="NI.LV.All.SourceOnly" Type="Bool">true</Property>
	<Item Name="My Computer" Type="My Computer">
		<Property Name="NI.SortType" Type="Int">3</Property>
		<Property Name="server.app.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.control.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="server.tcp.enabled" Type="Bool">false</Property>
		<Property Name="server.tcp.port" Type="Int">0</Property>
		<Property Name="server.tcp.serviceName" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.tcp.serviceName.default" Type="Str">My Computer/VI Server</Property>
		<Property Name="server.vi.callsEnabled" Type="Bool">true</Property>
		<Property Name="server.vi.propertiesEnabled" Type="Bool">true</Property>
		<Property Name="specify.custom.address" Type="Bool">false</Property>
		<Item Name="Modules" Type="Folder">
			<Item Name="Scripting Server.lvlib" Type="Library" URL="../Scripting Server/Scripting Server.lvlib"/>
		</Item>
		<Item Name="Testers" Type="Folder">
			<Item Name="Test Scripting Server API.vi" Type="VI" URL="../Scripting Server/Test Scripting Server API.vi"/>
			<Item Name="Test_Connect_Connector_Pane.vi" Type="VI" URL="../Scripting Server/Test_Connect_Connector_Pane.vi"/>
			<Item Name="Test_Create_Control.vi" Type="VI" URL="../Scripting Server/Test_Create_Control.vi"/>
			<Item Name="Test_EncloseSelection.vi" Type="VI" URL="../Scripting Server/Test_EncloseSelection.vi"/>
			<Item Name="Test_Rename_And_Set_Value.vi" Type="VI" URL="../Scripting Server/Test_Rename_And_Set_Value.vi"/>
			<Item Name="Test_RandomNumberGenerator.vi" Type="VI" URL="../Scripting Server/Test_RandomNumberGenerator.vi"/>
			<Item Name="Test_WireLoops.vi" Type="VI" URL="../Scripting Server/Test_WireLoops.vi"/>
			<Item Name="Test_StructureSubdiagramsAndDeleteObject.vi" Type="VI" URL="../Scripting Server/Test_StructureSubdiagramsAndDeleteObject.vi"/>
		</Item>
		<Item Name="Dependencies" Type="Dependencies"/>
		<Item Name="Build Specifications" Type="Build"/>
	</Item>
</Project>
