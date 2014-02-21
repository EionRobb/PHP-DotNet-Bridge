' Copyright (c) 2010 four.zero.one.unauthorized@gmail.com

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

' Author: four.zero.one.unauthorized@gmail.com
' Modified by: Eion Robb <eionrobb@gmail.com>
'
' Summary: Universal COM Callable Wrapper.  Create and handle most any .Net object, 
' EventArgs or type through this Generic COM Callable Wrapper library.  Object 
' events can be subscribed to by name with fired events being enqueued in Event_Queue.
' A Destroy() method and main Current_Memory_Usage
' property are provided for aiding in memory and resource management.  Universal_CCW_Factory 
' is the main COM object you create from your language, and with that main object, use New_x methods
' to spawn new Universal_CCW_Container objects wrapping your chosen .Net object or static class/type.

' New Classes:
' * Universal_CCW.Universal_CCW_Factory
' * Universal_CCW.Universal_CCW_Container

' Dependencies:
' 1) Windows OS
' 2) .Net ~4

' References:
' 1) About this source-code, manual: http://universalccw.sourceforge.net
' 2) Project hosted on: https://sourceforge.net/projects/universalccw/
' 3) MSDN .Net Reference: http://msdn.microsoft.com/en-us/library/ms229335.aspx


Option Strict On
Option Explicit On
Option Compare Binary

Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Collections
Imports System.Reflection
Imports System


	
Namespace Universal_CCW

	
	<Guid("A04368A6-737C-475B-B3A2-CFC58A188A8D"), ComVisible(True)> _
	Public Interface IUniversal_CCW_Factory
		ReadOnly Property Assembly_Is_Loaded(asmb_long_name As String) As Boolean
		ReadOnly Property Assembly_Registry_Count() As Integer
		ReadOnly Property Assembly_Registry_Item(asmb_long_name As String) As Assembly
		ReadOnly Property Calling_App_Name() As String
		ReadOnly Property Current_Memory_Usage() As Long
		ReadOnly Property Pending_Message_Count() As Integer
		<Description("Load assembly and store it in the project Assembly Registry for later reference.  New objects, event delegates, and static references will be loaded based off these stored assemblies.")> _
			Function Load_Assembly(asmb_long_name As String) As Boolean
		<Description("Create a Universal_CCW_Container object that allows indirect interaction with named static class or type.  Suggested method for any case.")> _
			Function New_Static(asmb_long_name As String, full_class_name As String) As Universal_CCW_Container
		<Description("Create a Universal_CCW_Container object that allows indirect interaction with a new object.  Suggested method for any case.")> _
			Function New_Object(item_handle As String, asmb_long_name As String, full_class_name As String, ByVal ParamArray args As Object()) As Universal_CCW_Container
		<Description("Displays the real type name of the passed object or value.")> _
			Function Type_Name(target_item as Object) As String
		<Description("Adds a message to the top of the message queue.  The message must be a HashTable type.")> _
			Sub Enqueue_Message(event_item as HashTable)
		<Description("Removes and returns the bottom queue message.  Returns a HashTable type.")> _
			Function Consume_Message() As HashTable
			
	End Interface


	<ClassInterface(ClassInterfaceType.None), Guid("4CAACF7C-6F81-47E5-A094-5AB1F69D5A6E")> _
	Public Class Universal_CCW_Factory 
		Implements IUniversal_CCW_Factory


		Private _Assembly_Registry As new HashTable()
		
		''' <value>Returns TRUE/FALSE if assembly object is loaded (Boolean).</value>
		''' <param name="asmb_long_name">Long assembly name</param>
		ReadOnly Public Property Assembly_Is_Loaded(asmb_long_name As String) As Boolean Implements IUniversal_CCW_Factory.Assembly_Is_Loaded
			Get
				Return _Assembly_Registry.ContainsKey(asmb_long_name)
				End Get
			End Property
			
		''' <value>Returns number of loaded assemblies in this project instance (Integer).</value>
		ReadOnly Public Property Assembly_Registry_Count() As Integer Implements IUniversal_CCW_Factory.Assembly_Registry_Count
			Get
				Return _Assembly_Registry.Count
				End Get
			End Property
			
		''' <value>Returns assembly object if loaded (Reflection.Assembly).
		''' Not recommended for external use as Assembly is not wrapped.</value>
		''' <param name="asmb_long_name">Long assembly name</param>
		ReadOnly Public Property Assembly_Registry_Item(asmb_long_name As String) As Assembly Implements IUniversal_CCW_Factory.Assembly_Registry_Item
			Get
				If Not Assembly_Is_Loaded(asmb_long_name) Then
					Throw New Exception("Assembly_Registry_Item: Assembly not loaded!")
					Return Nothing
					End If
				Return Ctype(_Assembly_Registry(asmb_long_name), Assembly)
				End Get
			End Property

		''' <value>Returns current memory allocated (Long).</value>
		ReadOnly Public Property Current_Memory_Usage() As Long Implements IUniversal_CCW_Factory.Current_Memory_Usage
			Get
				Return GC.GetTotalMemory(False)
				End Get
			End Property
		
		Private _Message_Queue As new Queue()
		
		''' <value>Returns count of items in the main message queue (Integer).</value>
		ReadOnly Public Property Pending_Message_Count() As Integer Implements IUniversal_CCW_Factory.Pending_Message_Count
			Get
				Return _Message_Queue.Count
				End Get
			End Property
		
		''' <value>Returns the name of the appication calling this library.  Useful for debugging and blacklist management (String).</value>
		ReadOnly Public Property Calling_App_Name() As String Implements IUniversal_CCW_Factory.Calling_App_Name
			Get
				Dim cmd_args() As String = Environment.GetCommandLineArgs()
				Return IO.Path.GetFileName(cmd_args(0))
				End Get
			End Property
		
		
        ''' <summary>
        ''' Run additional internal code on Factory object creation, including checking the calling application 
		''' against a white/blacklist.  See Reg key [HKEY_CLASSES_ROOT\Universal_CCW.Universal_CCW_Factory\Application_Security]
		''' for application permissions.  "Disallow" reg array value contains list of applications which will prompt the current 
		''' application user to permit spawning the Factory object.
        ''' </summary>
		Public Sub New()

			If String.IsNullOrEmpty(Calling_App_Name) Then
				Throw New Exception("Constructor: Calling application info is missing.  This is possibly due to an application error or a malicious attack.  Cannot continue.")
				End If
			' Warn user if calling app is on the blacklist
			Dim disallow_reg As String() = Ctype(Microsoft.Win32.Registry.GetValue("HKEY_CLASSES_ROOT\Universal_CCW.Universal_CCW_Factory\Application_Security", "Disallow", New String(){String.Empty}), String())
			For Each disallow_app In disallow_reg
				If disallow_app = Calling_App_Name Then
					If MsgBox("WARNING: " & disallow_app & " is attempting to call potentially dangerous code.  When this code is allowed to run on websites, it will give the website full access to your system.  Due to the extreme security risks this code poses, allowing " & disallow_app & " to use it is highly discouraged.  Do you wish to allow this?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo) <> vbYes Then 
						Throw New Exception(disallow_app & " is attempting to call this library.  Blocked by user.")
						End If
					End if
				Next
				
		End Sub
		
		
		
        ''' <param name="asmb_long_name">The full assembly string for the assembly to load.</param>
        ''' <summary>
        ''' Load assembly and store it in the project Assembly Registry for later reference.
        ''' New objects, event delegates, and static references will be loaded based off these 
        ''' stored assemblies.  Not usually necessary as creating objects and types automatically loads the assembly.
        ''' </summary>
		''' <returns>True if Assembly successfully or already loaded, false with exception if not.</returns>
		Public Function Load_Assembly(asmb_long_name As String) As Boolean Implements IUniversal_CCW_Factory.Load_Assembly
			
			If Assembly_Is_Loaded(asmb_long_name) Then
				Load_Assembly = true
				Exit Function
				End If
				
			Dim asmb = Assembly.Load(asmb_long_name)
			
			If IsNothing(asmb) Then
				Throw New Exception("Load_Assembly: asmb_long_name not found in your system assembly!")
				Load_Assembly = false
				Exit Function
				End If
				
			_Assembly_Registry.Add(asmb_long_name, asmb)
			Load_Assembly = true
			
			End Function
			
			
			
	''' <param name="asmb_long_name">The long assembly name for the assembly to load.</param>
	''' <param name="full_class_name">The full class name including namespace for this static reference.</param>
	''' <summary>
	''' Create a wrapper object that refers to a static class or type.
	''' Suggested method for any case.
	''' </summary>
	''' <returns>New Universal_CCW_Container object wrapping the static class (Universal_CCW_Container).</returns>
		Public Function New_Static(asmb_long_name As String, full_class_name As String) As Universal_CCW_Container Implements IUniversal_CCW_Factory.New_Static
		
			New_Static = new Universal_CCW_Container(asmb_long_name, full_class_name, Me)
		
		End Function
			
			
			
	''' <param name="item_handle">A new handle/name for this instance.  Used primarily here for identifying event sources.</param>
	''' <param name="asmb_long_name">The long assembly name for the assembly to load.</param>
	''' <param name="full_class_name">The full class name including namespace for this static reference.</param>
	''' <summary>
	''' Create a Universal_CCW_Container object that allows indirect interaction with a new object.
	''' Suggested method for any case.
	''' </summary>
	''' <returns>New Universal_CCW_Container object wrapping the specified class instance (Universal_CCW_Container).</returns>
		Public Function New_Object(item_handle As String, asmb_long_name As String, full_class_name As String, ByVal ParamArray args As Object()) As Universal_CCW_Container Implements IUniversal_CCW_Factory.New_Object
			
			New_Object = new Universal_CCW_Container(item_handle, asmb_long_name, full_class_name, Me, args)
		
		End Function


	''' <param name="target_item">Object you wish to check the type of.</param>
	''' <summary>
	''' Function to check the type of the externally passed object.
	''' </summary>
	''' <returns>Typename of object (String).</returns>
		Public Function Type_Name(target_item as Object) As String Implements IUniversal_CCW_Factory.Type_Name
		
			Type_Name = TypeName(target_item)
		
		End Function
		
		
		
	''' <param name="message_item">A short HashTable containing a message to add to the queue.
	''' 	For fired events, this consists of the handle of the event source, event name, and 
	'''		Universal_CCW_Container-wrapped EventArgs</param>
	''' <summary>
	''' Adds a message to the the message queue.  The queue is a Queue type, while
	''' the message is a HashTable type.
	''' </summary>
		Public Sub Enqueue_Message(message_item as HashTable) Implements IUniversal_CCW_Factory.Enqueue_Message
			
			_Message_Queue.Enqueue(message_item)
			
		End Sub

		
		
	''' <summary>
	''' Removes and returns the next queue message in line.  The queue is a Queue type, while
	''' the message is a HashTable type.
	''' </summary>
	''' <returns>The next queued item in _Message_Queue (HashTable).</returns>
		Public Function Consume_Message() As HashTable Implements IUniversal_CCW_Factory.Consume_Message

			Consume_Message = Ctype(IIF(Pending_Message_Count > 0, _Message_Queue.Dequeue(), Nothing), HashTable)

		End Function


	End Class
	
	

	<Guid("A04368A6-737C-475B-B3A2-CFC58A188A8E")> _
	Public Interface IUniversal_CCW_Container
		ReadOnly Property My_Handle() as String
		ReadOnly Property My_Object() As Object
		ReadOnly Property My_Static() As Type
		ReadOnly Property My_TypeName() As String
		ReadOnly Property My_Assembly_FullName() As String
		<Description("Gets the value of the contained object's named member property.  If return val " & _
			"is an object, the object wrapped in new Universal_CCW_Container is returned.")> _
			Function Get_Property_Value(member_name As String, Optional index As Object = Nothing) As Object
		<Description("Gets the type of the value of the contained object's named member property.  Returns a string.")> _
			Function Get_Property_TypeName(member_name As String, Optional index As Object = Nothing) As String
		<Description("Sets the value of the contained object's named member property.")> _
			Function Set_Property_Value(member_name As String, new_value As Object, Optional using_method As String = "Set") As Object
		<Description("Call the contained object's named method.  Must use exact number and type of args " & _
			"that the called method expects.   If return val is an object, the object wrapped in new " & _
			"Universal_CCW_Container is returned instead.")> _
			Function Call_Method(method_name As String, ByVal ParamArray extra_args() As Object) As Object
		<Description("Gets the value of the contained type's named property or subindex.")> _
			Function Get_Static_Property_Value(member_name As String, Optional index As Object = Nothing) As Object
		<Description("Gets the type of the value of the contained type's named property or subindex.")> _
			Function Get_Static_Property_TypeName(member_name As String, Optional index As Object = Nothing) As String
		<Description("Gets the value of the contained type's named field.")> _
			Function Get_Static_Field_Value(member_name As String) As Object
		<Description("Gets the typename of the contained type's named field.")> _
			Function Get_Static_Field_TypeName(member_name As String) As String
		<Description("Call the contained static class's named method.  Must use exact number and type of args " & _
			"that the called method expects.   If return val is an object, the object wrapped in new " & _
			"Universal_CCW_Container is returned instead.")> _
			Function Call_Static_Method(method_name As String, ByVal ParamArray extra_args() As Object) As Object
		<Description("Listens for the named event of the designated object.  A fired event is enqueued as a COM " & _
			"transferable message in the Universal_CCW_Factory._Message_Queue queue, retrievable by calling " & _
			"Universal_CCW_Factory.Consume_Message().  The message is a Hashtable type with the structure: " & _
			"{source=handle of the Universal_CCW_Container that contains the object that fired the event; " & _
			"event=name of the fired event; args=Universal_CCW_Container of the EventArgs derived object}.")> _
			Sub Subscribe_To_Event(event_name As String)
		<Description("Calls Dispose on the contained object, and sets the contained object to nothing.")> _
			Sub Destroy()
			
		End Interface
		
	' Container class for wrapping all .net objects
	<ClassInterface(ClassInterfaceType.None), Guid("4CAACF7C-6F81-47E5-A094-5AB1F69D5A6F")> _
	Public Class Universal_CCW_Container
		Implements IUniversal_CCW_Container, IDisposable

		
		Private _Universal_CCW_Factory_Reference as Universal_CCW_Factory
		
		Private _My_Handle As String
		
		''' <value>Returns the handle you assigned to this instance at creation (String).</value>
		ReadOnly Public Property My_Handle() As String Implements IUniversal_CCW_Container.My_Handle
			Get
				Return _My_Handle
				End Get
			End Property

		Private _Contained_Object As Object
		
		''' <value>Returns the actual object contained within (Object).</value>
		ReadOnly Public Property My_Object() As Object Implements IUniversal_CCW_Container.My_Object
			Get
				If IsNothing(_Contained_Object) Then
					Throw New Exception("My_Object: Object not set.")
					End If
				Return _Contained_Object
				End Get
		End Property
		
		Private _Contained_Static as Type
		
		''' <value>Returns the actual static class or type contained here (Type).</value>
		ReadOnly Public Property My_Static() As Type Implements IUniversal_CCW_Container.My_Static
			Get
				If IsNothing(_Contained_Static) Then
					Throw New Exception("My_Static: Static reference not set.")
					End If
				If Not IsNothing(_Contained_Object) AND _Contained_Object.GetType().ToString() <>  _Contained_Static.ToString() Then
					Throw New Exception("My_Static: Mismatched object and static references.")
					End If
				Return _Contained_Static
				End Get
		End Property
			
		''' <value>Returns the type name for this contained object (String).</value>
		ReadOnly Public Property My_TypeName() As String Implements IUniversal_CCW_Container.My_TypeName
			Get
				Return My_Static.ToString()
				End Get
			End Property

		''' <value>Returns the full assembly name for this contained object (String).</value>
		ReadOnly Public Property My_Assembly_FullName() As String Implements IUniversal_CCW_Container.My_Assembly_FullName
			Get
				Return My_Static.Assembly.FullName
				End Get
			End Property

			

	''' <param name="item_handle">A new handle/name for this instance.  Used primarily here for identifying event sources.</param>
	''' <param name="new_object">A VB.Net object to wrap.</param>
	''' <param name="parent_reference">Universal_CCW_Factory instance, for referencing the main message queue and loaded assemblies.</param>
	''' <summary>
	''' Constructor for wrapping objects created elsewhere.  Wrap .net new_obj created from elsewhere in the environment.  item_handle is its handle/name.
	''' This is meant to be an internal method only.  Suggest using Universal_CCW_Factory.New_Object() instead.
	''' </summary>
		Public Sub New(item_handle As String, new_object as Object, parent_reference as Universal_CCW_Factory)
			Dim asmb_long_name as String = new_object.GetType().Assembly.FullName

			_Universal_CCW_Factory_Reference = parent_reference
			_Universal_CCW_Factory_Reference.Load_Assembly(asmb_long_name)
			_Contained_Object = new_object
			_Contained_Static = _Contained_Object.GetType()
			_My_Handle = item_handle
		
		End Sub


		
	''' <param name="item_handle">A new handle/name for this instance.  Used primarily here for identifying event sources.</param>
	''' <param name="asmb_long_name">The short assembly name for the assembly to load.</param>
	''' <param name="full_class_name">The full class name including namespace for this static reference.</param>
	''' <param name="parent_reference">Universal_CCW_Factory instance, for referencing the main message queue and loaded assemblies.</param>
	''' <param name="args">The arguments to the constructor.</param>
	''' <summary>
	''' New_Object constructor.  Create a new object instance from assembly reference and wrap it.  item_handle is its handle/name.
	''' This is meant to be an internal method only.  Suggest using Universal_CCW_Factory.New_Object() instead.
	''' </summary>
		Public Sub New(item_handle As String, asmb_long_name As String, full_class_name As String, parent_reference as Universal_CCW_Factory, args As Object())
			
			_My_Handle = item_handle
			_Universal_CCW_Factory_Reference = parent_reference
			_Universal_CCW_Factory_Reference.Load_Assembly(asmb_long_name)
			
			Dim asmb as Assembly = _Universal_CCW_Factory_Reference.Assembly_Registry_Item(asmb_long_name)
			
			_Contained_Static = asmb.GetType(full_class_name)
			If IsNothing(_Contained_Static) Then
				Dim typesStr As String = ""
				Dim PotentialType As Type
				For Each PotentialType in asmb.GetTypes()
					If (PotentialType.Name = full_class_name) Then
						_Contained_Static = PotentialType
					End If
					typesStr = typesStr + ", " + PotentialType.Name.ToString()
				Next PotentialType
				If IsNothing(_Contained_Static) Then
					Throw New Exception("New_Object: full_class_name [" + full_class_name + "] not found in asmb_long_name [" + asmb_long_name + "] assembly! Valid types are: " + typesStr)
				End If
			End If
			
			'If full_class_name = "FetchDrTranParams" Then
			'	Dim consts As ConstructorInfo() = _Contained_Static.GetConstructors()
			'	_Contained_Object = consts(2).Invoke({1})
			'else
			
			Try
				If IsNothing(args) Or args.Length <= 0 Then
					_Contained_Object = Activator.CreateInstance(_Contained_Static)
				Else
					_Contained_Object = Activator.CreateInstance(_Contained_Static, args)
				End If
			Catch invokeException As TargetInvocationException
				Throw invokeException.InnerException
			End Try
			'End If
			
			If IsNothing(_Contained_Object) Then
				Throw New Exception("New_Object: failed to create object!")
			End If

		End Sub
		
	''' <param name="item_handle">A new handle/name for this instance.  Used primarily here for identifying event sources.</param>
	''' <param name="asmb_long_name">The short assembly name for the assembly to load.</param>
	''' <param name="full_class_name">The full class name including namespace for this static reference.</param>
	''' <param name="parent_reference">Universal_CCW_Factory instance, for referencing the main message queue and loaded assemblies.</param>
	''' <summary>
	''' New_Object constructor.  Create a new object instance from assembly reference and wrap it.  item_handle is its handle/name.
	''' This is meant to be an internal method only.  Suggest using Universal_CCW_Factory.New_Object() instead.
	''' </summary>
		Public Sub New(item_handle As String, asmb_long_name As String, full_class_name As String, parent_reference as Universal_CCW_Factory)
			
			_My_Handle = item_handle
			_Universal_CCW_Factory_Reference = parent_reference
			_Universal_CCW_Factory_Reference.Load_Assembly(asmb_long_name)
			
			Dim asmb as Assembly = _Universal_CCW_Factory_Reference.Assembly_Registry_Item(asmb_long_name)
			
			_Contained_Static = asmb.GetType(full_class_name)						
			If IsNothing(_Contained_Static) Then
				Throw New Exception("New_Object: full_class_name not found in asmb_long_name assembly!")
				End If
				
			_Contained_Object = asmb.CreateInstance(full_class_name)
			If IsNothing(_Contained_Object) Then
				Throw New Exception("New_Object: failed to create object!")
				End If

		End Sub

		
		
	''' <param name="asmb_long_name">The short assembly name for the assembly to load.</param>
	''' <param name="full_class_name">The full class name including namespace for this static reference.</param>
	''' <param name="parent_reference">Universal_CCW_Factory instance, for referencing the main message queue and loaded assemblies.</param>
	''' <summary>
	''' New Static constructor.  Create a reference to a static class or type from named assembly and wrap it.  My_Handle will be its full class name.
	''' This is meant to be an internal method only.  Suggest using Universal_CCW_Factory.New_Static() instead.
	''' </summary>
		Public Sub New(asmb_long_name As String, full_class_name As String, parent_reference as Universal_CCW_Factory)
			
			_Universal_CCW_Factory_Reference = parent_reference
			_Universal_CCW_Factory_Reference.Load_Assembly(asmb_long_name)
			Dim asmb as Assembly = _Universal_CCW_Factory_Reference.Assembly_Registry_Item(asmb_long_name)
			
			_Contained_Static = asmb.GetType(full_class_name)
			If IsNothing(_Contained_Static) Then
				Dim typesStr As String = ""
				Dim PotentialType As Type
				For Each PotentialType in asmb.GetTypes()
					If (PotentialType.Name = full_class_name) Then
						_Contained_Static = PotentialType
					End If
					typesStr = typesStr + ", " + PotentialType.Name.ToString()
				Next PotentialType
				If IsNothing(_Contained_Static) Then
					Throw New Exception("New_Static: full_class_name [" + full_class_name + "] not found in asmb_long_name [" + asmb_long_name + "] assembly! Valid types are: " + typesStr)
				End If
			End If
			
			_My_Handle = _Contained_Static.ToString()
			
		End Sub
		
		
		
	''' <param name="new_static">VB.Net static type to wrap.</param>
	''' <param name="parent_reference">Universal_CCW_Factory instance for use with assembly lookups and event queuing.</param>
	''' <summary>
	''' Constructor for wrapping static classes called from elsewhere.  Wrap .net static class or type already referenced elsewhere in environment.
	''' This is meant to be an internal method only.  Suggest using Universal_CCW_Factory.New_Static() instead.
	''' </summary>
		Public Sub New(new_static As Type, parent_reference as Universal_CCW_Factory)
			Dim asmb_long_name as String = new_static.GetType().Assembly.FullName

			_Universal_CCW_Factory_Reference = parent_reference
			_Universal_CCW_Factory_Reference.Load_Assembly(asmb_long_name)
			_Contained_Static = new_static
			_My_Handle = _Contained_Static.ToString()
			
		End Sub
			
			
			
		Private Function Get_Property(property_name As String, index As Object) As Object
			' private class utility function returning the true object property value.
			' used in Get_Member_Value and Get_Member_TypeName public functions.
			If IsNothing(_Contained_Object) Then
				Throw New Exception("Get_Property: object not set!")
				End If
				
			If IsNothing(index) Then
				Get_Property = CallByName(My_Object, property_name, CallType.Get)
				Else
				Get_Property = CallByName(My_Object, property_name, CallType.Get, index)
				End If
		
			End Function
			
			
			
	''' <param name="property_name">The name of the property to look up.</param>
	''' <param name="index">[Optional] index if the value to be returned is a member of a collection or array.</param>	
	''' <summary>
	''' Returns the value of named member of wrapped object.  Optional index to get value
	''' of item in a collection or array property.  If return val is object, returns object wrapped
	''' in new Universal_CCW_Container.
	''' </summary>
	''' <returns>The value of the object's named member.  If object is returned, a new Universal_CCW_Container wrapping this object will be returned instead (Scalar value or Universal_CCW_Container if Object).</returns>
		Public Function Get_Property_Value(property_name As String, Optional index As Object = Nothing) As Object Implements IUniversal_CCW_Container.Get_Property_Value
			Dim target_property_value as Object = Get_Property(property_name, index)

			If Not IsReference(target_property_value) OR TypeName(target_property_value) = "String" OR IsNothing(target_property_value)
				Get_Property_Value = target_property_value
				Else
				Get_Property_Value = new Universal_CCW_Container(_My_Handle & "." & TypeName(target_property_value), target_property_value, _Universal_CCW_Factory_Reference)
				End If
				
		End Function
			
			

	''' <param name="property_name">The name of the property to look up.</param>
	''' <param name="index">[Optional] index if the value to be returned is a member of a collection or array.</param>
	''' <summary>
	''' Gets typename of named member of wrapped object.
	''' </summary>	
	''' <returns>The typename of the named member (String).</returns>
		Public Function Get_Property_TypeName(property_name As String, Optional index As Object = Nothing) As String Implements IUniversal_CCW_Container.Get_Property_TypeName

			Get_Property_TypeName = TypeName(Get_Property(property_name, index))
			
		End Function

		

	''' <param name="property_name">The name of the property or field to set.</param>
	''' <param name="new_value">The new value to assign.  If VNW_Contained_x type, the wrapped thing will automatically be used instead.</param>
	''' <param name="using_method">[Optional] Assignment method to use.  'set' by default.  Can be 'add' or other basic method if a collection is the target.  Method must be naturally supported by the property.</param>
	''' <summary>
	''' Set named member of wrapped object using named using_method of the wrapped object's member.
	''' </summary>
	''' <returns>Whatever value is normally returned by any 'add', 'set', or other method, if any (Scalar value).</returns>
		Public Function Set_Property_Value(property_name As String, new_value As Object, Optional using_method As String = "Set") As Object Implements IUniversal_CCW_Container.Set_Property_Value

		If IsNothing(_Contained_Object) Then
			Throw New Exception("Set_Property_Value: object not set!")
			End If
			
		If typeof new_value is Universal_CCW_Container Then
			Try 
				new_value = Ctype(new_value, Universal_CCW_Container).My_Object
			Catch
				new_value = Ctype(new_value, Universal_CCW_Container).My_Static
			End Try
			End If
		If TypeName(new_value) = "DBNull" Then new_value = Nothing
		
		If typeof _Contained_Object is ValueType Then
			'unbox type
			Dim valueAsType As ValueType = DirectCast(_Contained_Object, ValueType)
			Dim valueType As Type = valueAsType.GetType()
			Dim propertyInfo As PropertyInfo = valueType.GetProperty(property_name)
			If IsNothing(propertyInfo) Then
				Dim field As FieldInfo = valueType.GetField(property_name)
				if IsNothing(field) Then
					Throw New Exception("Property " + property_name + " not found")
				End if
				
				field.SetValue(valueAsType, new_value)
			Else
				propertyInfo.SetValue(valueAsType, new_value)
			End If
			
			Set_Property_Value = Nothing
		Else
			If using_method = "Set" Then
				Set_Property_Value = CallByName(_Contained_Object, property_name, CallType.Set, new_value)
			Else
				Dim member_reference = CallByName(_Contained_Object, property_name, CallType.Get)
				Set_Property_Value = CallByName(member_reference, using_method, CallType.Method, new_value)
			End If
		End If
		
		End Function
		
		Private Function Call_Args_Filter(ByVal ParamArray extra_args() As Object) As Object()
			' private utility function.  Used in Call_Method and Call_Static_Method functions.
			' extracts objects or static references from any Universal_CCW_Container object passed as an argument
			
			If extra_args.Length > 0 Then
				Dim arg_count As Integer = extra_args.Length - 1
				Dim arg_counter As Integer
				For arg_counter = 0 to arg_count
					If typeof extra_args(arg_counter) is Universal_CCW_Container Then
						Try 
							extra_args(arg_counter) = Ctype(extra_args(arg_counter), Universal_CCW_Container).My_Object
						Catch
							extra_args(arg_counter) = Ctype(extra_args(arg_counter), Universal_CCW_Container).My_Static
						End Try
					End If
					If TypeName(extra_args(arg_counter)) = "DBNull" Then extra_args(arg_counter) = Nothing
				Next
			End If
			
			Call_Args_Filter = extra_args
			
			End Function

			
			
	''' <param name="method_name">The name of the method of the wrapped object to call.</param>
	''' <param name="extra_args">[Optional] ParamArray of arguments.  If any arguments are of VNW_Contained_x type, the thing it wraps will automatically be used in its place.</param>
	''' <summary>
	''' Call named method of wrapped object.  If return val is object, returns object wrapped in 
	''' new Universal_CCW_Container.
	''' </summary>
	''' <returns>Normal return value of object's method.  If object is returned, a new Universal_CCW_Container wrapping this object will be returned instead (Scalar value or Universal_CCW_Container if Object).</returns>
		Public Function Call_Method(method_name As String, ByVal ParamArray extra_args() As Object) As Object Implements IUniversal_CCW_Container.Call_Method

			If IsNothing(_Contained_Object) Then
				Throw New Exception("Call_Method: object not set!")
				End If
				
			Dim results as Object
			
			Try
				If extra_args.Length <= 0 Then
					results = CallByName(_Contained_Object, method_name, CallType.Method)
				Else
					results = CallByName(_Contained_Object, method_name, CallType.Method, Call_Args_Filter(extra_args))
				End If
			Catch exception As TargetInvocationException
				Throw exception.InnerException
			End Try
			
			If Not IsReference(results) OR TypeName(results) = "String" OR IsNothing(results)
				Call_Method = results
				Else
				Call_Method = new Universal_CCW_Container(_My_Handle & "_" & TypeName(results), results, _Universal_CCW_Factory_Reference)
				End If

		End Function



	''' <param name="method_name">The name of the method of the wrapped object to call.</param>
	''' <param name="extra_args">[Optional] ParamArray of arguments in order of what the method naturally requires.
	''' If any arguments are of Universal_CCW_Container type, the thing it wraps will automatically be used in its place.</param>
	''' <summary>
	''' Call named method of wrapped static class.
	''' </summary>
	''' <returns>Normal return value of static class's method.  If object is returned, a new Universal_CCW_Container wrapping this object will be returned instead (Scalar value or Universal_CCW_Container if Object).</returns>
		Public Function Call_Static_Method(method_name As String, ByVal ParamArray extra_args() As Object) As Object Implements IUniversal_CCW_Container.Call_Static_Method
					
			If IsNothing(_Contained_Static) Then
				Throw New Exception("Call_Static_Method: static reference not set!")
				End If
				
			Dim returned As Object = _Contained_Static.InvokeMember(method_name, BindingFlags.InvokeMethod, Nothing, _Contained_Static, Call_Args_Filter(extra_args))

			If Not IsReference(returned) OR TypeName(returned) = "String" OR IsNothing(returned)
				Call_Static_Method = returned
				Else
				Call_Static_Method = new Universal_CCW_Container("Obj_" & TypeName(returned), returned, _Universal_CCW_Factory_Reference)
				End If
				
		End Function

		
		Private Function Get_Static_Property(member_name As String, index As Object) As Object
			' private class utility function returning the true static class property value.
			' used in Get_Static_Member_Value and Get_Static_Member_TypeName public functions.
			
			If IsNothing(_Contained_Static) Then
				Throw New Exception("Get_Static_Property: static reference not set!")
				End If

			Dim p_info As PropertyInfo = _Contained_Static.GetProperty(member_name)
			
			If IsNothing(p_info) Then
				Throw New Exception("Get_Static_Property: property not found!")
				Get_Static_Property = Nothing
				Exit Function
				End If
			If IsNothing(index) Then
				Get_Static_Property = p_info.GetValue(_Contained_Static, nothing)
				Else
				Get_Static_Property = p_info.GetValue(_Contained_Static, New Object() {index})
				End If
				
			End Function


	''' <param name="member_name">The name of the property to look up.</param>
	''' <param name="index">[Optional] index if the value to be returned is a member of a collection or array.</param>
	''' <summary>
	''' Returns the value of named property of _Contained_Static. Optional index to get value of item in a collection or array.
	''' </summary>
	''' <returns>The value of the static class's named member.  If object is returned, a new Universal_CCW_Container wrapping this object will be returned instead (Scalar value or Universal_CCW_Container if Object).</returns>
		Public Function Get_Static_Property_Value(member_name As String, Optional index As Object = Nothing) As Object Implements IUniversal_CCW_Container.Get_Static_Property_Value
			Dim target_item = Get_Static_Property(member_name, index)
			
			If Not IsReference(target_item) OR TypeName(target_item) = "String" OR IsNothing(target_item)
				Get_Static_Property_Value = target_item
				Else
				Get_Static_Property_Value = new Universal_CCW_Container("Obj_" & member_name & "_" & TypeName(target_item), target_item, _Universal_CCW_Factory_Reference)
				End If
				
		End Function


		
	''' <param name="member_name">The name of the property or field to look up.</param>
	''' <param name="index">[Optional] index if the value to be returned is a member of a collection or array.</param>
	''' <summary>
	''' Gets typename of named member of _Contained_Static.
	''' </summary>
	''' <returns>The typename of the named member (String).</returns>
		Public Function Get_Static_Property_TypeName(member_name As String, Optional index As Object = Nothing) As String Implements IUniversal_CCW_Container.Get_Static_Property_TypeName

			Get_Static_Property_TypeName = TypeName(Get_Static_Property(member_name, index))
			
		End Function

		
			
		Private Function Get_Static_Field(member_name As String) As Object
			' private class utility function returning the true static class field value.
			' used in Get_Static_Field_Value and Get_Static_Field_TypeName public functions.
			
			If IsNothing(_Contained_Static) Then
				Throw New Exception("Get_Static_Field: static reference not set!")
				End If

			Dim f_info As FieldInfo = _Contained_Static.GetField(member_name)
			
			If IsNothing(f_info) Then
				Throw New Exception("Get_Static_Field: field not found!")
				Get_Static_Field = Nothing
				Exit Function
				End If
				
			Get_Static_Field = f_info.GetValue(_Contained_Static)

			End Function
			
			
	''' <param name="member_name">The name of the field to look up.</param>
	''' <summary>
	''' Returns the value of named field (enum value, etc) of wrapped static class or type.
	''' </summary>
	''' <returns>The value of the wrapped static class's or type's named field.  If object is returned, a new Universal_CCW_Container wrapping this object will be returned instead (Scalar value or Universal_CCW_Container if Object).</returns>
		Public Function Get_Static_Field_Value(member_name As String) As Object Implements IUniversal_CCW_Container.Get_Static_Field_Value
			Dim target_item = Get_Static_Field(member_name)
			
			If Not IsReference(target_item) OR TypeName(target_item) = "String" OR IsNothing(target_item)
				Get_Static_Field_Value = target_item
				Else
				Get_Static_Field_Value = new Universal_CCW_Container("Obj_" & member_name & "_" & TypeName(target_item), target_item, _Universal_CCW_Factory_Reference)
				End If
				
		End Function


		
	''' <param name="member_name">The name of the property or field to look up.</param>
	''' <summary>
	''' Returns the type name of named field (enum value, etc) of wrapped static class or type.
	''' </summary>
	''' <returns>The typename of the named field (String).</returns>
		Public Function Get_Static_Field_TypeName(member_name As String) As String Implements IUniversal_CCW_Container.Get_Static_Field_TypeName

			Get_Static_Field_TypeName = TypeName(Get_Static_Field(member_name))
			
		End Function
		
		
		
		
	''' <param name="event_name">The name of the event of the wrapped object to subscribe to.</param>
	''' <summary>
	''' Adds event handling for the named event of this wrapped object.
	''' Fired events are enqueued in the main Universal_CCW_Factory Message Queue.
	''' The enqueued message will be a Hashtable type using the following key/value pairs:
	''' "source": the chosen handle for this wrapped wrapper
	''' "event": the name of the event
	''' "args": a new Universal_CCW_Container wrapping EventArgs or derived type.
	''' </summary>
		Public Sub Subscribe_To_Event(event_name As String) Implements IUniversal_CCW_Container.Subscribe_To_Event
		
			If IsNothing(_Contained_Object) Then
				Throw New Exception("Subscribe_To_Event: object not set!")
				End If
				
			Dim obj_type = _Contained_Object.GetType()
			Dim obj_event_info = obj_type.GetEvent(event_name, BindingFlags.Instance OR BindingFlags.Public)
			
			If IsNothing(obj_event_info) Then
				Throw New Exception("Subscribe_To_Event: event type not available!")
				Exit Sub
			End If
			
			Dim generic_event_handler = Sub (source as Object, e as Object)
				Dim event_item as New Hashtable()
				Dim args_container = new Universal_CCW_Container(_My_Handle & "." & event_name, e, _Universal_CCW_Factory_Reference)
				event_item.Add("source", _My_Handle)
				event_item.Add("event", event_name)
				event_item.Add("args", args_container)
				_Universal_CCW_Factory_Reference.Enqueue_Message(event_item)
				End Sub
				
			Dim obj_event_handler_type = obj_event_info.EventHandlerType
			Dim obj_event_handler_conv as System.Delegate
			Select Case obj_event_handler_type.ToString()
				Case "System.EventHandler"
					obj_event_handler_conv = Ctype(generic_event_handler, EventHandler)
				Case "System.ComponentModel.CancelEventHandler"
					obj_event_handler_conv = Ctype(generic_event_handler, ComponentModel.CancelEventHandler)
				Case Else
					Throw New Exception("Subscribe_To_Event: " & obj_event_handler_type.ToString() & " handler type is not supported!")
			End Select
			obj_event_info.AddEventHandler(_Contained_Object, obj_event_handler_conv)
			
		End Sub


		
		Protected Overrides Sub Finalize()
			If TypeOf _Contained_Object Is IDisposable Then Call_Method("Dispose")
			_Contained_Object = Nothing
			_Contained_Static = Nothing
			_Universal_CCW_Factory_Reference = Nothing
			
			End Sub
			
			
			
	''' <summary>
	''' Allows calling Dispose method on contained object (if method exists)
	''' </summary>
		Public Sub Dispose() Implements IDisposable.Dispose
			If IsNothing(_Contained_Object) Then
				Throw New Exception("Dispose: object not set!")
				End If
				
			If TypeOf _Contained_Object Is IDisposable Then
				Call_Method("Dispose")
				Else
				Throw New Exception("Dispose: Contained object does not have a Dispose method!")
				End If
			
			End Sub
		
		

	''' <summary>
	''' Allows more forced freeing of object resources.  Dispose() will be called on the wrapped object
	''' (if exists), the wrapped object and static class will be set to nothing, and the parent Factory
	''' reference unlinked.
	''' </summary>
		Public Sub Destroy() Implements IUniversal_CCW_Container.Destroy
			Finalize()
			
			End Sub
			
			
		End Class

	End Namespace
