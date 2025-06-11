# MCP Server for LabVIEW

This Repo contains a LabVIEW Assistant MCP (Model Context Protocol) Server.

It enables an Large Language Model like ChatGPT or Claude Desktop to control software on your computer through defined commands.


# Getting Started:

To run it, fist install uv:
https://docs.astral.sh/uv/
Then install claude desktop:
https://claude.ai/download
Then open command prompt from this folder and run:
uv run mcp install main.py

It should take a few seconds and show a message that it successfully added mcp to Claude Desktop.

Then restart claude (rightlick task-icon and select "quit"), and the assistant should show up in your Claude Desktop app.

# Tools

The MCP Server contains a set of functions (tools) that the LLM can execute with different parameters. In this Server, the tools are defined by Request-and-wait-for-reply-Events of a DQMH-Module in LabVIEW.

More about DQMH here:
https://labviewwiki.org/wiki/Delacor_Queued_Message_Handler

Each Request-and-wait-for-reply-Event is essentially a SubVI with some in- and outputs. But it forwards the Request to a continuously running VI in the background (the server module), which then does the desired task, and replies with some values. When the values arrive back, the model reads them and continues.
Each of these SubVIs has a description which will be included in the MCP server code, so the AI can read it. The description includes all information the AI receives about the funciton and should well explain what it does.

There is a VI in the LabVIEW Project called "Generate Python Code.vi". It can be used to automatically generate the python code for the main.py file. So if anything on the Request VIs in-/outputs or description has changed, or one was created or removed, you can update the server that way.
Make sure to run "uv run mcp isntall main.py"and restart Claude Desktop, after you made changes to the Server code.

# Currently Available Tools:

### Start_Module
- Inputs: 
- Outputs: error out,Wait for Event Sync?,Scripting Server Broadcast Events,Module Was Already Running?,Module Name
- Description: Call this function first if you want to interact with LabVIEW. It will start the labview server that you can use to do actions in LabVIEW. You can use "Stop Module" to stop the server again. Make sure to call thsi function before any other LabVIEW specific functions.

### new_vi
- Inputs: 
- Outputs: error out,result,vi_id
- Description: Creates a new VI in the LabVIEW IDE. The VI Reference is stored and returned. This Reference can later be used to e.g. add modifications to the VI.

### add_object
- Inputs: object_name, vi_reference
- Outputs: error out,result,object_id
- Description: Adds an object to the block diagram or frontpanel of the referenced vi. Get a VI reference from "New VI".
Allowed object names are:
.NET Container
.NET Refnum
2D Picture
3D Picture
Absolute Value
Accept TLS
Access Rights
ActiveX Container
Add
Add (with error terminals)
Add Array Elements
Add Entropy
Add Private Key To TLS Configuration
Add Trusted Certificate To TLS Configuration
Agitator
AllSpoll
Always Copy
And
And Array Elements
Append True/False String
Application Refnum
Array (block diagram constant)
Array (classic)
Array (modern)
Array Interface Control
Array Max & Min
Array of Strings to Path
Array Size
Array Subset
Array To Cluster
Array To Spreadsheet String
ArrayMemInfo
Assert Structural Type Match
Assert Structural Type Mismatch
Atmospheric Tank (classic)
Atmospheric Tank (modern)
Automation Close
Automation Open
Automation Refnum
Avogadro Constant (1/mol)
Base 10 Logarithm Of e
Bin #1
Bin #2
Bitpack to Array
Bluetooth Close Connection
Bluetooth Create Listener
Bluetooth Discover
Bluetooth Network Connection Refnum
Bluetooth Open Connection
Bluetooth Read
Bluetooth Wait On Listener
Bluetooth Write
Boolean Array To Number
Boolean To (0,1)
Build Array
Build Cluster Array
Build Map
Build Matrix
Build Path
Build Set
Build Waveform
Bundle
Bundle By Name
Byte Array To String
Byte Array To UTF-8 String
Byte Stream File Refnum
Call By Reference
Call Chain
Call Library Function Node
Call MATLAB Function
Call Parent Class Method
CAN Channel (block diagram constant)
CAN Channel (classic)
CAN Channel (modern)
CAN Channel (system)
CAN Interface (block diagram constant)
CAN Interface (classic)
CAN Interface (modern)
CAN Interface (system)
CAN Task (block diagram constant)
CAN Task (classic)
CAN Task (modern)
CAN Task (system)
Cancel Button (modern)
Cancel Button (system)
Cancel Notification
Carriage Return Constant
Case Structure
Cast Unit Bases
Check Valve
Checkbox (classic)
CheckBox (modern)
Checkbox (system)
Class Specifier Constant
Classic Color Box
Classic Color Ramp
Classic Horizontal Slide Switch
Classic Vertical Slide
Clear Fixed-Point Overflow Status
Close File
Close File (deprecated)
Close MATLAB Session
Close Python Session
Close Reference
Close TLS Configuration
Close Variable Connection
Closure Structure
Cluster (classic)
Cluster (modern)
Cluster Constant
Cluster To Array
Code Interface Node
Coerce To Type
Collection Size
Color Box Constant
Color Box Thin Border
Colored Rectangle (broken?)
Combo Box (classic)
Combo Box (modern)
Combo Box (system)
Combo Box Constant
Complex Conjugate
Complex To Polar
Complex To Re/Im
Compound Arithmetic
Compressor
Concatenate Strings
Conditional Disable Structure
Constructor Node
Control & Simulation Loop
Control Help Window
Control Online Help
Control Refnum
Convert Unit
Conveyer
Copy
Copy (deprecated)
Cosecant
Cosine
Cotangent
CPU Information
Create Folder
Create Network Stream Endpoint
Create Network Stream Reader Endpoint
Create Network Stream Writer Endpoint
Create Stream
Create Unique Network Stream Endpoint
Create Unique Network Stream Reader Endpoint
Create Unique Network Stream Writer Endpoint
Create User Event
Current Processor ID
Current VI's Menubar
Current VI's Path
DAQmx Analog Trigger Source (block diagram constant)
DAQmx Analog Trigger Source (classic)
DAQmx Analog Trigger Source (modern)
DAQmx Analog Trigger Source (system)
DAQmx Device Name (block diagram constant)
DAQmx Device Name (classic)
DAQmx Device Name (modern)
DAQmx Device Name (system)
DAQmx Global Channel (block diagram constant)
DAQmx Global Channel (classic)
DAQmx Global Channel (modern)
DAQmx Global Channel (system)
DAQmx Physical Channel (block diagram constant)
DAQmx Physical Channel (classic)
DAQmx Physical Channel (modern)
DAQmx Physical Channel (system)
DAQmx Scale Name (block diagram constant)
DAQmx Scale Name (classic)
DAQmx Scale Name (modern)
DAQmx Scale Name (system)
DAQmx Switch (block diagram constant)
DAQmx Switch (classic)
DAQmx Switch (modern)
DAQmx Switch (system)
DAQmx Task Name (block diagram constant)
DAQmx Task Name (classic)
DAQmx Task Name (modern)
DAQmx Task Name (system)
DAQmx Terminal (block diagram constant)
DAQmx Terminal (classic)
DAQmx Terminal (modern)
DAQmx Terminal (system)
Data Cache Size
Data Log File Refnum
Data Value Reference
Data Value Reference Constant
DataSocket Close
DataSocket Open
DataSocket Read
DataSocket Refnum
DataSocket Write
Date/Time To Seconds
Decimal Digit?
Decimal String To Number
Decimate 1D Array
DecomposeLockDataValRef
DecomposeVariant
Decrement
Default Directory
Delete
Delete (deprecated)
Delete Data Value Reference
Delete From Array
Delete Menu Items
Delete Variant Attribute
Delete Waveform Attribute
Deny Access
Dequeue Element
Destroy Stream Endpoint
Destroy User Event
DevClear
DevClearList
Device Control/Status
Device Refnum
Diagram Disable Structure
Dial (classic)
Dial (modern)
Digital Data (classic)
Digital Data (modern)
Digital Waveform Graph (classic)
Digital Waveform Graph (modern)
Direct Variable Read
Direct Variable Write
Divide
Divide (with error terminals)
Dlg Horiz Scrollbar
Dlg Mac Group Box
Dlg Numeric #1
Dlg Numeric #2
Dlg String
Dlg Vert Scrollbar
Down Arrow
Down Bow Valve (flat)
Down Bow Valve (outline)
Down Bow Valve (shaded)
Down Rect Valve (flat)
Down Rect Valve (outline)
Down Rect Valve (shaded)
Down Tee (flat)
Down Tee (shaded)
Dynamic FPGA Interface Cast
Eased Corner Rectangle
Element of Set?
Elementary Charge (C)
Empty Array?
Empty Collection?
Empty Path Constant
Empty String Constant
Empty String/Path?
Enable Menu Tracking
EnableLocal
EnableRemote
End of Line Constant
Enqueue Element
Enqueue Element At Opposite End
Enum (classic)
Enum (modern)
Enum (system)
Enum Constant
EOF
Equal To 0?
Equal?
Error Ring Constant
Event Callback Refnum
Event Data Node
Event Structure
Exchanger
Exclusive Or
Exponential
Exponential (Arg) -1
Expression Node
False Constant
Feedback Node
FieldPoint IO Point (block diagram constant)
FieldPoint IO Point (classic)
FieldPoint IO Point (modern)
FIFO Refnum
File Dialog
File Dialog (deprecated)
File Path Control (classic)
File Path Control (modern)
File Path Control (system)
File Path Indicator (classic)
File Path Indicator (modern)
File/Directory Info
File/Directory Info (deprecated)
FindLstn
FindRQS
First Call?
Fixed-Point Overflow?
Fixed-Point to Integer Cast
Flat Box
Flat Circle
Flat Down Triangle
Flat Frame
Flat Left Triangle
Flat Right Triangle
Flat Round Button
Flat Rounded Box
Flat Sequence Structure
Flat Square Button
Flat Up Triangle
Flatten To JSON
Flatten To String
Flatten To XML
Flattened String To Variant
Floating Point Equal?
Flush Event Queue
Flush File
Flush File (deprecated)
Flush Queue
Flush Stream
For Loop
Fork
Format Date/Time String
Format Into File
Format Into String
Format Value
Formula Node
FPGA Refnum to Session
Fract/Exp String To Number
Framed Color Box (classic)
Framed Color Box (modern)
Free Label
Furnace
Gauge (classic)
Gauge (modern)
Gauge for Simon
Generate Front Panel Activity
Generate Occurrence
Generate User Event
Generate User-Defined Trace Event
Get Control Values by Index
Get Datalog Position
Get Date/Time In Seconds
Get Date/Time String
Get Drag Drop Data
Get File Position
Get File Size
Get Fixed-Point Components
Get Help Window Status
Get Matrix Diagonal
Get Matrix Elements
Get Menu Item Info
Get Menu Selection
Get Menu Short Cut Info
Get Notifier Status
Get Number of Records
Get Permissions
Get Queue Status
Get Submatrix
Get Type and Creator
Get Variant Attribute
Get Volume Info
Get Waveform Attribute
Get Waveform Components
Global Variable
GPIB Clear
GPIB Initialization
GPIB Misc
GPIB Read
GPIB Serial Poll
GPIB Status
GPIB Trigger
GPIB Wait
GPIB Write
Gravitational Constant (N m2/kg2)
Greater Or Equal To 0?
Greater Or Equal?
Greater Than 0?
Greater?
Handle Peek
Handle Poke
Hex Digit?
Hexadecimal String To Number
Hopper #1
Hopper #2
Horiz Line
Horiz Pipe (flat)
Horiz Pipe (shaded)
Horizontal Button Box
Horizontal Fill Slide (classic)
Horizontal Fill Slide (modern)
Horizontal Gradient Oval
Horizontal Gradient Rectangle
Horizontal Gradient Rounded Rectangle
Horizontal Graduated Bar
Horizontal Pointer Slide (classic)
Horizontal Pointer Slide (modern)
Horizontal Pointer Slide (system)
Horizontal Progress Bar (modern)
Horizontal Progress Bar (system)
Horizontal Scrollbar
Horizontal Slide (classic)
Horizontal Slide (system)
Horizontal Smooth Box
Horizontal Splitter Bar
Horizontal Splitter Bar (Classic)
Horizontal Splitter Bar (NXG Style)
Horizontal Splitter Bar (System)
Horizontal Switch
Horizontal Tank
Horizontal Toggle Switch (classic)
Horizontal Toggle Switch (modern)
Hyperbolic Cosecant
Hyperbolic Cosine
Hyperbolic Cotangent
Hyperbolic Secant
Hyperbolic Sine
Hyperbolic Tangent
IMAQ Session
Implies
In Place Element Structure
In Range and Coerce
Include Fixed-Point Overflow Status
Increment
Index & Bundle Cluster Array
Index Array
Index String Array
Initialize Array
Inline C Node
Insert Into Array
Insert Into Map
Insert Into Set
Insert Menu Items
Integer to Fixed-Point Cast
Intensity Chart (classic)
Intensity Chart (modern)
Intensity Graph (classic)
Intensity Graph (modern)
Interleave 1D Arrays
Interpolate 1D Array
Inverse Cosecant
Inverse Cosine
Inverse Cotangent
Inverse Hyperbolic Cosecant
Inverse Hyperbolic Cosine
Inverse Hyperbolic Cotangent
Inverse Hyperbolic Secant
Inverse Hyperbolic Sine
Inverse Hyperbolic Tangent
Inverse Secant
Inverse Sine
Inverse Tangent
Inverse Tangent (2 Input)
Invoke Node
IP To String
IrDA Close Connection
IrDA Create Listener
IrDA Discover
IrDA Network Connection Refnum
IrDA Open Connection
IrDA Read
IrDA Wait On Listener
IrDA Write
IsDebuggingActive
IVI Delete Session
IVI Logical Name (block diagram constant)
IVI Logical Name (classic)
IVI Logical Name (modern)
IVI New Session
Join
Join Numbers
Junction
Knob (classic)
Knob (modern)
Label (modern)
Label (system)
Labeled Oblong Button
Labeled Rectangular Button
Labeled Round Button
Labeled Square Button
LabVIEW Object
Leak Variant Value Reference
LED Button
Left Arrow
Left Bow Valve (flat)
Left Bow Valve (outline)
Left Bow Valve (shaded)
Left Down Elbow (flat)
Left Down Elbow (shaded)
Left Pump (flat)
Left Pump (shaded)
Left Rect Valve (flat)
Left Rect Valve (outline)
Left Rect Valve (shaded)
Left Tee (flat)
Left Tee (shaded)
Left Up Elbow (flat)
Left Up Elbow (shaded)
Less Or Equal To 0?
Less Or Equal?
Less Than 0?
Less?
Lexical Class
Line (chiseled, system)
Line (chiseled, thick)
Line (chiseled, thin)
Line (thick)
Line (thin)
Line Feed Constant
Line with Arrow (thick)
Line with Arrow (thin)
List Directory
List Folder
Listbox (classic)
Listbox (modern)
Listbox (system)
Listbox Symbol Ring Constant
Load Certificates Into Memory
Load Private Key Into Memory
Local Variable
Lock Range
Logarithm Base 10
Logarithm Base 2
Logarithm Base X
Logical Shift
Look In Map
Lookup Channel Probe
Lossy Enqueue Element
Lowered Box (thick)
Lowered Rounded Box
Machine Epsilon
Make TLS Configuration Immutable
MakeAddr
Mantissa & Exponent
Map
Map Constant
Match First String
Match Pattern
Match True/False String
MathScript Call By Reference
MathScript Node
MATLAB script
MATLAB Session Refnum
Matrix Size
Max & Min
Menu Refnum
Menu Ring (classic)
Menu Ring (modern)
Menu-Popup Ring
Merge Errors
Merge Signals
Meter (classic)
Meter (modern)
Mixed Checkbox (system)
Mixed Signal Graph
Molar Gas Constant (J/(mol K))
Motion Resource (block diagram constant)
Motion Resource (classic)
Motion Resource (modern)
Motor (flat)
Motor (shaded)
Move
Move (deprecated)
Multi-Segment Pipe
Multicolumn Listbox (classic)
Multicolumn Listbox (modern)
Multicolumn Listbox (system)
Multiply
Multiply (with error terminals)
Multiply Array Elements
Multirate Structure
Natural Logarithm
Natural Logarithm (Arg +1)
Natural Logarithm Base
Natural Logarithm Of 10
Natural Logarithm Of 2
Natural Logarithm Of Pi
Negate
Negative Infinity
New Data Value Reference
New Directory
New File
New TLS Configuration
New VI
New VI Object
Not
Not A Number/Path/Refnum?
Not A Path Constant
Not a Refnum Constant
Not a Shared Variable
Not And
Not Equal To 0?
Not Equal?
Not Exclusive Or
Not Or
Notifier Refnum
Number of Cache Levels
Number To Boolean Array
Number To Decimal String
Number To Engineering String
Number To Exponential String
Number To Fractional String
Number To Hexadecimal String
Number To Octal String
Numeric Constant
Numeric Control (classic)
Numeric Control (modern)
Numeric Control (system)
Numeric Indicator (classic)
Numeric Indicator (modern)
Obtain Notifier
Obtain Queue
Occurrence Refnum
Octal Digit?
Octal String To Number
OK Button (modern)
OK Button (system)
Old VISA Open
OLE Variant
One Button Dialog
Open and Verify Variable Connection
Open Application Reference
Open Device
Open Dynamic Bitfile Reference
Open File
Open MATLAB Session
Open Python Session
Open Tank (flat classic)
Open Tank (shaded classic)
Open Variable Connection
Open Variable Connection in Background
Open VI Object Reference
Open VI Reference
Open/Create/Replace Datalog
Open/Create/Replace File
Or
Or Array Elements
Oval with Shadow
Package Matrix
PassControl
Path Constant
Path to Array of Strings
Path To String
Path Type
Pi
Pi Divided By 2
Pi Multiplied By 2
Pick Line
Pict Ring (classic)
Pict Ring (modern)
Pict Ring (system)
PixMap
PixMap (warning: dangerous)
Planck's Constant (J/Hz)
Platform Flat Square Button
Plug-In Control
Polar To Complex
Polar To Re/Im
Positive Infinity
Power Of 10
Power Of 2
Power Of X
PPoll
PPollConfig
PPollUnconfig
Preallocated Read from Binary File
Preserve Run-Time Class
Pressure Tank (flat)
Pressure Tank (shaded)
Preview Queue Element
Printable?
Property Node
Pulldown Menu Ring
Push Button
Python Node
Python Session Refnum
Queue Refnum
Quit LabVIEW
Quotient & Remainder
Race Structure
Radio Button (system)
Radio Buttons (classic)
Radio Buttons (modern)
Radio Buttons (system)
Raised Box (classic)
Raised Box (modern)
Raised Circle
Raised Down Triangle
Raised Frame
Raised Left Triangle
Raised Right Triangle
Raised Rounded Box
Raised Up Triangle
Random Number (0-1)
RcvRespMsg
Re-accept TLS
Re/Im To Complex
Re/Im To Polar
Read Datalog
Read Device
Read File
Read from Binary File
Read from Text File
Read Map Max & Min Keys
Read Multiple Elements from Stream
Read Set Max & Min
Read Single Element from Stream
Read Variable
Read Variable with Timeout
ReadStatus
Receive
ReceiveSetup
Recessed Box
Recessed Circle
Recessed Down Triangle
Recessed Frame (modern)
Recessed Frame (system)
Recessed Left Triangle
Recessed Menu Ring
Recessed Right Triangle
Recessed Rounded Box
Recessed Up Triangle
Reciprocal
Reciprocal Of e
Reciprocal Of Pi
RecomposeUnlockDataValRef (warning: dangerous)
RecomposeVariant (warning: dangerous)
Rectangular Stop Button
Refnum to Path
Refnum to Session
Region
Register Event Callback
Register For Events
Register Session
Release Notifier
Release Queue
Remove Fixed-Point Overflow Status
Remove From Map
Remove From Set
Replace Array Subset
Replace Substring
Request Deallocation
ResetSys
Reshape Array
Resize Matrix
Resource Index
Reverse 1D Array
Reverse String
Right Arrow
Right Bow Valve (flat)
Right Bow Valve (outline)
Right Bow Valve (shaded)
Right Down Elbow (flat)
Right Down Elbow (shaded)
Right Pump (flat)
Right Pump (shaded)
Right Rect Valve (flat)
Right Rect Valve (outline)
Right Rect Valve (shaded)
Right Tee (flat)
Right Tee (shaded)
Right Up Elbow (flat)
Right Up Elbow (shaded)
Rocker
Rotary Feeder
Rotate
Rotate 1D Array
Rotate Left With Carry
Rotate Right With Carry
Rotate String
Round Button
Round LED (classic)
Round LED (modern)
Round Light
Round Push Button
Round Push Button (profile)
Round Radio Button
Round Stop Button
Round To Nearest
Round Toward +Infinity
Round Toward -Infinity
Rounded Rectangle with Shadow
RT FIFO Create
RT FIFO Delete
RT FIFO Read
RT FIFO Write
Rydberg Constant (1/m)
Scale By Power Of 2
Scan From File
Scan From String
Scan String For Tokens
Scan Value
Scanned Variable Read
Scanned Variable Write
Search 1D Array
Search and Replace String
Search Variable Container
Search/Split String
Secant
Seconds To Date/Time
Seek
Select
Send
Send Notification
SendCmds
SendDataBytes
SendIFC
SendList
SendLLO
SendSetup
Session to Refnum
Set
Set Constant
Set Control Values by Index
Set Datalog Position
Set File Position
Set File Size
Set Matrix Diagonal
Set Matrix Elements
Set Menu Item Info
Set Number of Records
Set Occurrence
Set Permissions
Set Submatrix
Set Type and Creator
Set Variant Attribute
Set Waveform Attribute
SetRWLS
SetTimeOut
Shared Variable
Shared Variable Constant
Shared Variable Control (classic)
Shared Variable Control (modern)
Shared Variable to String
Sign
Simple Horizontal Slide
Simple Numeric
Simple String
Simple Vertical Slide
Sinc
Sine
Sine & Cosine
Size Handle
Slide Switch
Smoothed Oval
Smoothed Rounded Rectangle
Sort 1D Array
Sort Array of String
Speed Of Light In Vacuum (m/sec)
Spin Control (system)
Split 1D Array
Split Button
Split Number
Split Signals
Spreadsheet String To Array
Square
Square Button
Square LED (classic)
Square LED (modern)
Square Light
Square Push Button
Square Push Button (profile)
Square Radio Button
Square Root
Stacked Sequence Structure
Start TLS
State
State Diagram
Static VI Reference
Stop
Stop Button
String Constant
String Control (classic)
String Control (modern)
String Control (system)
String Indicator (classic)
String Indicator (modern)
String Length
String Subset
String To Byte Array
String To IP
String To Path
String to Shared Variable
Strip Path
SubPanel
SubPanel (Classic)
SubPanel (NXG Style)
SubPanel (System)
Subtract
Subtract (with error terminals)
Swap Bytes
Swap Values
Swap Vector Element
Swap Words
Tab Constant
Tab Control (classic)
Tab Control (modern)
Tab Control (NXG Style)
Tab Control (system)
Table (classic)
Table (modern)
Table (system)
Tag Attribute Ring
Tangent
Tank (classic)
Tank (modern)
Target Structure
TCP Close Connection
TCP Create Listener
TCP Flattened Read
TCP Flattened Write
TCP Flex Read
TCP Flex Write
TCP Network Connection Refnum
TCP Open Connection
TCP Read
TCP Wait On Listener
TCP Write
TDMS Advanced Asynchronous Read (Data Ref)
TDMS Advanced Asynchronous Write (Data Ref)
TDMS Close
TDMS Configure Asynchronous Reads (Data Ref)
TDMS Configure Asynchronous Writes (Data Ref)
TDMS Defragment
TDMS Delete Data
TDMS File
TDMS Flush
TDMS Get Asynchronous Read Status (Data Ref)
TDMS Get Asynchronous Write Status (Data Ref)
TDMS Get Properties
TDMS In Memory Close
TDMS In Memory Open
TDMS In Memory Read Bytes
TDMS List Contents
TDMS Open
TDMS Read
TDMS Refnum To File ID
TDMS Set Properties
TDMS Write
TDMS Write IP
Temporary Directory
TestSRQ
TestSys
Text & Pict Ring (classic)
Text & Pict Ring (modern)
Text Button
Text Ring (block diagram constant)
Text Ring (classic)
Text Ring (modern)
Text Ring (system)
Text to UTF-8
Thermometer (classic)
Thermometer (modern)
Threshold 1D Array
Tick Count (ms)
Time Stamp (block diagram constant)
Time Stamp Control (classic)
Time Stamp Control (modern)
Time Stamp Control (system)
Time Stamp Indicator (classic)
Time Stamp Indicator (modern)
Timed Loop
Timed Sequence
TLS configuration
TLS Connection?
To Byte Integer
To Double Precision Complex
To Double Precision Float
To Extended Precision Complex
To Extended Precision Float
To Fixed-Point
To Long Integer
To Lower Case
To More Generic Class
To More Specific Class
To OLE Variant
To Probe String
To Quad Integer
To Single Precision Complex
To Single Precision Float
To Time Stamp
To Unsigned Byte Integer
To Unsigned Long Integer
To Unsigned Quad Integer
To Unsigned Word Integer
To Upper Case
To Variant
To Word Integer
Traditional DAQ Channel (block diagram constant)
Traditional DAQ Channel (classic)
Traditional DAQ Channel (modern)
Transpose 1D Array
Transpose 2D Array
Transpose Matrix
Tree (classic)
Tree (modern)
Tree (system)
Trigger
TriggerList
Two Button Dialog
Type and Creator
Type Cast
Type Error
Type Must Be Array
Type Must Be Cluster
Type Of
Type Specialization Structure
UDP Close
UDP Multicast Open
UDP Network Connection Refnum
UDP Open
UDP Read
UDP Write
Unbitpack from Array
Unbundle
Unbundle By Name
Unflatten From JSON
Unflatten From String
Unflatten From XML
Unleak Variant Value Reference
Unpackage Matrix
Unregister For Events
Unregister Session
Up Arrow
Up Bow Valve (flat)
Up Bow Valve (outline)
Up Bow Valve (shaded)
Up Rect Valve (flat)
Up Rect Valve (outline)
Up Rect Valve (shaded)
Up Tee (flat)
Up Tee (shaded)
User Defined Refnum
User Defined Refnum Tag (block diagram constant)
User Defined Refnum Tag (classic)
User Defined Refnum Tag (modern)
User Defined Refnum Tag Flatten (classic)
User Defined Refnum Tag Flatten (modern)
User Defined Tag (block diagram constant)
User Defined Tag (classic)
User Defined Tag (modern)
User Item
UTF-8 String To Byte Array
UTF-8 to Text
Variable Class Specifier Constant
Variable Name Constant
Variant
Variant To Data
Variant To Flattened String
Vert Line
Vert Pipe (flat)
Vert Pipe (shaded)
Vert Rocker
Vertical Fill Slide (classic)
Vertical Fill Slide (modern)
Vertical Gradient Oval
Vertical Gradient Rectangle
Vertical Gradient Rounded Rectangle
Vertical Graduated Bar
Vertical Pointer Slide (classic)
Vertical Pointer Slide (modern)
Vertical Pointer Slide (system)
Vertical Progress Bar (modern)
Vertical Progress Bar (system)
Vertical Scrollbar
Vertical Slide (system)
Vertical Slide Switch (classic)
Vertical Slide Switch (modern)
Vertical Smooth Box
Vertical Splitter Bar
Vertical Splitter Bar (Classic)
Vertical Splitter Bar (NXG Style)
Vertical Splitter Bar (System)
Vertical Switch
Vertical Toggle Switch (classic)
Vertical Toggle Switch (modern)
VI Library
VI Refnum
VI Server Reference
VISA Assert Interrupt Signal
VISA Assert Trigger
VISA Assert Utility Signal
VISA Clear
VISA Close
VISA Disable Event
VISA Discard Events
VISA Enable Event
VISA Find Resource
VISA Flush I/O Buffer
VISA GPIB Command
VISA GPIB Control ATN
VISA GPIB Control REN
VISA GPIB Pass Control
VISA GPIB Send IFC
VISA In 16
VISA In 32
VISA In 64
VISA In 8
VISA Lock
VISA Map Address
VISA Map Trigger
VISA Memory Allocation
VISA Memory Allocation Ex
VISA Memory Free
VISA Move
VISA Move In 16
VISA Move In 32
VISA Move In 64
VISA Move In 8
VISA Move Out 16
VISA Move Out 32
VISA Move Out 64
VISA Move Out 8
VISA Open
VISA Out 16
VISA Out 32
VISA Out 64
VISA Out 8
VISA Peek 16
VISA Peek 32
VISA Peek 64
VISA Peek 8
VISA Poke 16
VISA Poke 32
VISA Poke 64
VISA Poke 8
VISA Read
VISA Read STB
VISA Read To File
VISA Refnum to Session
VISA Resource Name
VISA resource name (classic)
VISA resource name (modern)
VISA Set I/O Buffer Size
VISA Status Description
VISA Unlock
VISA Unmap Address
VISA Unmap Trigger
VISA USB Control In
VISA USB Control Out
VISA VXI Cmd or Query
VISA Wait on Event
VISA Write
VISA Write From File
Volume Info
Wait (ms)
Wait For Activity
Wait For Front Panel Activity
Wait for GPIB RQS
Wait on Notification
Wait on Notification from Multiple
Wait on Notification from Multiple with Notifier History
Wait on Notification with Notifier History
Wait on Occurrence
Wait Until Next ms Multiple
WaitSRQ
Waveform Chart
Waveform Chart (classic)
Waveform Chart (modern)
Waveform Graph (classic)
Waveform Graph (modern)
While Loop #1
While Loop #2
White Space?
Write Datalog
Write Device
Write File
Write Multiple Elements to Stream
Write Single Element to Stream
Write to Binary File
Write to Text File
Write Variable
Write Variable with Timeout
XY Graph (classic)
XY Graph (modern)
Y-th Root of X


### connect_objects
- Inputs: to_object_terminal_index, from_object_terminal_index, to_object_reference, from_object_reference, vi_reference
- Outputs: error out,result
- Description: Connects two terminals of two objects with a wire on the block diagram of a labview vi. To get a new VI use "new vi" to add objects to a vi use "add object".

### get_object_terminals
- Inputs: object_id
- Outputs: error out,result
- Description: Returns the Terminals Names and/or descriptions as a string as well as their Index to be used in other functions like connect objects.

### get_vi_error_list
- Inputs: vi_reference
- Outputs: error out,result
- Description: Returns the current error list (list you see when clicking the run arrow) in a text format giving information about what on the block diagram needs to be fixed. Use this to see if your actions worked.

### cleanup_vi
- Inputs: vi_reference
- Outputs: error out,result
- Description: Cleans up the block diagram of a vi referenced by reference number. 

### create_control
- Inputs: constant, terminal_index, object_id
- Outputs: error out,created_object_id
- Description: Creates a control or indicator on the terminal specified by Object ID and Terminal Index, depending on if it's a in- or output of the node. If Constant is set to True, it can only be used on node inputs and will create a constant.
It will return the reference to the created control or indicator, that is a reference to the frontpanel element. If constant is True it returns a reference to the block diagram object (constant).
When creating a control or constant you can pass a value to be written to that element.

### run_vi
- Inputs: open_frontpanel, vi_id
- Outputs: error out
- Description: Runs a VI specified by VI ID and brings the frontpanel to foreground if open_frontpanel is set to True.

### Stop_Module
- Inputs: 
- Outputs: error out
- Description: Call this tool to stop the labview server. You must use "start module" to start it again to do actions in LabVIEW. Stopping the server will delete all temporary data, so after restarting it you cannot access the objects anymore.

