# server.py ----------------------------------------------------------
from mcp.server.fastmcp import FastMCP

import pythoncom
import time
from win32com.client import VARIANT
mcp = FastMCP("LabVIEW Assistant")

# ---- lazy LabVIEW handle ------------------------------------------
_labview = None
_labview_err = None

def get_labview():
    """
    Returns a cached LabVIEW.Application COM object.
    Import and Dispatch are delayed until the very first tool call.
    """
    global _labview, _labview_err
    if _labview is None and _labview_err is None:
        try:
            import win32com.client                # imported *inside* function
            _labview = win32com.client.Dispatch("LabVIEW.Application")
        except Exception as e:
            _labview_err = e
    if _labview_err is not None:
        raise RuntimeError(f"Cannot connect to LabVIEW: {_labview_err}")
    return _labview
# -------------------------------------------------------------------

@mcp.tool()
def connect_to_labview() -> str:
    """Use this function to establish a connection to labview. The other functions will only work after this
function was called once at least."""

    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\Start Module.vi" # Path to the Start Module.vi
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ( "error in",
        "Show Main VI Diagram on Init (F)",
        "Scripting Server Broadcast Events",
        "Module Was Already Running?", 
        "Wait for Event Sync?",
        "error out",
        "Module Name"
        )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        ("",
        False,
        "",
        False,
        False,
        "",
        "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?


    return f"Successfully connected to labview, you can now use the other tools to interact with it."


@mcp.tool()
def stop_labview_server() -> str:
    """Stops the LabVIEW Server application, can be used to restart by calling connect_to_labview again."""

    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\Stop Module.vi" # Path to the Start Module.vi
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ( "error in",
        "Wait for Module to Stop? (F)",
        "Origin",
        "Timeout to Wait for Stop (s) (-1: no timeout)", 
        "error out"
        )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        ("",
        True,
        "",
        -1,
        ""
        )
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?


    return f"Successfully stopped labview server."

@mcp.tool()
def echo(text: str) -> str:
    """Echoes back the provided text."""
    return f"You said: {text}"


@mcp.tool()
def new_vi() -> str:
    """
    Creates a new VI in the LabVIEW IDE. The VI Reference is stored and returned. This Reference can later be used to e.g. add modifications to the VI.

_____
Created using DQMH Framework: Event Scripter 7.1.0.1503.The Functions Inputs are: 
    """
    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\new_vi.vi"
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ("wait for reply (T)","error in","error out","timed out?","result" )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        (True, "", "", "", "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?
    
    return param_values
@mcp.tool()
def add_object(object_name: str,vi_reference: int) -> str:
    """
    Adds an object to the block diagram or frontpanel of the referenced vi. Get a VI reference from "New VI".
Allowed object names are:
.NET Container
.NET Refnum

_____
Created using DQMH Framework: Event Scripter 7.1.0.1503.The Functions Inputs are: parameter name: "object_name" - parameter description: ""
parameter name: "vi_reference" - parameter description: ""

    """
    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\add_object.vi"
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ("wait for reply (T)","error in","object_name","vi_reference","error out","timed out?","result" )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        (True, "", object_name, vi_reference, "", "", "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?
    
    return param_values
@mcp.tool()
def connect_objects(to_object_terminal_index: int,from_object_terminal_index: int,to_object_reference: int,from_object_reference: int,vi_reference: int) -> str:
    """
    Connects two terminals of two objects with a wire on the block diagram of a labview vi. To get a new VI use "new vi" to add objects to a vi use "add object".

_____
Created using DQMH Framework: Event Scripter 7.1.0.1503.The Functions Inputs are: parameter name: "to_object_terminal_index" - parameter description: ""
parameter name: "from_object_terminal_index" - parameter description: ""
parameter name: "to_object_reference" - parameter description: ""
parameter name: "from_object_reference" - parameter description: ""
parameter name: "vi_reference" - parameter description: ""

    """
    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\connect_objects.vi"
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ("to_object_terminal_index","from_object_terminal_index","wait for reply (T)","to_object_reference","error in","from_object_reference","vi_reference","error out","timed out?","result" )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        (to_object_terminal_index, from_object_terminal_index, True, to_object_reference, "", from_object_reference, vi_reference, "", "", "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?
    
    return param_values
@mcp.tool()
def get_object_terminals(object_id: int) -> str:
    """
    Returns the Terminals Names and/or descriptions as a string as well as their Index to be used in other functions like connect objects.

_____
Created using DQMH Framework: Event Scripter 7.1.0.1503.The Functions Inputs are: parameter name: "object_id" - parameter description: ""

    """
    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\get_object_terminals.vi"
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ("wait for reply (T)","error in","object_id","error out","timed out?","result" )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        (True, "", object_id, "", "", "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?
    
    return param_values
@mcp.tool()
def get_vi_error_list(vi_reference: int) -> str:
    """
    Returns the current error list (list you see when clicking the run arrow) in a text format giving information about what on the block diagram needs to be fixed. Use this to see if your actions worked.

_____
Created using DQMH Framework: Event Scripter 7.1.0.1503.The Functions Inputs are: parameter name: "vi_reference" - parameter description: ""

    """
    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\get_vi_error_list.vi"
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ("wait for reply (T)","error in","vi_reference","error out","timed out?","result" )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        (True, "", vi_reference, "", "", "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?
    
    return param_values
@mcp.tool()
def cleanup_vi(vi_reference: int) -> str:
    """
    Cleans up the block diagram of a vi referenced by reference number. 

_____
Created using DQMH Framework: Event Scripter 7.1.0.1503.The Functions Inputs are: parameter name: "vi_reference" - parameter description: ""

    """
    lv_app  = get_labview()
    vi_path = r"G:\My Drive\Coding\custom-mcp-server\labview_assistant\LabVIEW_Server\Scripting Server\cleanup_vi.vi"
    vi      = lv_app.GetVIReference(vi_path, "", False, 0)

    # Make sure win32com knows Call2 is a method
    vi._FlagAsMethod("Call2")

    # -------- parameter containers – one INPUT, one OUTPUT -------------
    param_names  = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_BSTR,
        ("wait for reply (T)","error in","vi_reference","error out","timed out?","result" )    # Control/Indicator Names
    )

    param_values = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
        (True, "", vi_reference, "", "", "")
    )
    # -------------------------------------------------------------------

    # Call the VI as a subVI (front panel stays closed, no suspend, etc.)
    vi.Call2(param_names, param_values,
            False,   # open FP?
            False,   # close FP after call?
            False,   # suspend on call?
            False)   # bring LabVIEW to front?
    
    return param_values

if __name__ == "__main__":
    # mcp.run()
    result = connect_to_labview()
    print(result)
    result = new_vi()
    print(result)
    time.sleep(5)
    stop_labview_server()