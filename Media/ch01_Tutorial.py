import uno
from datetime import datetime
from scriptforge import CreateScriptService
from com.sun.star.beans import PropertyValue

def move_to_cell(dispatcher, frame, cell_address):
    # Note that we must pass a list of arguments, even if it is just one
    args = [PropertyValue(Name="ToPoint", Value=cell_address)]
    dispatcher.executeDispatch(frame, ".uno:GoToCell", "", 0, args)

def copy_paste_example(args=None):
    doc = XSCRIPTCONTEXT.getDocument()
    frame = doc.CurrentController.Frame
    # Code needed to create Uno services
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    # Move to cell A1
    move_to_cell(dispatcher, frame, "A1")
    # Copy contents
    dispatcher.executeDispatch(frame, ".uno:Copy", "", 0, [])
    # Move to cell C1
    move_to_cell(dispatcher, frame, "C1")
    # Paste contents
    dispatcher.executeDispatch(frame, ".uno:Paste", "", 0, [])

def say_hello(args=None):
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.getActiveSheet()
    cell = sheet.getCellRangeByName("A1")
    cell.setString("Hello World")

def msg_get_date(args=None):
    # Create the Basic service from ScriptForge
    bas = CreateScriptService("Basic")
    # Shows a MsgBox with the current date
    today_str = datetime.today().strftime("%d/%m/%Y")
    bas.MsgBox(f"Today is {today_str}\nHave a nice day!")

def create_writer_file(args=None):
    ui = CreateScriptService("UI")
    sf_doc = ui.CreateDocument("Writer")
    doc = sf_doc.XComponent
    doc.Text.setString("Hello World")

