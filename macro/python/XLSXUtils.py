def InsertText(text, replace_text=""):
    """在当前电子表格中插入文本。
    如果有选中单元格或单元格中的部分文本，则替换选中内容。
    """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    ctx = XSCRIPTCONTEXT.getComponentContext()
    
    # 检查是否是电子表格文档
    if not hasattr(model, "Title") or not model.getTitle().endswith('.xlsx'):
        raise Exception("当前文档不是电子表格（.xlsx）！")
        
    controller = model.getCurrentController()
    active_sheet = controller.getActiveSheet()
    selection = model.getCurrentSelection()    
    
    dispatcher = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.frame.DispatchHelper", ctx)

    dispatcher.executeDispatch(controller.getFrame(), ".uno:Cancel", "", 0, ())
    if replace_text == "":
        selection.setFormula(text)
    else:
        the_text = selection.getString()
        replace_text = the_text.replace(replace_text, text)
        selection.setString(replace_text)

def InsertHello(event=None):
    # Calls the InsertText function to insert the "Hello" string
    InsertText("Hello")

def InsertAndReplace(event=None):
    InsertText("Hello", "中（B-F列）")

# Make InsertHello visible by the Macro Selector
g_exportedScripts = (InsertHello, InsertAndReplace)
