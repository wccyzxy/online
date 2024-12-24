import uno

def InsertText(text, replace_text=None):
    """在当前PPT幻灯片中插入文本。
    
    处理以下情况：
    1. 有选中shape但不在文字编辑状态：在shape下方创建新文本框
    2. 有选中shape且在文字编辑状态：
       - 有选中文字：替换选中文字
       - 无选中文字：在光标处插入
    3. 无选中shape：在幻灯片中央创建新文本框
    """
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    
    if not hasattr(model, "Title") or not model.getTitle().endswith('.pptx'):
        raise Exception("当前文档不是演示文稿（.pptx）！")
        
    controller = model.getCurrentController()
    current_page = controller.getCurrentPage()
    selection = controller.getSelection()
    
    def create_text_shape(text, below_shape=None):
        """创建新文本框"""
        text_shape = model.createInstance("com.sun.star.drawing.TextShape")
        text_shape.TextMaximumFrameWidth = 10000
        text_shape.TextAutoGrowWidth = True
        
        position = uno.createUnoStruct("com.sun.star.awt.Point")
        
        if below_shape:
            # 设置新文本框位置在现有shape下方
            position.X = below_shape.Position.X
            position.Y = below_shape.Position.Y + below_shape.Size.Height + 100
        else:
            # 居中放置
            position.X = (current_page.Width - text_shape.Size.Width) / 2
            position.Y = (current_page.Height - text_shape.Size.Height) / 2
            
        # 使用setPosition方法设置位置
        text_shape.setPosition(position)
        current_page.add(text_shape)
        text_shape.String = text

        return text_shape
    
    def replace_table_text(table, find_text, replace_text):
        """遍历表格中的所有单元格进行文本替换"""
        rows = table.getRows()
        columns = table.getColumns()
        for row in range(rows.getCount()):
            for col in range(columns.getCount()):
                cell = table.getCellByPosition(col, row)
                if hasattr(cell, "createTextCursor"):
                    cursor = cell.createTextCursor()
                    cursor.gotoStart(False)
                    cursor.gotoEnd(True)  # 选择所有文本
                    current_text = cursor.getString()
                    if find_text in current_text:
                        # 只在包含目标文本的单元格中进行替换
                        new_text = current_text.replace(find_text, replace_text)
                        cursor.setString(new_text)

    if selection:       
        # 检查是否在文字编辑状态
        if hasattr(selection, "Text"):            
            selection.setString(text)
            doc = XSCRIPTCONTEXT.getDocument()
            doc.Modified = True     
        elif hasattr(selection, "Count"):
            selected_shape = selection.getByIndex(0)
            if replace_text:
                if hasattr(selected_shape, "Model"):
                    replace_table_text(selected_shape.Model, replace_text, text)
                elif hasattr(selected_shape, "String"):
                    selected_shape.String = selected_shape.String.replace(replace_text, text)
            else:
                # 有选中shape但不在文字编辑状态
                create_text_shape(text, selected_shape)
    else:
        # 无选中shape，创建居中的新文本框
        create_text_shape(text)

def InsertHello(event=None):
    # Calls the InsertText function to insert the "Hello" string
    InsertText("Hello")

def InsertAndReplace(event=None):
    InsertText("Hello", "{Project Name}")

# Make InsertHello visible by the Macro Selector
g_exportedScripts = (InsertHello, InsertAndReplace)
