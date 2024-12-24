def InsertText(text):
    """Inserts the argument string into the current document.
    If there is a selection, the selection is replaced by it.
    """

    # Get the doc from the scripting context which is made available to
    # all scripts.
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    if not hasattr(model, "Text"):
        return

    # The context variable is of type XScriptContext and is available to
    # all BeanShell scripts executed by the Script Framework
    xModel = XSCRIPTCONTEXT.getDocument()

    # The writer controller impl supports the css.view.XSelectionSupplier
    # interface.
    xSelectionSupplier = xModel.getCurrentController()

    # See section 7.5.1 of developers' guide
    xIndexAccess = xSelectionSupplier.getSelection()
    count = xIndexAccess.getCount()

    if count >= 1:  # ie we have a selection
        i = 0

    while i < count:
        xTextRange = xIndexAccess.getByIndex(i)
        theString = xTextRange.getString()

        if not len(theString):
            # Nothing really selected, just insert.
            xText = xTextRange.getText()
            xWordCursor = xText.createTextCursorByRange(xTextRange)
            xWordCursor.setString(text)
            xSelectionSupplier.select(xWordCursor)
        else:
            # Replace the selection.
            xTextRange.setString(text)
            xSelectionSupplier.select(xTextRange)

        i += 1

def InsertHello(event=None):
    # Calls the InsertText function to insert the "Hello" string
    InsertText("Hello")

# Make InsertHello visible by the Macro Selector
g_exportedScripts = (InsertHello, )
