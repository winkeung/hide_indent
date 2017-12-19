import sys

# this script can be executed inside Libre Office, using uno or win32com.client (with different initialization code)
try:
    # #get the doc from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()

    import uno
    ctx = XSCRIPTCONTEXT.getComponentContext()
    smgr = ctx.ServiceManager

except:
    try:
        import socket  # only needed on win32-OOo3.0.0
        import uno

        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        # desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
        desktop = smgr.createInstance("com.sun.star.frame.Desktop")

        # access the current writer document
        # model = desktop.getCurrentComponent()
    except:
        # import win32com.client
        import comtypes.client

        # smgr = win32com.client.Dispatch("com.sun.star.ServiceManager")
        smgr = comtypes.client.CreateObject("com.sun.star.ServiceManager")
        desktop = smgr.CreateInstance("com.sun.star.frame.Desktop")

try:
    unicode
except:
    unicode = str

xSheet = None
oController = None
level = 0

# '
# 'Find no. of indentation char in ln
# '
def findNoIndentChar(ln):
    l = len(ln)
    
    for i in range(l):
        c = ln[i]
        if c == " " or c == "|" or c == "+" or c == "-" or c == "\\":
            #print "space"
            pass
        else:
            #print c
            break
    else: # end loop without break
        return -1 #incdicate a blank line
    return i

def get_struct():
    try:
        smgr._FlagAsMethod("Bridge_GetStruct")
        struct = smgr.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    except:
        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')

    return struct

#
#Find no. of indentation cell in row
#
def findNoIndentCell(start_col, end_col, row):
    
    for c in range(start_col, end_col+1):
        s = xSheet.getCellByPosition(c, row).getString()
        if s == "" :
            #print "empty cell"
            pass
        else:
            #print s
            break
    else:
        return -1 #incdicate a blank row

    return  c - start_col


def group_recursive(
        col,
        row,
        indent_cell,
        indent_char,
        end_col,
        end_row):
    """A recursive function, group items below with deeper indentation and stop until a equal or lesser intended row or end row is encountered.
        parameters:
            col           -- input col
            row           -- input row
            indent_cell
            indent_char   -- indent level of the input row (indent_cell : indent_char) when indent_cell is the same, compare with indent_char
            end_col
            end_row

        return value:
            row            -- this row is equal or lesser indented then the input row
            indent_cell
            indent_char    -- this is the indent level of the above row
            blank_row_cnt -- no. of blank rows above
    """
    global level

    last_row = row
    blank_row_cnt = 0

    while True: # skip blank rows (and count how many of them) and also check if end_row is reached
        last_row = last_row + 1
        if last_row > end_row:
            row = last_row
            indent_cell = -1
            return row, indent_cell, indent_char, blank_row_cnt

        last_indent_cell = findNoIndentCell(col, end_col, last_row)
        last_indent_char = 0

        if 0 <= last_indent_cell:
            last_indent_char = findNoIndentChar(xSheet.getCellByPosition(col + last_indent_cell, last_row).getString())
            if 0 <= last_indent_char:
                break # not blank row

        blank_row_cnt = blank_row_cnt + 1

    level = level + 1  # set before the 1st call to group()

    while indent_cell < last_indent_cell or ((indent_cell == last_indent_cell) and (indent_char < last_indent_char)):  # next item is deeper indented
        # isBlankLine = False
        last_row, last_indent_cell, last_indent_char, blank_row_cnt = group_recursive(col, last_row, last_indent_cell, last_indent_char, end_col, end_row)

    # do selection and grouping
    if row + 1 <= last_row - 1 - blank_row_cnt:
        oRange = xSheet.getCellRangeByPosition(0, row + 1, 0, last_row - 1 - blank_row_cnt)
        oController.select(oRange)

        # ----------------------------------------------------------------------
        # get access to the document
        model = desktop.getCurrentComponent()
        document = model.getCurrentController()
        dispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")

        # ----------------------------------------------------------------------
        struct = get_struct()

        struct.Name = 'RowOrCol'
        struct.Value = 'R'

        if level < 8:  # Libre Office Calc only support max 7 levels of nested groups
            dispatcher.executeDispatch(document, ".uno:Group", "", 0, tuple([struct]))

            # if level < 3 :
            #	print level
            #

    level = level - 1

    # row = last_row
    # row_indent = last_indent_cell
    return last_row, last_indent_cell, last_indent_char, blank_row_cnt


def group_selection():
    """Group the selected rows according to their indentations.
        # selection area: | range
        # -----------------+-----------------------------------
        # 1 row           | from selected row to whole sheet
        # >1 row          | selected rows
    """

    #get the first sheet of the spreadsheet doc
    #xSheet = ThisComponent.Sheets(iCurSheet)
    # xSheet=thiscomponent.getcurrentcontroller.activesheet

    doc = desktop.getCurrentComponent()
    global oController
    oController = doc.CurrentController
    global xSheet
    xSheet = doc.CurrentController.getActiveSheet()
    
    oSel = doc.getCurrentSelection() #or oView.getSelection()
    addr = oSel.getRangeAddress()
    
    #print addr.StartRow, addr.EndRow, addr.StartColumn, addr.EndColumn
    # dim col as long, row as long, row_indent as long, end_col as long, end_row as long
    col = addr.StartColumn
    row = addr.StartRow
    
    if addr.EndRow == addr.StartRow :
        c = xSheet.createCursor()
        c.gotoEndOfUsedArea(False)
        
        # end_col = c.RangeAddress.EndColumn
        end_row = c.RangeAddress.EndRow
    else:
        # end_col = addr.EndColumn
        end_row = addr.EndRow

    end_col = c.RangeAddress.EndColumn

    global level
    level = 0

    # blank_row_cnt = 0 #any value will do, this ByRef parameter serve as output only
    
    # if addr.StartColumn == addr.EndColumn : #single column is selected, char indentation mode
    #     row_indent = findNoIndentChar(xSheet.getCellByPosition(col, row).getString())
    #
    #     while True:
    #         row, row_indent, blank_row_cnt = group(
    #             col,
    #             row,
    #             row_indent,
    #             end_row)
    #         # print "here"
    #         if row_indent < 0:
    #             break
    # else:
    #     row_indent = findNoIndentCell(col, end_col, row)
    #
    #     while True:
    #         row, row_indent, blank_row_cnt = group_cell_indent(
    #             col,
    #             row,
    #             row_indent,
    #             end_col,
    #             end_row)
    #         if row_indent < 0:
    #             break

    indent_cell = findNoIndentCell(col, end_col, row)
    indent_char = findNoIndentChar(xSheet.getCellByPosition(col+indent_cell, row).getString())

    while True:
        row, indent_cell, indent_char, blank_row_cnt = group_recursive(
            col,
            row,
            indent_cell,
            indent_char,
            end_col,
            end_row)
        if indent_cell < 0:
            break
        # print "here"

def select(scol, srow, lcol, lrow):
    #'dim oSheet, oRange, oCell, oController
    model = desktop.getCurrentComponent()
    oController = model.getCurrentController()
    #oSheet = model.sheets(1)
    #oSheet = model.Sheets.getByIndex(0)  # access the active sheet
    oSheet = model.CurrentController.ActiveSheet
    #oRange = oSheet.getCellRangeByname("B2:D3")
    oRange = oSheet.getCellRangeByPosition(scol, srow, lcol, lrow)
    oController.select(oRange)

def set_selection_visible(isVisible):
    model = desktop.getCurrentComponent()
    document = model.getCurrentController()
    dispatcher = smgr.createInstance( "com.sun.star.frame.DispatchHelper")

    struct = get_struct()

    struct.Name = 'RowOrCol'
    struct.Value = 'R'

    if isVisible:
        cmd_str = ".uno:ShowRow"
    else:
        cmd_str = ".uno:HideRow"

    dispatcher.executeDispatch(document, cmd_str, "", 0, tuple([struct]))

def set_rows_visible(start_row, no_of_row, isVisible):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    Rows = sheet.Rows
    for r in range(start_row, start_row + no_of_row):
        try:
            Rows[r].IsVisible = isVisible
            # print("IsVisible")
        except:
            # backup current selection
            oSelection = doc.getCurrentSelection()
            oArea = oSelection.getRangeAddress()
            frow = oArea.StartRow
            lrow = oArea.EndRow
            fcol = oArea.StartColumn
            lcol = oArea.EndColumn

            select(fcol, start_row, fcol, start_row + no_of_row - 1)
            set_selection_visible(isVisible)
            # print("set_selection_visiable")

            # restore previous selection
            select(fcol, frow, lcol, lrow)
            break

def check_row_visible(r):
    doc = desktop.getCurrentComponent()
    sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    Row = sheet.Rows[r]
    # sheet.Rows.hideByIndex(i,1)
    return Row.IsVisible

def hide_selection():
    """Cycle between collapse all, expand one level, and expand all (treat indentation as indicator of tree relationship).
        Walk thru all rows and expand all immediate children if not yet expanded and check whether it is originally
        all expanded or all collapsed. And then do the following actions:

        all already collapsed:     expand 1 level (already done)
        all already expanded:      collapse all
        none of the above:         expand all
    """

    doc = desktop.getCurrentComponent()
    global oController
    oController = doc.CurrentController
    global xSheet
    xSheet = doc.CurrentController.getActiveSheet()

    oSel = doc.getCurrentSelection()  # or oView.getSelection()
    addr = oSel.getRangeAddress()

    # print addr.StartRow, addr.EndRow, addr.StartColumn, addr.EndColumn
    # dim col as long, row as long, row_indent as long, end_col as long, end_row as long
    col = addr.StartColumn
    row = addr.StartRow

    c = xSheet.createCursor()
    c.gotoEndOfUsedArea(False)

    end_row = c.RangeAddress.EndRow
    end_col = c.RangeAddress.EndColumn

    indent_cell = findNoIndentCell(col, end_col, row)
    indent_char = findNoIndentChar(xSheet.getCellByPosition(col + indent_cell, row).getString())

    isAlreadyAllExpanded = True
    isAlreadyAllCollapsed = True
    last_row = row
    child_indent_cell = 1025 # lastest encountered immediate child's indentation, init to a impossible big number
    child_indent_char = 0
    blank_row_cnt = 0 # no. blank row above the current visiting row
    while True:
        last_row += 1
        if end_row < last_row:
            break

        last_indent_cell = findNoIndentCell(col, end_col, last_row)
        last_indent_char = 0

        if 0 <= last_indent_cell:
            last_indent_char = findNoIndentChar(xSheet.getCellByPosition(col + last_indent_cell, last_row).getString())
            if 0 <= last_indent_char: # not blank row
                if indent_cell < last_indent_cell or (
                    (indent_cell == last_indent_cell) and (indent_char < last_indent_char)):  # next item is deeper indented
                    if check_row_visible(last_row):
                        # print ("here")
                        isAlreadyAllCollapsed = False
                    else:
                        # print ("there")
                        isAlreadyAllExpanded = False
                    if child_indent_cell < last_indent_cell or (
                                (child_indent_cell == last_indent_cell) and (
                                child_indent_char < last_indent_char)):  # next item is deeper indented then lastest encountered immediate child
                        pass
                    else: # encounter a new immediate child
                        child_indent_cell = last_indent_cell
                        child_indent_char = last_indent_char
                        set_rows_visible(last_row - blank_row_cnt, blank_row_cnt + 1, True)
                else:
                    break # this row is not deeper indented
                blank_row_cnt = 0
            else:
                blank_row_cnt += 1 # blank row
        else:
            blank_row_cnt += 1 # blank row

    if isAlreadyAllExpanded:
        # collapse all
        set_rows_visible(row + 1, last_row - row - 1 - blank_row_cnt, False)
    elif not isAlreadyAllCollapsed:
        # expand all
        set_rows_visible(row + 1, last_row - row - 1 - blank_row_cnt, True)

if __name__ == "__main__":
    group_selection()
