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


doc = None
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

rows = None

def getStringByPosition(c, r):
    # return xSheet.getCellByPosition(c, r).getString()
    global rows
    if rows == None:
        # doc = desktop.getCurrentComponent()
        # sheet = doc.CurrentController.getActiveSheet()

        cursor = xSheet.createCursor()
        cursor.gotoEndOfUsedArea(False)

        end_row = cursor.RangeAddress.EndRow
        end_col = cursor.RangeAddress.EndColumn

        # get real range to extract data
        oRange = xSheet.getCellRangeByPosition(0, 0, end_col, end_row)

        # Extract cell contents as DataArray
        rows = oRange.getDataArray()
    return unicode(rows[r][c])


def findNoIndent(start_col, end_col, row):
    """Find no. of indentation cell and char in row.
    :param start_col:
    :param end_col:
    :param row:
    :return: indent column, indent char
    """
    for c in range(start_col, end_col+1):
        # s = xSheet.getCellByPosition(c, row).getString()
        s = getStringByPosition(c, row)
        if s == "" :
            #print "empty cell"
            pass
        else:
            #print s
            break
    else:
        return -1, -1 #incdicate a blank row

    l = len(s)

    for i in range(l):
        ch = s[i]
        if ch == " " or ch == "|" or ch == "+" or ch == "-" or ch == "\\":
            # print "space"
            pass
        else:
            # print c
            break
    else:  # end loop without break
        return -1, -1  # incdicate a blank line

    return  c - start_col, i

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

def arrow_down():
    model = desktop.getCurrentComponent()
    document = model.getCurrentController()
    dispatcher = smgr.createInstance( "com.sun.star.frame.DispatchHelper")

    cmd_str = ".uno:GoDown"

    structs = (get_struct(), get_struct())

    structs[0].Name = 'By'
    structs[0].Value = '1'
    structs[1].Name = 'Sel'
    structs[1].Value = 'false'

    dispatcher.executeDispatch(document, cmd_str, "", 0, structs)

def set_rows_visible(start_row, no_of_row, isVisible):
    doc = desktop.getCurrentComponent()
    # sheet = doc.CurrentController.getActiveSheet()

    # Col = sheet.Columns[1]
    # Rows = sheet.Rows
    # for r in range(start_row, start_row + no_of_row):
    #     try:
    #         Rows[r].IsVisible = isVisible
    #         # print("IsVisible")
    #     except:

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
    # select(fcol, frow, lcol, lrow)
    # break

def next_visible_row(r):
    # doc = desktop.getCurrentComponent()
    # sheet = doc.CurrentController.getActiveSheet()
    #
    # # Col = sheet.Columns[1]
    # Row = sheet.Rows[r]
    # # sheet.Rows.hideByIndex(i,1)
    # return Row.IsVisible

    model = desktop.getCurrentComponent()
    document = model.getCurrentController()

    # backup current selection
    doc = desktop.getCurrentComponent()
    oSelection = doc.getCurrentSelection()
    oArea = oSelection.getRangeAddress()
    frow = oArea.StartRow
    lrow = oArea.EndRow
    fcol = oArea.StartColumn
    lcol = oArea.EndColumn

    # select previous/next row and then send a arrow key down/up event
    if r > 0:
        select(fcol, r - 1, fcol, r - 1)
        cmd_str = ".uno:GoDown"
    else:
        select(fcol, r + 1, fcol, r + 1)
        cmd_str = ".uno:GoUp"

    dispatcher = smgr.createInstance( "com.sun.star.frame.DispatchHelper")

    structs = (get_struct(), get_struct())

    structs[0].Name = 'By'
    structs[0].Value = '1'
    structs[1].Name = 'Sel'
    structs[1].Value = 'false'

    dispatcher.executeDispatch(document, cmd_str, "", 0, structs)

    # get current selection
    oSelection = doc.getCurrentSelection()
    oArea = oSelection.getRangeAddress()

    # restore previous selection
    select(fcol, frow, lcol, lrow)

    return oArea.StartRow

def hide_selection():
    """Cycle between collapse all, expand one level, and expand all (treat indentation as indicator of tree relationship).
        Walk thru all rows and expand all immediate children(not grand/grandgrand.. child) if not yet expanded and check
        whether it is originally all expanded or all collapsed. And then do the following actions:

        all already collapsed:     expand 1 level (already done)
        all already expanded:      collapse all (checking of this condition takes a lot of time)
        if grand child exist, at least one grand child expanded    collpase all
        else if grand child not exist, at least one child expanded      collapse all
        none of the above:         expand all
    """
    global rows
    rows = None
    global doc
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
    #print row

    # backup up current selection
    frow = addr.StartRow
    lrow = addr.EndRow
    fcol = addr.StartColumn
    lcol = addr.EndColumn

    c = xSheet.createCursor()
    c.gotoEndOfUsedArea(False)

    end_row = c.RangeAddress.EndRow
    end_col = c.RangeAddress.EndColumn

    # indent_cell = findNoIndentCell(col, end_col, row)
    # indent_char = findNoIndentChar(xSheet.getCellByPosition(col + indent_cell, row).getString())
    indent_cell, indent_char = findNoIndent(col, end_col, row)
    next_visible_r = next_visible_row(row)

    isUnHideRowFound = False
    isUnHideGrandChildFound = False
    isGrandChildFound = False
    last_row = row
    child_indent_cell = 1025 # lastest encountered immediate child's indentation, init to a impossible big number
    child_indent_char = 0
    blank_row_cnt = 0 # no. blank row above the current visiting row
    while True:
        last_row += 1
        if end_row < last_row:
            break

        # last_indent_cell = findNoIndentCell(col, end_col, last_row)
        # last_indent_char = 0

        last_indent_cell, last_indent_char = findNoIndent(col, end_col, last_row)

        if (0 <= last_indent_cell) and (0 <= last_indent_char): # not blank row
            if indent_cell < last_indent_cell or (
                (indent_cell == last_indent_cell) and (indent_char < last_indent_char)):  # next item is deeper indented
                if not isUnHideRowFound:
                    while True:
                        if last_row < next_visible_r:
                            # print ("exp")
                            break
                        elif last_row == next_visible_r:
                            # print ("col")
                            isUnHideRowFound = True
                            break
                        else:
                            # print("next visible")
                            next_visible_r = next_visible_row(last_row)

                if child_indent_cell < last_indent_cell or (
                            (child_indent_cell == last_indent_cell) and (
                            child_indent_char < last_indent_char)):  # next item is deeper indented then lastest encountered immediate child
                    isGrandChildFound = True
                    if not isUnHideGrandChildFound:
                        while True:
                            if last_row < next_visible_r:
                                # print ("exp")
                                break
                            elif last_row == next_visible_r:
                                # print ("col")
                                isUnHideGrandChildFound = True
                                break
                            else:
                                # print("next visible")
                                next_visible_r = next_visible_row(last_row)

                else: # encounter a new immediate child
                    child_indent_cell = last_indent_cell
                    child_indent_char = last_indent_char
                    set_rows_visible(last_row - blank_row_cnt, blank_row_cnt + 1, True)
            else:
                break # this row is not deeper indented
            blank_row_cnt = 0
        else:
            blank_row_cnt += 1 # blank row
    #print("here")
    if isUnHideRowFound:
        if isUnHideGrandChildFound or (not isGrandChildFound):
            # collapse all
            set_rows_visible(row + 1, last_row - row - 1 - blank_row_cnt, False)
        else:
            # expand all
            set_rows_visible(row + 1, last_row - row - 1 - blank_row_cnt, True)

    # restore previous selection
    select(fcol, frow, lcol, lrow)

def hide_all_elder_brothers():
    """Thide all elder brothers (treat indentation as indicator of tree relationship).
        Walk thru all rows and expand all immediate children(not grand/grandgrand.. child) if not yet expanded and check
        whether it is originally all expanded or all collapsed. And then do the following actions:

        all already collapsed:     expand 1 level (already done)
        all already expanded:      collapse all (checking of this condition takes a lot of time)

    if not all collapsed
        if grand child exist, at least one grand child expanded    collpase all
        else if grand child not exist, at least one child expanded      collapse all
        else:         expand all
    """
    global rows
    rows = None
    global doc
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

    # backup up current selection
    frow = addr.StartRow
    lrow = addr.EndRow
    fcol = addr.StartColumn
    lcol = addr.EndColumn

    c = xSheet.createCursor()
    c.gotoEndOfUsedArea(False)

    end_row = c.RangeAddress.EndRow
    end_col = c.RangeAddress.EndColumn

    # indent_cell = findNoIndentCell(col, end_col, row)
    # indent_char = findNoIndentChar(xSheet.getCellByPosition(col + indent_cell, row).getString())
    indent_cell, indent_char = findNoIndent(col, end_col, row)
    next_visible_r = next_visible_row(row)

    isUnHideRowFound = False
    isUnHideGrandChildFound = False
    isGrandChildFound = False
    last_row = row
    child_indent_cell = 1025 # latest encountered immediate child's indentation, init to a impossible big number
    child_indent_char = 0
    blank_row_cnt = 0 # no. blank row above the current visiting row
    while True:
        last_row -= 1
        if last_row < 0:
            break

        last_indent_cell, last_indent_char = findNoIndent(col, end_col, last_row)

        if (0 <= last_indent_cell) and (0 <= last_indent_char): # not blank row
            if last_indent_cell < indent_cell or (
                (indent_cell == last_indent_cell) and (last_indent_char < indent_char)):  # is shallower indented
                break

    if row - last_row - 1 > 0:
        set_rows_visible(last_row + 1, row - last_row - 1, False)
    if 0 <= last_row:
        set_rows_visible(last_row, 1, True)

    # restore previous selection
    select(fcol, frow, lcol, lrow)

if __name__ == "__main__":
    group_selection()
