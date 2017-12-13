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
# A Recursive Function
#
# parameters:
#	col as long, \
#	ByRef row as long, _           'input & output
#	ByRef row_indent as long, _    'input & output
#	ByRef blank_row_cnt as long, _ 'ouput, no. of blank rows above row
#	end_row as long
def group(
    col,
    row,
    row_indent,
    end_row):

    global level

    last_row = row
    blank_row_cnt = 0

    while True:
        last_row = last_row + 1
        if last_row > end_row :
            row = last_row
            row_indent = -1
            return row, row_indent, blank_row_cnt

        last_row_indent = findNoIndentChar(xSheet.getCellByPosition(col, last_row).string)
        if 0 <= last_row_indent :
            break
        
        blank_row_cnt = blank_row_cnt + 1
    
    level = level + 1 #set before the 1st call to group()
    
    while row_indent < last_row_indent: # next item is deeper indented 
        #isBlankLine = False
        last_row, last_row_indent, blank_row_cnt = group(col, last_row, last_row_indent, end_row)
    

    # do selection and grouping 
    if row + 1 <= last_row - 1 - blank_row_cnt :
        oRange = xSheet.getCellRangeByPosition(0, row + 1, 0, last_row - 1 - blank_row_cnt)
        oController.select(oRange)

        #----------------------------------------------------------------------
        #get access to the document
        model = desktop.getCurrentComponent()
        document = model.getCurrentController()
        dispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")
        
        #----------------------------------------------------------------------
        # dim args1(0) as new com.sun.star.beans.PropertyValue
        # args1(0).Name = "RowOrCol"
        # args1(0).Value = "R"

        struct = get_struct()

        struct.Name = 'RowOrCol'
        struct.Value = 'R'

        if level < 8 : #Libre Office Calc only support max 7 levels of nested groups
            dispatcher.executeDispatch(document, ".uno:Group", "", 0, tuple([struct]))
        
        #if level < 3 :
        #	print level
        #
    
    
    level = level - 1
    
    row = last_row
    row_indent = last_row_indent

    return row, row_indent, blank_row_cnt

#
#Find no. of indentation cell in row
#
def findNoIndentCell(start_col, end_col, row):
    
    for c in range(start_col, end_col+1):
        s = xSheet.getCellByPosition(c, row).string
        if s == "" :
            #print "empty cell"
            pass
        else:
            #print s
            break
    else:
        return -1 #incdicate a blank row

    return  c - start_col

#
# A Recursive Function, Group Indentation Use Cell as Indent Unit, not char
#
# parameters:
#	col as long, \
#	ByRef row as long, _           'input & output
#	ByRef row_indent as long, _    'input & output
#	ByRef blank_row_cnt as long, _ 'ouput, no. of blank rows above row
#	end_row as long, \
#	end_col as long
def group_cell_indent(
    col,
    row,
    row_indent,
    end_col,
    end_row):

    global level

    last_row = row
    blank_row_cnt = 0

    while True:
        last_row = last_row + 1
        if last_row > end_row :
            row = last_row
            row_indent = -1
            return row, row_indent, blank_row_cnt
        

        last_row_indent = findNoIndentCell(col, end_col, last_row)
        if 0 <= last_row_indent :
            break
        
        blank_row_cnt = blank_row_cnt + 1

    
    level = level + 1 #set before the 1st call to group()
    
    while row_indent < last_row_indent: # next item is deeper indented
        #isBlankLine = False
        last_row, last_row_indent, blank_row_cnt = group_cell_indent(col, last_row, last_row_indent, end_col, end_row)


    # do selection and grouping 
    if row + 1 <= last_row - 1 - blank_row_cnt :
        oRange = xSheet.getCellRangeByPosition(0, row + 1, 0, last_row - 1 - blank_row_cnt)
        oController.select(oRange)
        
        #----------------------------------------------------------------------
        #get access to the document
        model = desktop.getCurrentComponent()
        document = model.getCurrentController()
        dispatcher = smgr.createInstance("com.sun.star.frame.DispatchHelper")

        #----------------------------------------------------------------------
        struct = get_struct()

        struct.Name = 'RowOrCol'
        struct.Value = 'R'
        
        
        if level < 8 : #Libre Office Calc only support max 7 levels of nested groups
            dispatcher.executeDispatch(document, ".uno:Group", "", 0, tuple([struct]))
        
        #if level < 3 :
        #	print level
        #
    
    
    level = level - 1
    
    row = last_row
    row_indent = last_row_indent
    return row, row_indent, blank_row_cnt


# TODO: merge group() and group_cell_indent() into 1 function
#
# selection area:     | indentation unit | range 
#---------------------+------------------+--------------------------------------------
# 1 row, 1 column     | character        | from selected row to whole sheet
# 1 row, >1 columns   | cell             | from selected top left cell to whole sheet
# >1 row, 1 columns   | character        | selected rows
# >1 row, >1 columns  | cell             | selected range 
def group_selection():
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
        
        end_col = c.RangeAddress.EndColumn	
        end_row = c.RangeAddress.EndRow
    else:
        end_col = addr.EndColumn
        end_row = addr.EndRow

    global level
    level = 0

    blank_row_cnt = 0 #any value will do, this ByRef parameter serve as output only
    
    if addr.StartColumn == addr.EndColumn : #single column is selected, char indentation mode
        row_indent = findNoIndentChar(xSheet.getCellByPosition(col, row).string) 
        
        while True:
            row, row_indent, blank_row_cnt = group(
                col,
                row,
                row_indent,
                end_row)
            # print "here"
            if row_indent < 0:
                break
    else:
        row_indent = findNoIndentCell(col, end_col, row) 

        while True:
            row, row_indent, blank_row_cnt = group_cell_indent(
                col,
                row,
                row_indent,
                end_col,
                end_row)
            if row_indent < 0:
                break

if __name__ == "__main__":
    group_selection()