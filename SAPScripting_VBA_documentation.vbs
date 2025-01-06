


' For initializing Codes. In this instance, it considers SAPGUI is already open and loged in (Less conveluted code):

Public session as Object
Private Sub initializeSAPConnection()
     Dim SapGuiAuto As Object: Set SapGuiAuto = GetObject("SAPGUI")
     Dim Applic As Object: Set Applic = SapGuiAuto.GetScriptingEngine
     Dim Connection As Object: Set Connection = Applic.Children(0)
     Set session = Connection.Children(0)
End Sub




' Then just reffer to session as below:
    session.findById("wnd[0]").sendVKey 0 'For example



' We use the Find by id to identify the elements for using:
.findById("wnd[0]") '-> Main Screen
.findById("wnd[1]") '-> The second window opened
.findById("wnd[0]/tbar[0]/okcd") '-> The element okcd
.findById("wnd[0]/tbar[1]/btn[45]") '-> The button with id 45
.findById("wnd[0]/usr/chkEINA-LOEKZ") '-> The checkbox LOEKZ on table EINA
.findById("wnd[0]/usr/ctxtEINA-MATNR") '-> The Text field MATNR on table EINA
.findById("wnd[2]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell") '-> The table (Shell) on this pathway
.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]") '-> The element [4,0] of the table
' Then we have common atributes

Examples of atributes:

.example '-> What does the code do
example.example 0 '-> In code use
' Other Commentaries



' Main Screens
.maximize '-> Maximaze the page screen size
session.findById("wnd[0]").maximize
.resizeWorkingPane '-> Resize the page
session.findById("wnd[0]").resizeWorkingPane 125,41,false
.sendVKey ' -> Input a Key Button
Session.findById("wnd[0]").sendVKey 0 ' For list, check #Key Mappings#
.doubleClick '-> Double Click
.close '-> Close...

' Text and value fields



.text '-> Get or Set the text of a field
session.findById("wnd[0]/tbar[0]/okcd").text = "ME23N"
.caretPosition '-> Change the cursosr placement on the text field
session.findById("wnd[0]/tbar[0]/okcd").caretPosition
.Value'-> Get or Set the value of a field
session.findById("wnd[0]/usr/ctxtEINA-MATNR").Value = "123456"
.setFocus '-> Focus on an item. [Normally ignored]
session.findById("wnd[0]/usr/ctxtEINA-MATNR").setfocus



' Buttons and Selection Boxes / Checkbox



.Selected '-> Get or Set a checkbox state
session.findById("wnd[0]/usr/chkEINA-LOEKZ").Selected = True
.press '-> Simulate a click
session.findById("wnd[0]/tbar[1]/btn[45]").press



' Lists



.Select '-> Select an item from a list
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select


' Tables


.setCurrentCell '-> Setting current cell, needs the row as a INT and the collumn name as String
session.findById("wnd[0]/.../shell").setCurrentCell 3, "TEXT"
.selectRows '-> Select a Row
.selectColumn '-> Select a Collumn
session.findById("wnd[0]/.../shell").selectRow  3
.currentCellRow '-> Set the current cell row
.currentCellColumn '-> Set the current cell Column
session.findById("wnd[0]/.../shell").currentCellRow
.verticalScrollbar.Position '-> Scroll vertically
session.findById("wnd[0]/.../shell").verticalScrollbar.Position 15
.firstVisibleRow '-> Scroll down to the row 
session.findById("wnd[0]/.../shell").firstVisibleRow 14
.selectAll '-> Select all itens
session.findById("wnd[0]/.../shell").selectAll 
.getAbsoluteRow '-> Select absolute row of a table
session.findById("wnd[0]/.../shell").getAbsoluteRow(0).Selected = True
.clickCurrentCell '-> Click current Cell
session.findById("wnd[0]/.../shell").clickCurrentCell
.pressColumnHeader '-> Press Column Header
session.findById("wnd[0]/.../shell").pressColumnHeader "BITM_DESCR"


' Attachments:



' Pressing toolbar context button:
session.findById("wnd[1]/usr/cntlCONT_111/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
' Selecting context menu item:
session.findById("wnd[1]/usr/cntlCONT_111/shellcont/shell").selectContextMenuItem "&LOAD"



' Clicking current cell:


    ' #Key Mappings#:
    00  -> Enter / 01~10 -> F1~F10
    11  -> Ctrl+S / 12  -> F12
    13~21 -> Shift+F1~Shift+F9
    22  -> Shift+Ctrl+0
    23~24  -> Shift+F11~Shift+F12
    25~36 -> Ctrl+F1~Ctrl+F12
    37~48 -> Ctrl+Shift+F1~Ctrl+Shift+F12
    70  -> Ctrl+E
    71  -> Ctrl+F
    72  -> Ctrl+/
    73  -> Ctrl+\
    74  -> Ctrl+N
    75  -> Ctrl+O
    76  -> Ctrl+X
    77  -> Ctrl+C
    78  -> Ctrl+V
    79  -> Ctrl+Z
    80/83  -> Ctrl+PageUp/ Ctrl+PageDown
    81/82  -> PageUp/PageDown
    84  -> Ctrl+G
    85  -> Ctrl+R
    86  -> Ctrl+P   