Attribute VB_Name = "Opening"
'Credits to Alexandru Ionescu for the idea and some code :)

Option Explicit

Public Const FILEMENUID         As Integer = 1
Public Const ABOUTMENUID        As Integer = 2
Public Const FILETOGGLEID       As Integer = 100
Public Const FILESEPARATOR1ID   As Integer = 101
Public Const FILEEXITID         As Integer = 102

Public hForm                    As Long
Public hCmd                     As Long
Public hText                    As Long
Public hLabel1                  As Long
Public hLabel2                  As Long
Public hRadBut1                 As Long
Public hRadBut2                 As Long
Public hRadBut3                 As Long
Public hChk1                    As Long
Public hMenu                    As Long
Public hFileSubMenu             As Long
Public hBrush                   As Long
Public blnButtonEnabled         As Boolean
Public intRadButValue           As Integer

Public Sub Main()

Dim X As Long, Y As Long, mnuItem As MENUITEMINFO

    intRadButValue = 1
    
    'We must create the form first
    X = ((Screen.width / Screen.TwipsPerPixelX) - 350) / 2
    Y = ((Screen.height / Screen.TwipsPerPixelY) - 170) / 2
    hForm = CreateForm("All API form", FixedMinMax, X, Y, 350, 180)
    
    'Then the controls
    hCmd = CreateCmdButton(hForm, "Show text", 30, 30, 70, 20)
    EnableWindow hCmd, False 'Disable the button
    blnButtonEnabled = False
    hText = CreateTextbox(hForm, 30, 90, 100, 25)
    hLabel1 = CreateLabel(hForm, "Enter some text and press the button", 30, 55, 150, 25)
    hLabel2 = CreateLabel(hForm, "Textbox background color", 200, 5, 140, 20)
    hRadBut1 = CreateRadioButton(hForm, "Default color", 200, 25, 100, 25, True)
    hRadBut2 = CreateRadioButton(hForm, "Yellow", 200, 47, 100, 25, , False)
    hRadBut3 = CreateRadioButton(hForm, "Green", 200, 69, 100, 25, , False)
    SendMessage hRadBut1, BM_SETCHECK, BST_CHECKED, 0& 'Check the first one
    hChk1 = CreateCheckbox(hForm, "Toggle textbox", 200, 101, 100, 25)
    SendMessage hChk1, BM_SETCHECK, BST_CHECKED, 0& 'Textbox is enabled, check the checkbox
    
    'Create the File Sub Menu
    hFileSubMenu = CreatePopupMenu()
    mnuItem = CreateMenuItem("Toggle textbox", FILETOGGLEID, 0&, , , True)
    InsertMenuItem hFileSubMenu, FILETOGGLEID, False, mnuItem
    mnuItem = CreateMenuItem("", FILESEPARATOR1ID, 0&, True)
    InsertMenuItem hFileSubMenu, FILESEPARATOR1ID, False, mnuItem
    mnuItem = CreateMenuItem("Exit", FILEEXITID, 0&)
    InsertMenuItem hFileSubMenu, FILEEXITID, False, mnuItem
    'Create the main menu
    hMenu = CreateMenu()
    mnuItem = CreateMenuItem("File", FILEMENUID, hFileSubMenu)
    InsertMenuItem hMenu, FILEMENUID, False, mnuItem
    mnuItem = CreateMenuItem("About", ABOUTMENUID, 0)
    InsertMenuItem hMenu, ABOUTMENUID, False, mnuItem
    'Attach the menu to the form, that way we will not have to destroy the menu
    'before quitting, it will be destroyed with the form
    SetMenu hForm, hMenu
    
    'Watch for events
    GetAndDispatch hForm

End Sub
