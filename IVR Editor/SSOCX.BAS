Attribute VB_Name = "Mdl_PrintConst"
'----------------------------------------------------------
'
' File: SSOCX.BAS
'
' Copyright (C) 2004 FarPoint Technologies.
' All rights reserved.
'
'----------------------------------------------------------
'  ********* MISCELLANEOUS DEFINES **********
Public Const SS_SHEET_MAX = 1000

' ********** SPREADSHEET PROPERTY SETTINGS **********

' Action property settings
Public Const SS_ACTION_ACTIVE_CELL = 0
Public Const SS_ACTION_GOTO_CELL = 1
Public Const SS_ACTION_SELECT_BLOCK = 2
Public Const SS_ACTION_CLEAR = 3
Public Const SS_ACTION_DELETE_COL = 4
Public Const SS_ACTION_DELETE_ROW = 5
Public Const SS_ACTION_INSERT_COL = 6
Public Const SS_ACTION_INSERT_ROW = 7
Public Const SS_ACTION_RECALC = 11
Public Const SS_ACTION_CLEAR_TEXT = 12
Public Const SS_ACTION_PRINT = 13
Public Const SS_ACTION_DESELECT_BLOCK = 14
Public Const SS_ACTION_DSAVE = 15
Public Const SS_ACTION_SET_CELL_BORDER = 16
Public Const SS_ACTION_ADD_MULTISELBLOCK = 17
Public Const SS_ACTION_GET_MULTI_SELECTION = 18
Public Const SS_ACTION_COPY_RANGE = 19
Public Const SS_ACTION_MOVE_RANGE = 20
Public Const SS_ACTION_SWAP_RANGE = 21
Public Const SS_ACTION_CLIPBOARD_COPY = 22
Public Const SS_ACTION_CLIPBOARD_CUT = 23
Public Const SS_ACTION_CLIPBOARD_PASTE = 24
Public Const SS_ACTION_SORT = 25
Public Const SS_ACTION_COMBO_CLEAR = 26
Public Const SS_ACTION_COMBO_REMOVE = 27
Public Const SS_ACTION_RESET = 28
Public Const SS_ACTION_SEL_MODE_CLEAR = 29
Public Const SS_ACTION_VMODE_REFRESH = 30
Public Const SS_ACTION_SMARTPRINT = 32

' ActiveCellHighlight property settings
Public Const SS_ACTIVECELLHIGHLIGHTSTYLE_OFF = 0
Public Const SS_ACTIVECELLHIGHLIGHTSTYLE_NORMAL = 1

' Appearance property settings
Public Const SS_APPEARANCE_FLAT = 0
Public Const SS_APPEARANCE_3D = 1
Public Const SS_APPEARANCE_3DWITHBORDER = 2

' BackColorStyle property settings
Public Const SS_BACKCOLORSTYLE_OVERGRID = 0
Public Const SS_BACKCOLORSTYLE_UNDERGRID = 1
Public Const SS_BACKCOLORSTYLE_OVERHORZGRIDONLY = 2
Public Const SS_BACKCOLORSTYLE_OVERVERTGRIDONLY = 3

' ButtonDrawMode property settings
Public Const SS_BDM_ALWAYS = 0
Public Const SS_BDM_CURRENT_CELL = 1
Public Const SS_BDM_CURRENT_COLUMN = 2
Public Const SS_BDM_CURRENT_ROW = 4
Public Const SS_BDM_ALWAYS_BUTTON = 8
Public Const SS_BDM_ALWAYS_COMBO = 16

' CellBorderStyle property settings
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_DASH = 2
Public Const SS_BORDER_STYLE_DOT = 3
Public Const SS_BORDER_STYLE_DASH_DOT = 4
Public Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Public Const SS_BORDER_STYLE_BLANK = 6
Public Const SS_BORDER_STYLE_FINE_SOLID = 11
Public Const SS_BORDER_STYLE_FINE_DASH = 12
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' CellBorderType property settings
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8
Public Const SS_BORDER_TYPE_OUTLINE = 16

' CellNoteIndicator property settings
Public Const SS_CELLNOTEINDICATOR_SHOWANDFIREEVENT = 0
Public Const SS_CELLNOTEINDICATOR_SHOWANDDONOTFIREEVENT = 1
Public Const SS_CELLNOTEINDICATOR_DONOTSHOWANDFIREEVENT = 2
Public Const SS_CELLNOTEINDICATOR_DONOTSHOWANDDONOTFIREEVENT = 3

' CellType property settings
Public Const SS_CELL_TYPE_DATE = 0
Public Const SS_CELL_TYPE_EDIT = 1
Public Const SS_CELL_TYPE_FLOAT = 2
Public Const SS_CELL_TYPE_INTEGER = 3
Public Const SS_CELL_TYPE_PIC = 4
Public Const SS_CELL_TYPE_STATIC_TEXT = 5
Public Const SS_CELL_TYPE_TIME = 6
Public Const SS_CELL_TYPE_BUTTON = 7
Public Const SS_CELL_TYPE_COMBOBOX = 8
Public Const SS_CELL_TYPE_PICTURE = 9
Public Const SS_CELL_TYPE_CHECKBOX = 10
Public Const SS_CELL_TYPE_OWNER_DRAWN = 11
Public Const SS_CELL_TYPE_CURRENCY = 12
Public Const SS_CELL_TYPE_NUMBER = 13
Public Const SS_CELL_TYPE_PERCENT = 14
Public Const SS_TYPE_SCIENTIFIC = 15


' ClipboardOptions property settings
Public Const SS_CLIP_NOHEADERS = 0
Public Const SS_CLIP_COPYROWHEADERS = 1
Public Const SS_CLIP_PASTEROWHEADERS = 2
Public Const SS_CLIP_COPYCOLHEADERS = 4
Public Const SS_CLIP_PASTECOLHEADERS = 8
Public Const SS_CLIP_COPYPASTEALLHEADERS = 15

' ColHeadersAutoText and RowHeadersAutoText property settings
Public Const SS_HEADER_BLANK = 0
Public Const SS_HEADER_NUMBERS = 1
Public Const SS_HEADER_LETTERS = 2

' ColMerge and RowMerge property settings
Public Const SS_MERGE_NONE = 0
Public Const SS_MERGE_ALWAYS = 1
Public Const SS_MERGE_RESTRICTED = 2

' ColUserSortIndicator property settings
Public Const SS_COLUSERSORTINDICATOR_NONE = 0
Public Const SS_COLUSERSORTINDICATOR_ASCENDING = 1
Public Const SS_COLUSERSORTINDICATOR_DESCENDING = 2
Public Const SS_COLUSERSORTINDICATOR_DISABLED = 3

' CursorStyle property settings
Public Const SS_CURSOR_STYLE_USER_DEFINED = 0
Public Const SS_CURSOR_STYLE_DEFAULT = 1
Public Const SS_CURSOR_STYLE_ARROW = 2
Public Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Public Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Public Const SS_CURSOR_TYPE_DEFAULT = 0
Public Const SS_CURSOR_TYPE_COLRESIZE = 1
Public Const SS_CURSOR_TYPE_ROWRESIZE = 2
Public Const SS_CURSOR_TYPE_BUTTON = 3
Public Const SS_CURSOR_TYPE_GRAYAREA = 4
Public Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Public Const SS_CURSOR_TYPE_COLHEADER = 6
Public Const SS_CURSOR_TYPE_ROWHEADER = 7
Public Const SS_CURSOR_TYPE_DRAGDROPAREA = 8
Public Const SS_CURSOR_TYPE_DRAGDROP = 9

' DAutoSizeCols property settings
Public Const SS_AUTOSIZE_NO = 0
Public Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Public Const SS_AUTOSIZE_BEST_GUESS = 2

' EditEnterAction property settings
Public Const SS_CELL_EDITMODE_EXIT_NONE = 0
Public Const SS_CELL_EDITMODE_EXIT_UP = 1
Public Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Public Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Public Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Public Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Public Const SS_CELL_EDITMODE_EXIT_SAME = 7
Public Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' OperationMode property settings
Public Const SS_OP_MODE_NORMAL = 0
Public Const SS_OP_MODE_READONLY = 1
Public Const SS_OP_MODE_ROWMODE = 2
Public Const SS_OP_MODE_SINGLE_SELECT = 3
Public Const SS_OP_MODE_MULTI_SELECT = 4
Public Const SS_OP_MODE_EXT_SELECT = 5

' Position property settings
Public Const SS_POSITION_UPPER_LEFT = 0
Public Const SS_POSITION_UPPER_CENTER = 1
Public Const SS_POSITION_UPPER_RIGHT = 2
Public Const SS_POSITION_CENTER_LEFT = 3
Public Const SS_POSITION_CENTER_CENTER = 4
Public Const SS_POSITION_CENTER_RIGHT = 5
Public Const SS_POSITION_BOTTOM_LEFT = 6
Public Const SS_POSITION_BOTTOM_CENTER = 7
Public Const SS_POSITION_BOTTOM_RIGHT = 8

' PrintOrientation property settings
Public Const SS_PRINTORIENT_DEFAULT = 0
Public Const SS_PRINTORIENT_PORTRAIT = 1
Public Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintPageOrder property settings
Public Const SS_PAGEORDER_AUTO = 0
Public Const SS_PAGEORDER_DOWNTHENOVER = 1
Public Const SS_PAGEORDER_OVERTHENDOWN = 2

' PrintType property settings
Public Const SS_PRINT_ALL = 0
Public Const SS_PRINT_CELL_RANGE = 1
Public Const SS_PRINT_CURRENT_PAGE = 2
Public Const SS_PRINT_PAGE_RANGE = 3

' ScrollBars property settings
Public Const SS_SCROLLBAR_NONE = 0
Public Const SS_SCROLLBAR_H_ONLY = 1
Public Const SS_SCROLLBAR_V_ONLY = 2
Public Const SS_SCROLLBAR_BOTH = 3

' ScrollBarTrack property settings
Public Const SS_SCROLLBARTRACK_OFF = 0
Public Const SS_SCROLLBARTRACK_VERTICAL = 1
Public Const SS_SCROLLBARTRACK_HORIZONTAL = 2
Public Const SS_SCROLLBARTRACK_BOTH = 3

' SelBackColor property settings
Public Const SPREAD_COLOR_NONE = &H8000000B

' SelectBlockOptions property settings
Public Const SS_SELBLOCKOPT_COLS = 1
Public Const SS_SELBLOCKOPT_ROWS = 2
Public Const SS_SELBLOCKOPT_BLOCKS = 4
Public Const SS_SELBLOCKOPT_ALL = 8

' ScrollTipTextMax
Public Const SS_SCROLLTIP_TEXTMAX = 1024

' ShowScrollTips property settings
Public Const SS_SHOWSCROLLTIPS_OFF = 0
Public Const SS_SHOWSCROLLTIPS_VERT = 1
Public Const SS_SHOWSCROLLTIPS_HORZ = 2
Public Const SS_SHOWSCROLLTIPS_BOTH = 3

' SortKeyOrder property settings
Public Const SS_SORT_ORDER_NONE = 0
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_SORT_ORDER_DESCENDING = 2

' TabStripPolicy property settings
Public Const SS_TABSTRIPPOLICY_ASNEEDED = 0
Public Const SS_TABSTRIPPOLICY_ALWAYS = 1
Public Const SS_TABSTRIPPOLICY_NEVER = 2

' TextTip property settings
Public Const SS_TEXTTIP_OFF = 0
Public Const SS_TEXTTIP_FIXED = 1
Public Const SS_TEXTTIP_FLOATING = 2
Public Const SS_TEXTTIP_FIXEDFOCUSONLY = 3
Public Const SS_TEXTTIP_FLOATINGFOCUSONLY = 4

' TypeButtonAlign property settings
Public Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Public Const SS_CELL_BUTTON_ALIGN_TOP = 1
Public Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Public Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' TypeButtonType property settings
Public Const SS_CELL_BUTTON_NORMAL = 0
Public Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeCheckTextAlign property settings
Public Const SS_CHECKBOX_TEXT_LEFT = 0
Public Const SS_CHECKBOX_TEXT_RIGHT = 1

' TypeCheckType property settings
Public Const SS_CHECKBOX_NORMAL = 0
Public Const SS_CHECKBOX_THREE_STATE = 1

' TypeComboBoxAutoSearch property settings
Public Const SS_COMBOBOX_AUTOSEARCH_NONE = 0
Public Const SS_COMBOBOX_AUTOSEARCH_SINGLECHAR = 1
Public Const SS_COMBOBOX_AUTOSEARCH_MULTIPLECHAR = 2
Public Const SS_COMBOBOX_AUTOSEARCH_SINGLECHARGREATER = 3

'TypeComboBoxWidth property settings
Public Const SS_COMBOWIDTH_CELLWIDTH = 0
Public Const SS_COMBOWIDTH_AUTORIGHT = 1
Public Const SS_COMBOWIDTH_AUTOLEFT = -1

' TypeCurrencyLeadingZero, TypeNumberLeadingZero,
' TypePercentLeadingZero property settings
Public Const SS_LEADINGZERO_INTL = 0
Public Const SS_LEADINGZERO_NO = 1
Public Const SS_LEADINGZERO_YES = 2

' TypeCurrencyNegStyle property settings
Public Const SS_CELL_CURRENCY_NEGSTYLE_INTL = 0
Public Const SS_CELL_CURRENCY_NEGSTYLE_1 = 1
Public Const SS_CELL_CURRENCY_NEGSTYLE_2 = 2
Public Const SS_CELL_CURRENCY_NEGSTYLE_3 = 3
Public Const SS_CELL_CURRENCY_NEGSTYLE_4 = 4
Public Const SS_CELL_CURRENCY_NEGSTYLE_5 = 5
Public Const SS_CELL_CURRENCY_NEGSTYLE_6 = 6
Public Const SS_CELL_CURRENCY_NEGSTYLE_7 = 7
Public Const SS_CELL_CURRENCY_NEGSTYLE_8 = 8
Public Const SS_CELL_CURRENCY_NEGSTYLE_9 = 9
Public Const SS_CELL_CURRENCY_NEGSTYLE_10 = 10
Public Const SS_CELL_CURRENCY_NEGSTYLE_11 = 11
Public Const SS_CELL_CURRENCY_NEGSTYLE_12 = 12
Public Const SS_CELL_CURRENCY_NEGSTYLE_13 = 13
Public Const SS_CELL_CURRENCY_NEGSTYLE_14 = 14
Public Const SS_CELL_CURRENCY_NEGSTYLE_15 = 15
Public Const SS_CELL_CURRENCY_NEGSTYLE_16 = 16

' TypeCurrencyPosStyle property settings
Public Const SS_CELL_CURRENCY_POSSTYLE_INTL = 0
Public Const SS_CELL_CURRENCY_POSSTYLE_1 = 1
Public Const SS_CELL_CURRENCY_POSSTYLE_2 = 2
Public Const SS_CELL_CURRENCY_POSSTYLE_3 = 3
Public Const SS_CELL_CURRENCY_POSSTYLE_4 = 4

' TypeDateFormat property settings
Public Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Public Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Public Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Public Const SS_CELL_DATE_FORMAT_YYMMDD = 3
Public Const SS_CELL_DATE_FORMAT_YYMM = 4
Public Const SS_CELL_DATE_FORMAT_MMDD = 5
Public Const SS_CELL_DATE_FORMAT_DEFAULT = 99

' TypeEditCharCase property settings
Public Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Public Const SS_CELL_EDIT_CASE_NO_CASE = 1
Public Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Public Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Public Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeHAlign property settings
Public Const SS_CELL_H_ALIGN_LEFT = 0
Public Const SS_CELL_H_ALIGN_RIGHT = 1
Public Const SS_CELL_H_ALIGN_CENTER = 2

' TypeNumberNegStyle property settings
Public Const SS_CELL_NUMBER_NEGSTYLE_INTL = 0
Public Const SS_CELL_NUMBER_NEGSTYLE_1 = 1
Public Const SS_CELL_NUMBER_NEGSTYLE_2 = 2
Public Const SS_CELL_NUMBER_NEGSTYLE_3 = 3
Public Const SS_CELL_NUMBER_NEGSTYLE_4 = 4
Public Const SS_CELL_NUMBER_NEGSTYLE_5 = 5

' TypePercentNegStyle property settings
Public Const SS_CELL_PERCENT_NEGSTYLE_INTL = 0
Public Const SS_CELL_PERCENT_NEGSTYLE_1 = 1
Public Const SS_CELL_PERCENT_NEGSTYLE_2 = 2
Public Const SS_CELL_PERCENT_NEGSTYLE_3 = 3
Public Const SS_CELL_PERCENT_NEGSTYLE_4 = 4
Public Const SS_CELL_PERCENT_NEGSTYLE_5 = 5
Public Const SS_CELL_PERCENT_NEGSTYLE_6 = 6
Public Const SS_CELL_PERCENT_NEGSTYLE_7 = 7
Public Const SS_CELL_PERCENT_NEGSTYLE_8 = 8

' TypeTextAlignVert property settings
Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Public Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Public Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTextOrient property settings
Public Const SS_CELL_TEXTORIENT_HORIZONTAL = 0
Public Const SS_CELL_TEXTORIENT_VERTICAL_LTR = 1
Public Const SS_CELL_TEXTORIENT_DOWN = 2
Public Const SS_CELL_TEXTORIENT_UP = 3
Public Const SS_CELL_TEXTORIENT_INVERT = 4
Public Const SS_CELL_TEXTORIENT_VERTICAL_RTL = 5

' TypeTime24Hour property settings
Public Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Public Const SS_CELL_TIME_24_HOUR_CLOCK = 1
Public Const SS_CELL_TIME_24_HOUR_DEFAULT = 2

' TypeVAlign property settings
Public Const SS_CELL_V_ALIGN_TOP = 0
Public Const SS_CELL_V_ALIGN_BOTTOM = 1
Public Const SS_CELL_V_ALIGN_VCENTER = 2

' UnitType property settings
Public Const SS_CELL_UNIT_NORMAL = 0
Public Const SS_CELL_UNIT_VGA = 1
Public Const SS_CELL_UNIT_TWIPS = 2

' UserColAction property settings
Public Const SS_USERCOLACTION_DEFAULT = 0
Public Const SS_USERCOLACTION_SORT = 1
Public Const SS_USERCOLACTION_SORTNOINDICATOR = 2
Public Const SS_USERCOLACTION_AUTOSIZE = 3

' UserResize property settings
Public Const SS_USER_RESIZE_NONE = 0
Public Const SS_USER_RESIZE_COL = 1
Public Const SS_USER_RESIZE_ROW = 2
Public Const SS_USER_RESIZE_BOTH = 3

' UserResizeCol and UserResizeRow property settings
Public Const SS_USER_RESIZE_DEFAULT = 0
Public Const SS_USER_RESIZE_ON = 1
Public Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Public Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4



' ********** SPREADSHEET METHOD SETTINGS ***********

' ActionKey method settings
Public Const SS_KBA_CLEAR = 0
Public Const SS_KBA_CURRENT = 1
Public Const SS_KBA_POPUP = 2

' AddCustomFunctionExt, GetCustomFunction method Flags parameter settings
Public Const SS_CUSTFUNC_WANTCELLREF = 1
Public Const SS_CUSTFUNC_WANTRANGEREF = 2

' CFGetParamInfo method Type parameter settings
Public Const SS_VALUE_TYPE_LONG = 0
Public Const SS_VALUE_TYPE_DOUBLE = 1
Public Const SS_VALUE_TYPE_STR = 2
Public Const SS_VALUE_TYPE_CELL = 3
Public Const SS_VALUE_TYPE_RANGE = 4

' CFGetParamInfo method Status parameter settings
Public Const SS_VALUE_STATUS_OK = 0
Public Const SS_VALUE_STATUS_ERROR = 1
Public Const SS_VALUE_STATUS_EMPTY = 2

' GetCellSpan method return values
Public Const SS_SPAN_NO = 0
Public Const SS_SPAN_YES = 1
Public Const SS_SPAN_ANCHOR = 2

' ExportExcelBookEx, ExportExcelToFileEx
Public Const SS_EXCELSAVEFLAG_NONE = 0
Public Const SS_EXCELSAVEFLAG_NOFORMULAS = 1

' ExportTextFile, ExportRangeToTextFile, ExportToXML and  LoadTextFile
Public Const SS_EXPORTTEXT_CREATE = 1
Public Const SS_EXPORTTEXT_APPEND = 2
Public Const SS_EXPORTTEXT_UNFORMATTED = 4
Public Const SS_EXPORTTEXT_COLHEADERS = 8
Public Const SS_EXPORTTEXT_ROWHEADERS = 16

Public Const SS_EXPORTXML_FORMATTED = 0
Public Const SS_EXPORTXML_UNFORMATTED = 1

Public Const SS_LOADTEXT_NOHEADERS = 0
Public Const SS_LOADTEXT_COLHEADERS = 1
Public Const SS_LOADTEXT_ROWHEADERS = 2
Public Const SS_LOADTEXT_CLEARDATAONLY = 4

' GetRefStyle/SetRefStyle methods return values/parameter settings
Public Const SS_REFSTYLE_DEFAULT = 0
Public Const SS_REFSTYLE_A1 = 1
Public Const SS_REFSTYLE_R1C1 = 2

' PrintScaling method values
Public Const SS_PRINTSCALINGMETHOD_NONE = 0
Public Const SS_PRINTSCALINGMETHOD_SMARTPRINT = 1
Public Const SS_PRINTSCALINGMETHOD_ZOOM = 2
Public Const SS_PRINTSCALINGMETHOD_BESTFIT = 3

' PrintSheet flags
Public Const SS_PRINTFLAGS_NONE = 0
Public Const SS_PRINTFLAGS_SHOWCOMMONDIALOG = 1

' SearchCol and SearchRow method's SearchFlags values
Public Const SS_SEARCHFLAGS_NONE = 0
Public Const SS_SEARCHFLAGS_GREATEROREQUAL = 1
Public Const SS_SEARCHFLAGS_PARTIALMATCH = 2
Public Const SS_SEARCHFLAGS_VALUE = 4
Public Const SS_SEARCHFLAGS_CASESENSITIVE = 8
Public Const SS_SEARCHFLAGS_SORTEDASCENDING = 16
Public Const SS_SEARCHFLAGS_SORTEDDESCENDING = 32

' Sort method's SortBy parameter settings
Public Const SS_SORT_BY_ROW = 0
Public Const SS_SORT_BY_COL = 1



' ********** SPREADSHEET EVENT SETTINGS **********
' BeforeEditMode event - UserAction parameter settings
Public Const SS_BEM_KEYBOARD = 0
Public Const SS_BEM_MOUSE = 1
Public Const SS_BEM_CODE = 2

' BeforeEditMode event - CursorPos parameter settings
Public Const SS_BEM_CURSORPOS_DEFAULT = 0
Public Const SS_BEM_CURSORPOS_SELECTALL = 1
Public Const SS_BEM_CURSORPOS_BEGINNING = 2
Public Const SS_BEM_CURSORPOS_END = 3
Public Const SS_BEM_CURSORPOS_MOUSE = 4
Public Const SS_BEM_CURSORPOS_DECIMAL = 5

' BeforeUserSort event - DefaultAction parameter settings
Public Const SS_BEFOREUSERSORT_DEFAULTACTION_CANCEL = 0
Public Const SS_BEFOREUSERSORT_DEFAULTACTION_AUTOSORT = 1
Public Const SS_BEFOREUSERSORT_DEFAULTACTION_MANUALSORT = 2

' BeforeUserSort event - DefaultAction parameter settings
Public Const SS_BEFOREUSERSORT_STATE_NONE = 0
Public Const SS_BEFOREUSERSORT_STATE_ASCENDING = 1
Public Const SS_BEFOREUSERSORT_STATE_DESCENDING = 2

' TextTipFetch event MultiLine parameter settings
Public Const SS_TT_MULTILINE_SINGLE = 0
Public Const SS_TT_MULTILINE_MULTI = 1
Public Const SS_TT_MULTILINE_AUTO = 2


' ********** PRINT PREVIEW PROPERTY SETTINGS **********

' GrayAreaMarginType property values
Public Const SPV_GRAYAREAMARGINTYPE_SCALED = 0
Public Const SPV_GRAYAREAMARGINTYPE_ACTUAL = 1

' MousePointer property values
Public Const SPV_MOUSEPOINTER_DEFAULT = 0
Public Const SPV_MOUSEPOINTER_ARROW = 1
Public Const SPV_MOUSEPOINTER_CROSS = 2
Public Const SPV_MOUSEPOINTER_I_BEAM = 3
Public Const SPV_MOUSEPOINTER_ICON = 4
Public Const SPV_MOUSEPOINTER_SIZE = 5
Public Const SPV_MOUSEPOINTER_SIZE_NE_SW = 6
Public Const SPV_MOUSEPOINTER_SIZE_N_S = 7
Public Const SPV_MOUSEPOINTER_SIZE_NW_SE = 8
Public Const SPV_MOUSEPOINTER_SIZE_W_E = 9
Public Const SPV_MOUSEPOINTER_UP_ARROW = 10
Public Const SPV_MOUSEPOINTER_HOURGLASS = 11
Public Const SPV_MOUSEPOINTER_NO_DROP = 12

' PageViewType property values
Public Const SPV_PAGEVIEWTYPE_WHOLE_PAGE = 0
Public Const SPV_PAGEVIEWTYPE_NORMAL_SIZE = 1
Public Const SPV_PAGEVIEWTYPE_PERCENTAGE = 2
Public Const SPV_PAGEVIEWTYPE_PAGE_WIDTH = 3
Public Const SPV_PAGEVIEWTYPE_PAGE_HEIGHT = 4
Public Const SPV_PAGEVIEWTYPE_MULTIPLE_PAGES = 5

' ScrollBarH property values
Public Const SPV_SCROLLBARH_SHOW = 0
Public Const SPV_SCROLLBARH_AUTO = 1
Public Const SPV_SCROLLBARH_HIDE = 2

' ScrollBarV property values
Public Const SPV_SCROLLBARV_SHOW = 0
Public Const SPV_SCROLLBARV_AUTO = 1
Public Const SPV_SCROLLBARV_HIDE = 2

' ZoomState property values
Public Const SPV_ZOOMSTATE_INDETERMINATE = 0
Public Const SPV_ZOOMSTATE_IN = 1
Public Const SPV_ZOOMSTATE_OUT = 2
Public Const SPV_ZOOMSTATE_SWITCH = 3
