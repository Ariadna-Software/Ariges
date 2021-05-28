Attribute VB_Name = "modCodejock17"
Public frmInbox As frmInbox
Public frmShortBar As frmShortcutBar2
Public frmEditEvent As frmEditEvent
Public frmPaneContacts As frmPaneContacts2
Public frmPaneCalendar As frmPaneCalendar2
Public pageBackstageHelp As pageBackstageHelp

'
'Public Const ID_SWITCH_PRINTLAYOUT = 7700
'Public Const ID_SWITCH_FULLSCREENREADING = 7701
'Public Const ID_SWITCH_WEBLAYOUT = 7702
'Public Const ID_SWITCH_OUTLINE = 7703
'Public Const ID_SWITCH_DRAFT = 7704
'
'Public Const ID_SWITCH_NORMAL = 7705
'Public Const ID_SWITCH_CALENAR_AND_TASK = 7706
'Public Const ID_SWITCH_CALENDAR = 7707
'Public Const ID_SWITCH_CLASSIC = 7708
'Public Const ID_SWITCH_READING = 7709
'
'Public Const ID_GALLERY_QUICKSTEP = 7750
'Public Const ID_QUICKSTEP_MOVE_TO = 7755
'Public Const ID_QUICKSTEP_TEAM_EMAIL = 7756
'Public Const ID_QUICKSTEP_REPLAY_DELETE = 7757
'Public Const ID_QUICKSTEP_TO_MANAGER = 7758
'Public Const ID_QUICKSTEP_DONE = 7759
'Public Const ID_QUICKSTEP_CREATE_NEW = 7760
'Public Const ID_GROUP_QUICKSTEP = 7761

'Public Const ID_QUICKSTEP_CATEGORIZE = 7762
'Public Const ID_QUICKSTEP_FLAG_MOVE = 7763

'Public Const ID_GROUP_MAIL_NEW = 7764
'Public Const ID_GROUP_MAIL_NEW_NEW = 7765
'Public Const ID_GROUP_MAIL_NEW_NEW_ITEMS = 7766
'Public Const ID_GROUP_MAIL_NEW_APPLOINTMENT = 7767
'Public Const ID_GROUP_MAIL_NEW_CONTACT = 7768
'Public Const ID_GROUP_MAIL_NEW_TASK = 7769
'Public Const ID_GROUP_MAIL_DELETE = 7770
'Public Const ID_GROUP_MAIL_DELETE_CLEANUP = 7771
'Public Const ID_GROUP_MAIL_DELETE_JUNK = 7772
'Public Const ID_GROUP_MAIL_DELETE_DELETE = 7773
'Public Const ID_GROUP_MAIL_RESPOND = 7774
'Public Const ID_GROUP_MAIL_RESPOND_REPLY = 7775
'Public Const ID_GROUP_MAIL_RESPOND_REPLY_ALL = 7776
'Public Const ID_GROUP_MAIL_RESPOND_FORWARD = 7777
'Public Const ID_GROUP_MAIL_RESPOND_MEETING = 7778
'Public Const ID_GROUP_MAIL_RESPOND_IM = 7779
'Public Const ID_GROUP_MAIL_RESPOND_MORE = 7780
'Public Const ID_GROUP_MAIL_MOVE = 7781
'Public Const ID_GROUP_MAIL_MOVE_MOVE = 7782
'Public Const ID_GROUP_MAIL_MOVE_ONENOTE = 7783
'Public Const ID_GROUP_MAIL_TAGS = 7784
'Public Const ID_GROUP_MAIL_TAGS_UNREAD = 7785
'Public Const ID_GROUP_MAIL_TAGS_CATEGORIZE = 7786
'Public Const ID_GROUP_MAIL_TAGS_FOLLOWUP = 7787
'Public Const ID_GROUP_MAIL_FIND = 7788
'Public Const ID_GROUP_MAIL_FIND_CONTACT = 7789
'Public Const ID_GROUP_MAIL_FIND_ADDRESSBOOK = 7790
'Public Const ID_GROUP_MAIL_FIND_FILTER = 7791

Public Const ID_INDICATOR_PAGENUMBER = 220
Public Const ID_INDICATOR_WORDCOUNT = 221
Public Const ID_INDICATOR_LANGUAGE = 222
Public Const ID_INDICATOR_TRACKCHANGES = 223
Public Const ID_INDICATOR_CAPSLOCK = 224
Public Const ID_INDICATOR_OVERTYPE = 225
Public Const ID_INDICATOR_MACRORECORDING = 226
Public Const ID_INDICATOR_VIEWSHORTCUTS = 227
Public Const ID_INDICATOR_ZOOM = 228
Public Const ID_INDICATOR_ZOOMSLIDER = 229

Public Const ID_OPTIONS_RTL = 3004
Public Const ID_OPTIONS_ANIMATION = 3005

Public Const ID_OPTIONS_STYLEBLUE2007 = 2890
Public Const ID_OPTIONS_STYLESILVER2007 = 2891
Public Const ID_OPTIONS_STYLEBLACK2007 = 2892
Public Const ID_OPTIONS_STYLEAQUA2007 = 2893
Public Const ID_OPTIONS_STYLESCENIC7 = 2894
Public Const ID_OPTIONS_STYLEBLUE2010 = 2896
Public Const ID_OPTIONS_STYLESILVER2010 = 2897
Public Const ID_OPTIONS_STYLEBLACK2010 = 2898
Public Const ID_OPTIONS_STYLESYSTEM = 2899

Public Const ID_RIBBON_MINIMIZE = 4567
Public Const ID_RIBBON_EXPAND = 4568
Public Const ID_RIBBON_QUICKACCESSEMPTYICON = 302

'Public Const IDR_CNTR_INPLACE = 6
'Public Const IDD_ABOUTBOX = 100
'Public Const IDP_OLE_INIT_FAILED = 100
'Public Const IDP_FAILED_TO_CREATE = 102
'Public Const IDR_MAINFRAME = 128
'Public Const IDR_SMALLICONS = 128
'Public Const IDR_RIBBONTYPE = 129
'Public Const IDR_LARGEICONS = 131
'Public Const IDR_LAYOUTTABSMALL = 143
'Public Const IDR_LAYOUTTABLARGE = 145
'Public Const IDR_MENU_CONTEXT = 147

'Public Const ID_APP_THEME = 148
'Public Const IDB_BITMAP_PICTURE = 149
'Public Const IDB_BITMAP_GRAPHIC = 150
'Public Const IDB_BITMAP_CHART = 151
'Public Const IDB_BITMAP_TABLE = 152
'Public Const IDB_INSERTTAB = 200
'Public Const IDB_WRITETAB = 201
'Public Const IDB_BITMAPS_GROUPS = 202
'Public Const IDB_GEAR = 300
'Public Const ID_GROUP_BUTTONPOPUP = 2000
'Public Const ID_INSERT_HYPERLINK = 2710
'Public Const ID_INSERT_CROSS_REFERENCE = 2712
'Public Const ID_TEXT_SIGNATURE = 2713
'Public Const ID_TEXT_DATETIME = 2714
'Public Const ID_TEXT_INSERTOBJECT = 2715
'Public Const ID_CANCEL_EDIT_CNTR = 2768
'Public Const ID_PAGES_NEW = 2772
'Public Const ID_PAGES_COVRE = 2773
'Public Const ID_PAGES_BREAK = 2774
'Public Const ID_TABLE_NEW = 2775


'Public Const ID_ILLUSTRATION_PICTURE = 2776
'Public Const ID_ILLUSTRATION_GRAPHIC = 2777
'Public Const ID_ILLUSTRATION_CHART = 2778
'Public Const ID_TABLE_INSERTTABLE = 2779
'Public Const ID_INSERT_HEADER = 2780
'Public Const ID_INSERT_FOOTER = 2781
'Public Const ID_INSERT_PAGENUMBER = 2782
'Public Const ID_TEXT_TEXTBOX = 2783
'Public Const ID_TEXT_PARTS = 2784
'Public Const ID_TEXT_WORDART = 2785
'Public Const ID_TEXT_DROPCAP = 2786
'Public Const ID_SYMBOL_EQUATIONS = 2787
'Public Const ID_SYMBOL_SYMBOL = 2788
'Public Const ID_PAGES_COVER = 2789

'Public Const ID_ILLUSTRATION_CLIPART = 2790
'Public Const ID_ILLUSTRATION_FROMCAMERA = 2791
'Public Const ID_INSERT_BOOKMARK = 2791
'Public Const ID_PAGENUMBER_FORMATPAGENUMBERS = 2792
'Public Const ID_FONT_GROW = 2792
'Public Const ID_PAGENUMBER_REMOVEPAGENUMBERS = 2793
'Public Const ID_FONT_SHRINK = 2793
'Public Const ID_NEWPAGE_BLANKPAGE = 2794
'Public Const ID_FONT_CLEAR = 2794
'Public Const ID_NEWPAGE_SELECTION = 2795
'Public Const ID_TEXT_CHANGECASE = 2795
'Public Const ID_INSERT_NUMBERING = 2796
'Public Const ID_INSERT_LIST = 2797
'Public Const ID_PARA_DECREASEINDENT = 2798
'Public Const ID_PARA_INCREASEINDENT = 2799
'Public Const ID_PARA_SORT = 2800
'Public Const ID_PARA_JUSTIFY = 2801
'Public Const ID_PARA_SHOWMARKS = 2802
'Public Const ID_DOCUMENTPARTS_AUTOTEXT = 2803
'Public Const ID_PARA_LINESPACING = 2803
'Public Const ID_DOCUMENTPARTS_PROPERTY = 2804
'Public Const ID_PARA_SHADING = 2804
'Public Const ID_DOCUMENTPARTS_FIELD = 2805
'Public Const ID_BORDERS_NOBORDER = 2805
'Public Const ID_DOCUMENTPARTS_BUILDINGBLOCKORGANIZER = 3806
'Public Const ID_TEXT_HIGHLIGHTCOLOR = 2806
'Public Const ID_VIEW_RULER = 2807
'Public Const ID_VIEW_GRIDLINES = 2808
'Public Const ID_VIEW_PROPERTIES = 2809
'Public Const ID_VIEW_DOCUMENTMAP = 2810
'Public Const ID_VIEW_THUMBNAILS = 2811
'Public Const ID_VIEW_ACTINBAR = 2812
'Public Const ID_TEXT_COLOR_SELECTOR = 2813
'Public Const ID_GROUP_PARAGRAPH = 5000
'Public Const ID_GROUP_CLIPBOARD = 5001
'Public Const ID_GROUP_FONT = 5002
'Public Const ID_GROUP_FIND = 5003

'Public Const ID_TAB_WRITE = 5004
'Public Const ID_TAB_INSERT = 5005
'Public Const ID_TAB_PAGELAYOUT = 5006
'Public Const ID_TAB_ADDINS = 5007
'Public Const ID_TAB_TABLEDESIGN = 5008
'Public Const ID_TAB_TABLELAYOUT = 5009
'Public Const ID_TAB_CHARTDESIGN = 5010
'Public Const ID_TAB_CHARTFORMAT = 5011
'Public Const ID_TAB_CHARTLAYOUT = 5012
'Public Const ID_TAB_CONTEXTCHART = 5013
'Public Const ID_GROUP_PAGES = 5014
'Public Const ID_GROUP_TABLE = 5015
'Public Const ID_GROUP_ILLUSTRATIONS = 5016
'Public Const ID_GROUP_HEADERFOOTERS = 5017
'Public Const ID_GROUP_LINKS = 5018
'Public Const ID_GROUP_TEXT = 5019
'Public Const ID_GROUP_SYMBOLS = 5020
'Public Const ID_GROUP_THEMES = 5021
'Public Const ID_GROUP_PAGESETUP = 5022
'Public Const ID_GROUP_PAGEBACKGROUND = 5023
'Public Const ID_GROUP_ARRANGE = 5024
'Public Const ID_GROUP_SHOWHIDE = 5025

'Public Const ID_TAB_VIEW = 5026
'Public Const ID_TAB_REFERENCES = 5027
'Public Const ID_TAB_MAILINGS = 5028
'Public Const ID_TAB_REVIEW = 5029
'Public Const ID_CHAR_BOLD = 7608
'Public Const ID_CHAR_ITALIC = 7610
'Public Const ID_CHAR_UNDERLINE = 7611
'Public Const ID_EDIT_GOTO = 7612
'Public Const ID_EDIT_SELECT_OBJECTS = 7613
'Public Const ID_EDIT_SELECT = 7614
'Public Const ID_EDIT_SELECT_MULTIPLE_OBJECTS = 7615
'Public Const ID_FORMAT_PAINTER = 7616
'Public Const ID_TEXT_FONT = 7617
'Public Const ID_FONT_FACE = 7618
'Public Const ID_FONT_SIZE = 7619

'Public Const ID_CHAR_STRIKETHROUGH = 7620
'Public Const ID_TEXT_SUBSCRIPT = 7621
'Public Const ID_TEXT_SUPERSCRIPT = 7622
'Public Const ID_TEXT_COLOR = 7623
'Public Const ID_INSERT_BULLET = 32777
'Public Const ID_COVERPAGE_REMOVECURRENTCOVERPAGE = 32796
'Public Const ID_COVERPAGE_SAVESELECTIONASNEWCOVERPAGE = 32797
'Public Const ID_TABLE_DRAWTABLE = 32799
'Public Const ID_TABLE_CONVERTTEXTTOTABLE = 32800
'Public Const ID_TEXTBOX_DRAWTEXTBOX = 32801
'Public Const ID_TEXTBOX_SAVESELECTIONASNEWTEXTBOX = 32802
'Public Const ID_PARA_LEFT = 32803
'Public Const ID_PARA_CENTER = 32804
'Public Const ID_PARA_RIGHT = 32805
'Public Const ID_EQUATIONS_MATH = 32807
'Public Const ID_THEMES_COLORS = 32808
'Public Const ID_THEMES_FONTS = 32809
'Public Const ID_THEMES_EFFECTS = 32810
'Public Const ID_PAGE_ORIENTATIONS = 32811
'Public Const ID_PAGE_ORIENTATION = 32812
'Public Const ID_PAGE_SIZE = 32813
'Public Const ID_PAGE_COLUMNS = 32814
'Public Const ID_PAGE_BREAKS = 32815
'Public Const ID_PAGE_LINENUMBERS = 32816
'Public Const ID_PAGE_HYPHENATATION = 32817
'Public Const ID_PAGE_WATERMARK = 32818
'Public Const ID_PAGE_COLOR = 32819
'Public Const ID_PAGE_BORDERS = 32820

'Public Const ID_ARRANGE_FRONT = 32821
'Public Const ID_ARRANGE_BACK = 32822
'Public Const ID_ARRANGE_ALIGN = 32823
'Public Const ID_ARRANGE_GROUP = 32824
'Public Const ID_ARRANGE_UNGROUP = 32825
'Public Const ID_ARRANGE_ROTATE = 32826
'Public Const ID_THEMES_THEMES = 32827
'Public Const ID_PAGE_MARGINS = 32828
'Public Const ID_ARRANGE_POSITION = 32829
'Public Const ID_ARRANGE_TEXTWRAPPING = 32831
'Public Const ID_CONTEXT_FONT = 32832
'Public Const ID_PARA_PARAGRAPH = 32833
'Public Const ID_THEME_OFFICE2003 = 32834
'Public Const ID_THEME_OFFICE2007 = 32835
'Public Const IDB_CLIENT_FACE = 3010
Public Const ID_SYSTEM_ICON = 1200

Public Const ID_FILE_PREPARE = 1230
Public Const ID_FILE_SEND_MAIL = 1231
Public Const ID_FILE_PUBLISH = 1232
Public Const ID_FILE_CLOSE = 1233
Public Const ID_FILE_SEND_INTERNETFAX = 1234
Public Const ID_FILE_SEND = 1235
Public Const ID_FILE_OPTIONS = 1236

Public Const ID_OPTIONS_FONT_SYSTEM = 42883
Public Const ID_OPTIONS_FONT_NORMAL = 42884
Public Const ID_OPTIONS_FONT_LARGE = 42885
Public Const ID_OPTIONS_FONT_EXTRALARGE = 42886
Public Const ID_OPTIONS_FONT_AUTORESIZEICONS = 42887


Public Const ID_GROUP_HEADERANDFOOTER = 2003
Public Const ID_GROUP_POPUPICON = 2004
Public Const ID_GROUP_STYLES = 2005
Public Const ID_GALLERY_STYLES = 2006
Public Const ID_GALLERY_SHAPES = 2007
Public Const ID_GALLERY_COLORS = 2010
Public Const ID_GALLERY_LARGE_COLORS_POPUP = 2012
Public Const ID_GROUP_SHAPES = 2205

Public Const ID_APP_ABOUT = 4000
Public Const ID_EDIT_PASTE = 4001
Public Const ID_EDIT_PASTE_SPECIAL = 4002
Public Const ID_EDIT_COPY = 4003
Public Const ID_EDIT_CUT = 4004
Public Const ID_EDIT_FIND = 57636
Public Const ID_EDIT_REPLACE = 4006
Public Const ID_EDIT_SELECT_ALL = 4007
Public Const ID_FILE_NEW = 4008
Public Const ID_FILE_OPEN = 4009
Public Const ID_FILE_SAVE = 4010
Public Const ID_FILE_PRINT = 4011
Public Const ID_FILE_SAVE_AS = 57604
Public Const ID_FILE_PRINT_PREVIEW = 57609
Public Const ID_FILE_PRINT_SETUP = 57606
Public Const ID_FILE_MRU_FILE1 = 57616
Public Const ID_APP_EXIT = 57665

Public Const ID_SEARCH_ICON = 57783

Public Const ID_SAMPLE_MENU_ITEM = 60006

Public Const ID_GROUP_CLIPBOARD_OPTION = 3400
Public Const ID_GROUP_FONT_OPTION = 3401


Public Const ID_OPTIONS_STYLEBLUE = 3000
Public Const ID_OPTIONS_STYLEBLACK = 3001
Public Const ID_OPTIONS_STYLEAQUA = 3002
Public Const ID_OPTIONS_STYLESILVER = 3003

Public Const ID_PARAGRAPH_INDENTLEFT = 4500
Public Const ID_PARAGRAPH_INDENTRIGHT = 4501
Public Const ID_PARAGRAPH_SPACINGBEFORE = 4502
Public Const ID_PARAGRAPH_SPACINGAFTER = 4503


'Commandbars public constants






Public Const ID_FILE_EXIT2 = 10004
Public Const ID_VIEW_STATUSBAR = 10016
Public Const IDS_ARRANGE_BY = 220

Public Const ID_THEME_CLIENTPANE = 190
Public Const ID_THEME_OFFICE2000_PLAIN = 191
Public Const ID_THEME_OFFICEXP_PLAIN = 192
Public Const ID_THEME_OFFICE2003_PLAIN = 193
Public Const ID_THEME_NATIVE_PLAIN = 194





Public Const ID_CALENDAREVENT_OPEN = 6050
Public Const ID_CALENDAREVENT_DELETE = 6051
Public Const ID_CALENDAREVENT_NEW = 6052
Public Const ID_CALENDAREVENT_CHANGE_TIMEZONE = 6053
Public Const ID_CALENDAREVENT_60 = 6054
Public Const ID_CALENDAREVENT_30 = 6055
Public Const ID_CALENDAREVENT_15 = 6056
Public Const ID_CALENDAREVENT_10 = 6057
Public Const ID_CALENDAREVENT_5 = 6058

Public Const ID_INDICATOR_CAPS = 59137
Public Const ID_INDICATOR_NUM = 59138
Public Const ID_INDICATOR_SCRL = 59139

Public Const FCONTROL = 8

'Report Control public constants

'public constants used to identify columns, this will be the column ItemIndex
Public Const COLUMN_IMPORTANCE = 0
Public Const COLUMN_ICON = 1
Public Const COLUMN_ATTACHMENT = 2
Public Const COLUMN_FROM = 3
Public Const COLUMN_SUBJECT = 4
Public Const COLUMN_SENT = 5
Public Const COLUMN_SIZE = 6
Public Const COLUMN_CHECK = 7
Public Const COLUMN_PRICE = 8
Public Const COLUMN_CREATED = 9
Public Const COLUMN_RECEIVED = 10
Public Const COLUMN_CONVERSATION = 11
Public Const COLUMN_CONTACTS = 12
Public Const COLUMN_MESSAGE = 13
Public Const COLUMN_CC = 14
Public Const COLUMN_CATEGORIES = 15
Public Const COLUMN_AUTOFORWARD = 16
Public Const COLUMN_DO_NOT_AUTOARCH = 17
Public Const COLUMN_DUE_BY = 18
  
'public constants used to identify icons used in the ReportControl
Public Const COLUMN_MAIL_ICON = 1
Public Const COLUMN_IMPORTANCE_ICON = 2
Public Const COLUMN_CHECK_ICON = 3
Public Const RECORD_UNREAD_MAIL_ICON = 4
Public Const RECORD_READ_MAIL_ICON = 5
Public Const RECORD_REPLIED_ICON = 6
Public Const RECORD_IMPORTANCE_HIGH_ICON = 7
Public Const COLUMN_ATTACHMENT_ICON = 8
Public Const COLUMN_ATTACHMENT_NORMAL_ICON = 9
Public Const RECORD_IMPORTANCE_LOW_ICON = 10

Public Const IMPORTANCE_HIGH = 0
Public Const IMPORTANCE_NORMAL = 1
Public Const IMPORTANCE_LOW = 2

Public Const CHECKED_TRUE = 1
Public Const CHECKED_FALSE = 0

Public Const READ_TRUE = 1
Public Const READ_FALSE = 0

Public Const ATTACHMENTS_TRUE = 1
Public Const ATTACHMENTS_FALSE = 0

'Docking Pane Constants
Public Const PANE_SHORTCUTBAR = 1
Public Const PANE_REPORT_CONTROL = 2
Public Const PANE_READING_PANE = 3
Public Const PANE_FINDBAR = 4
Public ShowEventInPane As Boolean

'Shortcutbar constants
Public Const SHORTCUT_INBOX = 4300
Public Const SHORTCUT_CALENDAR = 4301
Public Const SHORTCUT_CONTACTS = 4302
Public Const SHORTCUT_TASKS = 4303
Public Const SHORTCUT_NOTES = 4304
Public Const SHORTCUT_FOLDER_LIST = 4305
Public Const SHORTCUT_SHORTCUTS = 4306
Public Const SHORTCUT_JOURNAL = 4307

Public Const SHORTCUT_SHOW_MORE = 4308
Public Const SHORTCUT_SHOW_FEWER = 4309

Public Const SHORTCUT_NAVIGATE_PANE_OPTIONS = 4310
Public Const SHORTCUT_ADD_REMOVE_BUTTONS = 4311

'Email Page constants
'Public Const ID_EMAIL_SEND = 4330
'Public Const ID_EMAIL_ADDRESS_BOOK = 4331
'Public Const ID_EMAIL_IMPORTANCE_HIGH = 4332
'Public Const ID_EMAIL_ATTACHMENT = 4333
'Public Const ID_EMAIL_IMPORTANCE_LOW = 4334
'Public Const ID_EMAIL_CHECK_NAMES = 4335
'Public Const ID_EMAIL_PERMISSION = 4336
'Public Const ID_EMAIL_FLAG = 337
'Public Const ID_EMAIL_CREATE_RULE = 338
'Public Const ID_EMAIL_OPTIONS = 339
'Public Const ID_EMAIL_CLOSE = 340
'Public Const ID_EMAIL_SAVE = 341
'Public Const ID_EMAIL_BOLD = 346
'Public Const ID_EMAIL_ITALIC = 347
'Public Const ID_EMAIL_UNDERSCORE = 348
'Public Const ID_EMAIL_BULLETS = 349
'Public Const ID_EMAIL_NUMBERING = 350
'Public Const ID_EMAIL_DECREASE_INDENT = 351
'Public Const ID_EMAIL_INCREASE_INDENT = 352
'Public Const ID_EMAIL_TRANSLATE = 354
'Public Const ID_EMAIL_LTR = 355
'Public Const ID_EMAIL_RTL = 356
'Public Const ID_EMAIL_LEFT = 357
'Public Const ID_EMAIL_CENTER = 358
'Public Const ID_EMAIL_RIGHT = 359
'
'Public Const ID_EMAIL_FILE_NEW = 370
'Public Const ID_EMAIL_FILE_OPEN = 371
'Public Const ID_EMAIL_FILE_SAVE_AS = 372
'Public Const ID_EMAIL_FILE_PRINT_SETUP = 373
'Public Const ID_EMAIL_FILE_PRINT_PREVIEW = 374
'Public Const ID_EMAIL_FILE_PRINT = 375
'Public Const ID_EMAIL_FILE_EXIT = 376

'Public Const ID_EMAIL_EDIT_UNDO = 400
'Public Const ID_EMAIL_EDIT_REDO = 401
'Public Const ID_EMAIL_EDIT_CUT = 402
'Public Const ID_EMAIL_EDIT_COPY = 403
'Public Const ID_EMAIL_EDIT_OFFICE_CLIPBOARD = 404
'Public Const ID_EMAIL_EDIT_PASTE = 405
'Public Const ID_EMAIL_EDIT_PASTE_SPECIAL = 406
'Public Const ID_EMAIL_EDIT_PASTE_AS_HYPERLINK = 407
'Public Const ID_EMAIL_EDIT_CLEAR = 408
'Public Const ID_EMAIL_EDIT_SELECT_ALL = 409
'Public Const ID_EMAIL_EDIT_FIND = 410
'Public Const ID_EMAIL_EDIT_REPLACE = 411
'Public Const ID_EMAIL_EDIT_GO_TO = 412
'Public Const ID_EMAIL_EDIT_UPDATE_IME_DICTIONARY = 413
'Public Const ID_EMAIL_EDIT_RECONVERT = 414
'Public Const ID_EMAIL_EDIT_LINKS = 415
'Public Const ID_EMAIL_EDIT_OBJECT = 416

'Public Const ID_EMAIL_VIEW_NORMAL = 417
'Public Const ID_EMAIL_VIEW_WEB_LAYOUT = 418
'Public Const ID_EMAIL_VIEW_PRINT_LAYOUT = 419
'Public Const ID_EMAIL_VIEW_READING_LAYOUT = 420
'Public Const ID_EMAIL_VIEW_OUTLINE = 421
'Public Const ID_EMAIL_VIEW_TASK_PANE = 422
'Public Const ID_EMAIL_VIEW_TOOLBARS = 423
'Public Const ID_EMAIL_VIEW_RULER = 424
'Public Const ID_EMAIL_VIEW_SHOW_PARAGRAPHMARKS = 425
'Public Const ID_EMAIL_VIEW_GRIDLINES = 426
'Public Const ID_EMAIL_VIEW_DOCUMENT_MAP = 427
'Public Const ID_EMAIL_VIEW_THUMBNAILS = 428
'Public Const ID_EMAIL_VIEW_HEADER_AND_FOOTER = 429
'Public Const ID_EMAIL_VIEW_FOOTNOTES = 430
'Public Const ID_EMAIL_VIEW_MARKUP = 431
'Public Const ID_EMAIL_VIEW_FULL_SCREEN = 432
'Public Const ID_EMAIL_VIEW_ZOOM = 433
'
'Public Const ID_EMAIL_INSERT_BREAK = 434
'Public Const ID_EMAIL_INSERT_PAGE_NUMBERS = 435
'Public Const ID_EMAIL_INSERT_DATE_AND_TIME = 436
'Public Const ID_EMAIL_INSERT_AUTOTEXT = 437
'Public Const ID_EMAIL_INSERT_FIELD = 438
'Public Const ID_EMAIL_INSERT_SYMBOL = 439
'Public Const ID_EMAIL_INSERT_COMMENT = 440
'Public Const ID_EMAIL_INSERT_NUMBER = 441
'Public Const ID_EMAIL_INSERT_REFERENCE = 442
'Public Const ID_EMAIL_INSERT_WEBCOMPONENT = 443
'Public Const ID_EMAIL_INSERT_PICTURE = 444
'Public Const ID_EMAIL_INSERT_DIAGRAM = 445
'Public Const ID_EMAIL_INSERT_TEXT_BOX = 446
'Public Const ID_EMAIL_INSERT_FILE = 447
'Public Const ID_EMAIL_INSERT_OBJECT = 448
'Public Const ID_EMAIL_INSERT_BOOKMARK = 449
'Public Const ID_EMAIL_INSERT_HYPERLINK = 450
'
'Public Const ID_EMAIL_FORMAT_FONT = 451
'Public Const ID_EMAIL_FORMAT_PARAGRAPH = 452
'Public Const ID_EMAIL_FORMAT_BULLETS_AND_NUMBERING = 453
'Public Const ID_EMAIL_FORMAT_BORDERS_AND_SHADING = 454
'Public Const ID_EMAIL_FORMAT_COLUMNS = 455
'Public Const ID_EMAIL_FORMAT_TABS = 456
'Public Const ID_EMAIL_FORMAT_DROP_CAP = 457
'Public Const ID_EMAIL_FORMAT_TEXT_DIRECTION = 458
'Public Const ID_EMAIL_FORMAT_CHANGE_CASE = 459
'Public Const ID_EMAIL_FORMAT_FIT_TEXT = 460
'Public Const ID_EMAIL_FORMAT_ASIAN_LAYOUT = 461
'Public Const ID_EMAIL_FORMAT_BACKGROUND = 462
'Public Const ID_EMAIL_FORMAT_THEME = 463
'Public Const ID_EMAIL_FORMAT_FRAMES = 464
'Public Const ID_EMAIL_FORMAT_AUTOFORMAT = 465
'Public Const ID_EMAIL_FORMAT_STYLES_AND_FORMATTING = 466
'Public Const ID_EMAIL_FORMAT_REVEAL_FORMATTING = 467
'Public Const ID_EMAIL_FORMAT_FORMAT_AUTOSHAPE_PICTURE = 468
'
'Public Const ID_EMAIL_TOOLS_SPELLING_AND_GRAMMAR = 469
'Public Const ID_EMAIL_TOOLS_RESEARCH = 470
'Public Const ID_EMAIL_TOOLS_LANGUAGE = 471
'Public Const ID_EMAIL_TOOLS_FIX_BROKEN_TEXT = 472
'Public Const ID_EMAIL_TOOLS_WORDCOUNT = 473
'Public Const ID_EMAIL_TOOLS_AUTOSUMMARIZE = 474
'Public Const ID_EMAIL_TOOLS_SPEECH = 475
'Public Const ID_EMAIL_TOOLS_SHAREDWORKSPACE = 476
'Public Const ID_EMAIL_TOOLS_TRACK_CHANGES = 477
'Public Const ID_EMAIL_TOOLS_COMPARE_AND_MERGE_DOCUMENTS = 478
'Public Const ID_EMAIL_TOOLS_PROTECT_DOCUMENT = 479
'Public Const ID_EMAIL_TOOLS_ONLINE_COLLABORATION = 480
'Public Const ID_EMAIL_TOOLS_LETTERS_AND_MAILINGS = 481
'Public Const ID_EMAIL_TOOLS_MACRO = 482
'Public Const ID_EMAIL_TOOLS_TEMPLATES_AND_ADDINS = 483
'Public Const ID_EMAIL_TOOLS_AUTOCORRECT_OPTIONS = 484
'Public Const ID_EMAIL_TOOLS_CUSTOMIZE = 485
'Public Const ID_EMAIL_TOOLS_OPTIONS = 486
'
'Public Const ID_EMAIL_TABLE_DRAW_TABLE = 487
'Public Const ID_EMAIL_TABLE_INSERT = 488
'Public Const ID_EMAIL_TABLE_DELETE = 489
'Public Const ID_EMAIL_TABLE_SELECT = 490
'Public Const ID_EMAIL_TABLE_MERGE_CELLS = 491
'Public Const ID_EMAIL_TABLE_SPLIT_CELLS = 492
'Public Const ID_EMAIL_TABLE_SPLITTABLE = 493
'Public Const ID_EMAIL_TABLE_TABLE_AUTOFORMAT = 494
'Public Const ID_EMAIL_TABLE_AUTOFIT = 495
'Public Const ID_EMAIL_TABLE_HEADING_ROWS_REPEAT = 496
'Public Const ID_EMAIL_TABLE_CONVERT = 497
'Public Const ID_EMAIL_TABLE_SORT = 498
'Public Const ID_EMAIL_TABLE_FORMULA = 499
'Public Const ID_EMAIL_TABLE_SHOW_GRIDLINES = 500
'Public Const ID_EMAIL_TABLE_TABLE_PROPERTIES = 501
'
'Public Const ID_EMAIL_WINDOW_ARRANGE_ALL = 502
'Public Const ID_EMAIL_WINDOW_COMPARE_SIDE_BY_SIDE_WITH = 503
'Public Const ID_EMAIL_WINDOW_SPLIT = 504
'
'Public Const ID_EMAIL_HELP_MICROSOFT_OFFICE_WORD_HELP = 505
'Public Const ID_EMAIL_HELP_SHOW_THE_OFFICE_ASSISTANT = 506
'Public Const ID_EMAIL_HELP_MICROSOFT_OFFICE_ONLINE = 507
'Public Const ID_EMAIL_HELP_CONTACT_US = 508
'Public Const ID_EMAIL_HELP_WORDPERFECT_HELP = 509
'Public Const ID_EMAIL_HELP_CHECK_FOR_UPDATES = 510
'Public Const ID_EMAIL_HELP_DETECT_AND_REPAIR = 511
'Public Const ID_EMAIL_HELP_ACTIVATE_PRODUCT = 512
'Public Const ID_EMAIL_HELP_CUSTOMER_FEEDBACK_OPTIONS = 513
'Public Const ID_EMAIL_HELP_ABOUT_MICROSOFT_OFFICE_WORD = 514
'
'Public Const ID_FINDBAR_COMBO = 600
'Public Const ID_FINDBAR_SEARCHIN = 601
'Public Const ID_FINDBAR_EDIT = 602
'Public Const ID_FINDBAR_FINDNOW = 603
'Public Const ID_FINDBAR_CLEAR = 604
'Public Const ID_FINDBAR_OPTIONS = 605
'Public Const ID_FINDBAR_CLOSE = 606
'
'




Public Const ID_TAB_HOME = 130

Public Const ID_TAB_EDIT = 133
Public Const ID_TAB_PRINT_PREVIEW = 134


Public Const ID_GROUP_FILE = 130
Public Const ID_GROUP_DOCUMENTVIEWS = 134

Public Const ID_GROUP_WINDOW = 136


Public Const ID_GROUP_EDITING = 139

Public Const ID_VIEW_NORMAL = 141
Public Const ID_VIEW_FULLSCREEN = 142
Public Const ID_WINDOW_SWITCH = 143



Public Const ID_VIEW_WORKSPACE = 59394
Public Const ID_PREVIEW_PRINT_PRINT = 5050
Public Const ID_PREVIEW_PRINT_OPTIONS = 5051
Public Const ID_PREVIEW_PAGESETUP_MARGINS = 5052
Public Const ID_PREVIEW_PAGESETUP_ORIENTATION = 5053
Public Const ID_PREVIEW_PAGESETUP_SIZE = 5054
Public Const ID_PREVIEW_ZOOM_ZOOM = 5055
Public Const ID_PREVIEW_ZOOM_100_PERCENT = 5056
Public Const ID_PREVIEW_ZOOM_1PAGE = 5057
Public Const ID_PREVIEW_ZOOM_2PAGES = 5058
Public Const ID_PREVIEW_ZOOM_PAGE_WIDTH = 5059
Public Const ID_PREVIEW_PREVIEW_RULER = 5060
Public Const ID_PREVIEW_PREVIEW_MAGNIFIER = 5061
Public Const ID_PREVIEW_PREVIEW_SHRINK = 5062
Public Const ID_PREVIEW_PREVIEW_NEXT = 5063
Public Const ID_PREVIEW_PREVIEW_PREVIOUS = 5064
Public Const ID_PREVIEW_PREVIEW_CLOSE = 5065
Public Const ID_GROUP_PREVIEW = 5070
Public Const ID_GROUP_ZOOM = 5071
Public Const ID_GROUP_PRINT = 5072
Public Const ID_MARGINS_CUSTOM_MARGINS = 5073
Public Const ID_ORIENTATION_PORTRAIT = 5074
Public Const ID_ORIENTATION_LANDSCAPE = 5075
Public Const ID_SIZE_MORE_PAPER_SIZES = 5076

Public Const ID_VIEW_MESSAGEBAR = 2815
Public Const ID_GROUP_ADVANCED = 3431
Public Const ID_GROUP_HYPERLINK = 3432
Public Const ID_GROUP_MARKUPLABEL = 3433
Public Const ID_GROUP_BITMAP = 3434
Public Const ID_TAB_ADVANCED = 3435

Public Const ID_VIEW_STATUS_BAR = 2808

Public Const ID_TAB_CALENDAR_HOME = 12000
Public Const ID_GROUP_NEW = 12001
Public Const ID_GROUP_NEW_APPOINTMENT = 12002
Public Const ID_GROUP_NEW_MEETING = 12003
Public Const ID_GROUP_NEW_ITEMS = 12044
Public Const ID_GROUP_NEW_ALLDAY = 12051
Public Const ID_GROUP_GOTO = 12005
Public Const ID_GROUP_GOTO_TODAY = 12006
Public Const ID_GROUP_GOTO_NEXT7DAYS = 12007
Public Const ID_GROUP_ARRANGE2 = 12008
Public Const ID_GROUP_ARRANGE_DAY = 12009
Public Const ID_GROUP_ARRANGE_WORK_WEEK = 12010
Public Const ID_GROUP_ARRANGE_WEEK = 12012
Public Const ID_GROUP_ARRANGE_MONTH = 12012
Public Const ID_GROUP_ARRANGE_MONTH_LOW = 12052
Public Const ID_GROUP_ARRANGE_MONTH_MEDIUM = 12053
Public Const ID_GROUP_ARRANGE_MONTH_HIGH = 12054
Public Const ID_GROUP_ARRANGE_SCHEDULE_VIEW = 12013
Public Const ID_GROUP_MANAGE = 12023
Public Const ID_GROUP_MANAGE_CALENDARS_OPEN = 12014
Public Const ID_GROUP_MANAGE_CALENDARS_GROUPS = 12015
Public Const ID_GROUP_SHARE = 12024
Public Const ID_GROUP_SHARE_EMAIL = 12016
Public Const ID_GROUP_SHARE_SHARE = 12017
Public Const ID_GROUP_SHARE_PUBLISH = 12018
Public Const ID_GROUP_SHARE_PERMISSIONS = 12019
Public Const ID_GROUP_FIND2 = 12020
Public Const ID_GROUP_FIND2_CONTACT = 12021
Public Const ID_GROUP_FIND2_ADDRESSBOOK = 12022

Public Const ID_TAB_MAIL_HOME = 12114
Public Const ID_TAB_SEND_RECEIVE = 12110
Public Const ID_TAB_FOLDER = 12111
Public Const ID_TAB_VIEW2 = 12112
Public Const ID_TAB_ADDINS2 = 12113


Public Const FSHIFT = 4
Public Const FALT = 16

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B



Public Const ID_WINDOW_NEW = 57648
Public Const ID_WINDOW_ARRANGE = 57649






'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
'
'
'                   Constantes de ARIGES
'
'
'****************************************************************************************
'****************************************************************************************
'
'select concat('Public Const id_',
'replace(replace(replace(menus.descripcion,' ',''),'.',''),'-','_')
', ' = ',menus.codigo ,'     ',substring('                             ',1,30- length(menus.descripcion)),'\'', m2.descripcion,'(',menus.padre,')') as datos
' from menus left join menus m2 on menus.padre=m2.codigo where menus.aplicacion='ariges' and menus.aplicacion=m2.aplicacion and menus.padre >0 order by menus.padre,menus.grupo,menus.orden


Public Const id_Empresa = 101                            'Configuración(1)
Public Const id_ParámetrosAplicación = 102              'Configuración(1)
Public Const id_TiposMovimiento = 104                   'Configuración(1)
Public Const id_TiposDocumentos = 105                   'Configuración(1)
Public Const id_Usuarios = 106                           'Configuración(1)

Public Const id_acciones_inicio = 149   ' No tiene punto de menu. Es para centralizar todas las llamadas de abrirforms


'Del 150 al 200 los reservo para los agrupados2

Public Const id_Marcas = 201                             'Almacen(2)
Public Const id_AlmacenesPropios = 202                  'Almacen(2)
Public Const id_TiposUnidad = 203                       'Almacen(2)
Public Const id_TiposArtículos = 204                    'Almacen(2)
Public Const id_Ubicaciones = 205                        'Almacen(2)
Public Const id_Familias = 206                           'Almacen(2)
Public Const id_Categorias = 207                         'Almacen(2)
Public Const id_Artículos = 208                          'Almacen(2)
Public Const id_Númerosdelote = 209                    'Almacen(2)
Public Const id_Telematel = 238                          'Almacen(2)
Public Const id_TraspasoAlmacenes = 210                 'Almacen(2)
Public Const id_HcoTraspasoAlmacenes = 211             'Almacen(2)
Public Const id_MovimientosAlmacén = 212                'Almacen(2)
Public Const id_HcoMovimientosAlmacén = 213            'Almacen(2)
Public Const id_MovimientosArtículos = 214              'Almacen(2)
Public Const id_MovimientosStockdesdeInv = 215       'Almacen(2)
Public Const id_InfControlStock = 216                  'Almacen(2)
Public Const id_InfArtículosinactivos = 217            'Almacen(2)
Public Const id_InfArtículoscomponentes = 218          'Almacen(2)
Public Const id_InfValoraciónstocks = 219              'Almacen(2)
Public Const id_InfStocksmax_min = 220                 'Almacen(2)
Public Const id_InfStocksaFecha = 221                 'Almacen(2)
Public Const id_InfStocksxmeses = 222                 'Almacen(2)
Public Const id_InfAlertasPPedido = 223               'Almacen(2)
Public Const id_InfReposiciónAlmacén = 224             'Almacen(2)
Public Const id_InfStockmínimo = 225                   'Almacen(2)
Public Const id_MovimientosLotes = 226                  'Almacen(2)
Public Const id_TomadeInventario = 227                 'Almacen(2)
Public Const id_EntradaExistencia = 228                 'Almacen(2)
Public Const id_ListadoDiferencias = 229                'Almacen(2)
Public Const id_Actualizardiferencias = 230             'Almacen(2)
Public Const id_ValoraciónStocksInv = 231             'Almacen(2)
Public Const id_RectificarúltimoInv = 232             'Almacen(2)
Public Const id_InventariarArtículo = 233               'Almacen(2)
Public Const id_RecálculoPrStandard = 234              'Almacen(2)
Public Const id_RecálculoPrMedioP = 235              'Almacen(2)
Public Const id_RecálculoUltPrCompra = 236            'Almacen(2)
Public Const id_HistóricoInventario = 237               'Almacen(2)
Public Const id_Actividades = 301                        'Datos Básicos Ventas(3)
Public Const id_Zonas = 302                              'Datos Básicos Ventas(3)
Public Const id_Rutas = 303                              'Datos Básicos Ventas(3)
Public Const id_Portes = 304                             'Datos Básicos Ventas(3)
Public Const id_Descuentosporcantidad = 305            'Datos Básicos Ventas(3)
Public Const id_FormasdeEnvío = 306                    'Datos Básicos Ventas(3)
Public Const id_FormasdePago = 307                     'Datos Básicos Ventas(3)
Public Const id_Bancospropios = 308                     'Datos Básicos Ventas(3)
Public Const id_SituacionesEspeciales = 309             'Datos Básicos Ventas(3)
Public Const id_Agentes = 310                            'Datos Básicos Ventas(3)
Public Const id_Clientesvarios = 311                    'Datos Básicos Ventas(3)
Public Const id_Clientes = 312                           'Datos Básicos Ventas(3)
Public Const id_ClientesPotenciales = 313               'Datos Básicos Ventas(3)
Public Const id_TiposdeCartas = 314                    'Datos Básicos Ventas(3)
Public Const id_Incidencias = 315                        'Datos Básicos Ventas(3)
Public Const id_ClientesInactivos = 316                 'Datos Básicos Ventas(3)
Public Const id_AltasClientes = 317                     'Datos Básicos Ventas(3)
Public Const id_EtiquetasdeClientes = 318              'Datos Básicos Ventas(3)
Public Const id_CartasaClientes = 319                  'Datos Básicos Ventas(3)
Public Const id_Etiquetasdebultos = 320                'Datos Básicos Ventas(3)
Public Const id_InfTeléfonosxCliente = 321            'Datos Básicos Ventas(3)
Public Const id_InfCuotastelefonía = 322               'Datos Básicos Ventas(3)
Public Const id_TarifasVenta = 331                      'Datos Básicos Ventas(3)
Public Const id_ListaPrecios = 332                      'Datos Básicos Ventas(3)
Public Const id_PreciosEspeciales = 333                 'Datos Básicos Ventas(3)
Public Const id_Promociones = 334                        'Datos Básicos Ventas(3)
Public Const id_DtosFamilia_Marca = 335                'Datos Básicos Ventas(3)
Public Const id_DtosxActividad = 336                  'Datos Básicos Ventas(3)
Public Const id_ActualizarPrecios = 337                 'Datos Básicos Ventas(3)
Public Const id_Copiarpreciosdesdecompra = 338        'Datos Básicos Ventas(3)
Public Const id_InfControlmargenTarifa = 339          'Datos Básicos Ventas(3)
Public Const id_CorregirerroresyActTarifas = 340     'Datos Básicos Ventas(3)
Public Const id_ControlerrorDtosCliente = 341         'Datos Básicos Ventas(3)
Public Const id_TarifasTaxímetros = 342                 'Datos Básicos Ventas(3)

Public Const id_Ofertas = 401                            'Ofertas-Pedidos Ventas(4)
Public Const id_GrupodePlantillas = 402                'Ofertas-Pedidos Ventas(4)
Public Const id_Plantillas = 403                         'Ofertas-Pedidos Ventas(4)
Public Const id_InfOfertasefectuadas = 404             'Ofertas-Pedidos Ventas(4)
Public Const id_Pedidos = 405                            'Ofertas-Pedidos Ventas(4)
Public Const id_HcoPedidosanulados = 406               'Ofertas-Pedidos Ventas(4)
Public Const id_CartasConfirmaciónPedidos = 407       'Ofertas-Pedidos Ventas(4)
Public Const id_InfPedidosxArtículo = 408             'Ofertas-Pedidos Ventas(4)
Public Const id_InfPedidosxCliente = 409              'Ofertas-Pedidos Ventas(4)
Public Const id_InfDisponibilidadStocks = 410          'Ofertas-Pedidos Ventas(4)
Public Const id_ImpresiónPedidosxZona = 411           'Ofertas-Pedidos Ventas(4)
Public Const id_ConsultaPreciosxCliente = 412         'Ofertas-Pedidos Ventas(4)
Public Const id_Devoluciónmaterial = 413                'Ofertas-Pedidos Ventas(4)
Public Const id_InfPedidosxdia = 414                  'Ofertas-Pedidos Ventas(4)
Public Const id_Alertas = 415                            'Ofertas-Pedidos Ventas(4)

Public Const id_Albaranes = 501                          'Albaranes Ventas(5)
Public Const id_HcoAlbaranesanulados = 502             'Albaranes Ventas(5)
Public Const id_AlbaranesDevolución = 503               'Albaranes Ventas(5)
Public Const id_InfAlbaranesxArtículo = 504           'Albaranes Ventas(5)
Public Const id_InfIncumplimientoentrega = 505         'Albaranes Ventas(5)
Public Const id_InfSituaciónAlbaranes = 506            'Albaranes Ventas(5)
Public Const id_ControlAlbaranes = 507                  'Albaranes Ventas(5)
Public Const id_ControlAlbaranesFact = 508            'Albaranes Ventas(5)
Public Const id_ControlDirecEnvío = 509                'Albaranes Ventas(5)
Public Const id_ImpresiónAlbTransporte = 510           'Albaranes Ventas(5)
Public Const id_InfAlbaranes = 511                      'Albaranes Ventas(5)
Public Const id_InfAlbaranesentregados = 512           'Albaranes Ventas(5)
Public Const id_FacturasMostrador = 513                 'Albaranes Ventas(5)
Public Const id_FacturasRectificativas = 514            'Albaranes Ventas(5)
Public Const id_AlbaranesServicios = 515                'Albaranes Ventas(5)
Public Const id_FacturaciónServicios = 516              'Albaranes Ventas(5)
Public Const id_AlbaranesInternos = 517                 'Albaranes Ventas(5)
Public Const id_FacturaciónInternos = 518               'Albaranes Ventas(5)
Public Const id_InfAlbaranesInternos = 519             'Albaranes Ventas(5)
Public Const id_AlbaranesGasolinera = 520               'Albaranes Ventas(5)
Public Const id_AlbaranesTienda = 521                   'Albaranes Ventas(5)
Public Const id_ImportarficheroGasolinera = 523        'Albaranes Ventas(5)
Public Const id_CambiarAlbaranes_Facturas = 524         'Albaranes Ventas(5)
Public Const id_Previsión = 525                          'Albaranes Ventas(5)
Public Const id_Combustible = 526                        'Albaranes Ventas(5)
Public Const id_Tienda = 527                             'Albaranes Ventas(5)
Public Const id_AjusteFormasdePago = 528              'Albaranes Ventas(5)
Public Const id_AlbaranesTelefonía = 529                'Albaranes Ventas(5)
Public Const id_ImportarficheroTelefonía = 530         'Albaranes Ventas(5)
Public Const id_Datospendientesfacturar = 531          'Albaranes Ventas(5)
Public Const id_CargosVarios = 534                      'Albaranes Ventas(5)
Public Const id_Modificaciónmasivacuotas = 535         'Albaranes Ventas(5)
Public Const id_Comparativadescuentos = 536             'Albaranes Ventas(5)
Public Const id_Facturaciónporsoporte = 537            'Albaranes Ventas(5)
Public Const id_Resumenporsoporte = 538                'Albaranes Ventas(5)
Public Const id_Datosimportaciónfichero = 539          'Albaranes Ventas(5)
Public Const id_Conceptosconsumos = 540                 'Albaranes Ventas(5)
Public Const id_Descuentosconsumos = 541                'Albaranes Ventas(5)
Public Const id_Conceptoscuotas = 542                   'Albaranes Ventas(5)
Public Const id_Descuentoscuotas = 543                  'Albaranes Ventas(5)
Public Const id_Cuotaspropias = 544                     'Albaranes Ventas(5)
Public Const id_Parámetros = 545                         'Albaranes Ventas(5)
Public Const id_Calibres = 546                           'Albaranes Ventas(5)
Public Const id_Contadores = 547                         'Albaranes Ventas(5)
Public Const id_Importarfichero = 548                   'Albaranes Ventas(5)
Public Const id_Facturaciónagua = 549                   'Albaranes Ventas(5)
Public Const id_Resumenfacturación = 550                'Albaranes Ventas(5)
Public Const id_InfFacturaciónxperiodo = 551          'Albaranes Ventas(5)
Public Const id_InfContadoresexportación = 552         'Albaranes Ventas(5)
Public Const id_Modificarcuotavarios = 553             'Albaranes Ventas(5)
Public Const id_Declaracióndetalladaejercicio = 554     'Albaranes Ventas(5)
Public Const id_InfTasaspendcobro = 555               'Albaranes Ventas(5)
Public Const id_Materiasactivas = 556                   'Albaranes Ventas(5)
Public Const id_ADR = 557                                'Albaranes Ventas(5)
Public Const id_Plagas = 558                             'Albaranes Ventas(5)
Public Const id_Flotas = 559                             'Albaranes Ventas(5)
Public Const id_Tratamientos = 560                       'Albaranes Ventas(5)
Public Const id_Partesdetrabajo = 561                  'Albaranes Ventas(5)
Public Const id_InfFitosanitarios_Campos = 562          'Albaranes Ventas(5)
Public Const id_Ajustecomprastrat = 563               'Albaranes Ventas(5)
Public Const id_Capítulos = 564                          'Albaranes Ventas(5)
Public Const id_Actuaciones = 565                        'Albaranes Ventas(5)
Public Const id_Partesdetrabajo2 = 566                  'Albaranes Ventas(5)
Public Const id_Tiposordenes = 567                      'Albaranes Ventas(5)
Public Const id_Reloj = 568                              'Albaranes Ventas(5)
Public Const id_InfCompras_Ventasactuación = 569       'Albaranes Ventas(5)
Public Const id_ImpresiónCertificación = 570            'Albaranes Ventas(5)
Public Const id_InfHuertos_Hanegadas = 571              'Albaranes Ventas(5)
Public Const id_Facturaciónderrama = 572                'Albaranes Ventas(5)

Public Const id_PrevisiónFActuración = 601              'Facturación Ventas(6)
Public Const id_Facturaciónalbaranes = 602              'Facturación Ventas(6)
Public Const id_Facturarcliente = 603                   'Facturación Ventas(6)
Public Const id_HistóricoFacturasVenta = 604           'Facturación Ventas(6)
Public Const id_ReimpresiónFacturas = 605               'Facturación Ventas(6)
Public Const id_EnvíoFacturasxmail = 606              'Facturación Ventas(6)
Public Const id_EnvíoFacturasweb = 607                 'Facturación Ventas(6)
Public Const id_ContabilizarFacturas = 608              'Facturación Ventas(6)
Public Const id_Contabilizarticketsagrupados = 609     'Facturación Ventas(6)
Public Const id_InfTicketsfacturados = 610             'Facturación Ventas(6)
Public Const id_InfporCliente = 611                   'Facturación Ventas(6)
Public Const id_InfporTrabajador = 612                'Facturación Ventas(6)
Public Const id_Infpormeses = 613                     'Facturación Ventas(6)
Public Const id_InfporFamilia_Artículo = 614          'Facturación Ventas(6)
Public Const id_InfporArtículo = 615                  'Facturación Ventas(6)
Public Const id_InfporProveedor = 616                 'Facturación Ventas(6)
Public Const id_InfporAgente = 617                    'Facturación Ventas(6)
Public Const id_DetalledeFacturación = 618             'Facturación Ventas(6)
Public Const id_Margenventas = 619                      'Facturación Ventas(6)
Public Const id_Infportipoprecio = 620               'Facturación Ventas(6)
Public Const id_Artículosmayorventa = 621              'Facturación Ventas(6)
Public Const id_InfporFamiliaagrupado = 622          'Facturación Ventas(6)
Public Const id_Infportipopedido = 623               'Facturación Ventas(6)
Public Const id_InfCostes = 624                        'Facturación Ventas(6)

Public Const id_Proveedores = 901                        'Compras(9)
Public Const id_ProveedoresVarios = 902                 'Compras(9)
Public Const id_DireccionesCompra = 903                 'Compras(9)
Public Const id_EtiquetasdeProveedores = 904           'Compras(9)
Public Const id_CartasaProveedores = 905               'Compras(9)
Public Const id_EtiquetasdebultosCompra = 906                'Compras(9)
Public Const id_PreciosProveedor = 907                  'Compras(9)
Public Const id_DescuentosProveedor = 908               'Compras(9)
Public Const id_CopiarPreciosdesdeventa = 909         'Compras(9)
Public Const id_ActualizarPreciosCompra = 910          'Compras(9)
Public Const id_PedidosProveedor = 911                  'Compras(9)
Public Const id_HcoPedidosanuladosCompra = 912               'Compras(9)
Public Const id_InfMaterialpdterecibir = 913        'Compras(9)
Public Const id_PropuestadePedido = 914                'Compras(9)
Public Const id_InfReaprovisionamiento = 915           'Compras(9)
Public Const id_AlbaranesProveedor = 916                'Compras(9)
Public Const id_HcoAlbaranesanuladosCompra = 917      'Compras(9)
Public Const id_InfPendientefacturar = 918            'Compras(9)
Public Const id_ControlAlbaranesCompra = 919                  'Compras(9)
Public Const id_ControlAlbaranesfacturados = 920       'Compras(9)
Public Const id_RecepciónFacturas = 921                 'Compras(9)
Public Const id_HistóricoFacturasCompra = 922          'Compras(9)
Public Const id_ContabilizarFacturasCompra = 923              'Compras(9)
Public Const id_InfporProveedorCompra = 924                 'Compras(9)
Public Const id_InfporFamilia_ArtículoCompra = 925          'Compras(9)
Public Const id_InfpormesesCompra = 926                     'Compras(9)
Public Const id_InfAlbaranesxProveedor = 927         'Compras(9)
Public Const id_InfPrevisiónpagos = 928               'Compras(9)
Public Const id_InfProveedor_Marca_Familia = 929       'Compras(9)
Public Const id_Trabajadores = 1001                       'Administración(10)
Public Const id_GastosTécnicos = 1002                    'Administración(10)
Public Const id_NóminasyGastos = 1003                   'Administración(10)
Public Const id_CálculoRiesgo = 1004                     'Administración(10)
Public Const id_InfRiesgo = 1005                        'Administración(10)
Public Const id_Correccióncostesvariosfactura = 1007     'Administración(10)
Public Const id_CorreccióncostesEstVentas = 1008       'Administración(10)
Public Const id_InfVentasacrédito = 1009              'Administración(10)
Public Const id_BeneficioProveedor = 1010                'Administración(10)
Public Const id_BeneficioCliente = 1011                  'Administración(10)
Public Const id_BeneficioMarca_Agente_Proveedor = 1012     'Administración(10)
Public Const id_InfArtículosenpromoción = 1013        'Administración(10)
Public Const id_InfArtículosconDtoEspecial = 1014     'Administración(10)
Public Const id_InfVentasTrabajadordía = 1015         'Administración(10)
Public Const id_InfVentasxFPago = 1016               'Administración(10)
Public Const id_InfComparativoDtosCompra_Venta = 1017     'Administración(10)
Public Const id_ResumenVentasAgente = 1018              'Administración(10)
Public Const id_BeneficioAgente = 1019                   'Administración(10)
Public Const id_InfVentasAgente_Trabajador = 1020      'Administración(10)
Public Const id_InfComisionesECO = 1021                'Administración(10)
Public Const id_InfAgente_Familia_Marca = 1022          'Administración(10)
Public Const id_InfAgente_Marca_Familia = 1023          'Administración(10)
Public Const id_RegistroGastos = 1024                    'Administración(10)
Public Const id_FlotasAdm = 1025                             'Administración(10)
Public Const id_Conceptos = 1026                          'Administración(10)
Public Const id_TiposdeContrato = 1101                  'Mantenimientos(11)
Public Const id_Mantenimientos = 1102                     'Mantenimientos(11)
Public Const id_InfMantenimientos = 1103                'Mantenimientos(11)
Public Const id_InfRevisiones = 1104                    'Mantenimientos(11)
Public Const id_FichaMantenimientos = 1105               'Mantenimientos(11)
Public Const id_InfAltas = 1106                         'Mantenimientos(11)
Public Const id_InfTeórico = 1107                       'Mantenimientos(11)
Public Const id_Etiquetas = 1108                          'Mantenimientos(11)
Public Const id_Cartasrenovación = 1109                  'Mantenimientos(11)
Public Const id_Traspasosiguienteaactual = 1110        'Mantenimientos(11)
Public Const id_HcoMantenimientos = 1111                'Mantenimientos(11)
Public Const id_InfAnulados = 1112                      'Mantenimientos(11)
Public Const id_PrevisiónMantenimientos = 1113           'Mantenimientos(11)
Public Const id_FacturaciónMantenimientos = 1114         'Mantenimientos(11)
Public Const id_PrevisiónRenting = 1115                  'Mantenimientos(11)
Public Const id_FacturaciónRenting = 1116                'Mantenimientos(11)

Public Const id_NúmerosdeSerie = 1201                   'Reparaciones(12)
Public Const id_Motivosbaja = 1202                       'Reparaciones(12)
Public Const id_MotivosPendienteRep = 1203             'Reparaciones(12)
Public Const id_Tiposaveria = 1204                       'Reparaciones(12)
Public Const id_Trabajosrealizados = 1205                'Reparaciones(12)
Public Const id_Serviciosasistenciatécnica = 1206       'Reparaciones(12)
Public Const id_EntradaReparación = 1210                 'Reparaciones(12)
Public Const id_ControlReparación = 1211                 'Reparaciones(12)
Public Const id_HcoReparaciones = 1212                  'Reparaciones(12)
Public Const id_Infpordía = 1213                       'Reparaciones(12)
Public Const id_InfporClienteRepar = 1214                   'Reparaciones(12)
Public Const id_FrecuenciadeReparación = 1215           'Reparaciones(12)
Public Const id_InfporTécnico = 1216                   'Reparaciones(12)
Public Const id_InfReparacionesefectuadas = 1217       'Reparaciones(12)
Public Const id_InfGarantíaproveedor = 1218            'Reparaciones(12)
Public Const id_AlbaránOrdendetrabajo = 1219           'Reparaciones(12)
Public Const id_AlbaránServExterior = 1220              'Reparaciones(12)
Public Const id_AlbaránReparación = 1221                 'Reparaciones(12)
Public Const id_Proyectos = 1222                          'Reparaciones(12)
Public Const id_Frecuencias = 1207                        'Reparaciones(12)
Public Const id_AlbaránReparaciónRepara = 1223                 'Reparaciones(12)
Public Const id_PrevisiónFActuraciónRepara = 1224              'Reparaciones(12)
Public Const id_Facturación = 1225                        'Reparaciones(12)

Public Const id_TiposAcciones = 1301                     'CRM(13)
Public Const id_Conceptosllamadas = 1302                 'CRM(13)
Public Const id_Accionescomerciales = 1303               'CRM(13)
Public Const id_Generaracciones = 1304                   'CRM(13)
Public Const id_InfMasivo = 1305                        'CRM(13)
Public Const id_InfresumenCRM = 1306                    'CRM(13)
Public Const id_InfClientesporacción = 1307             'CRM(13)
Public Const id_AvisosdeClientes = 1308                 'CRM(13)
Public Const id_InfAvisospendientes = 1309             'CRM(13)
Public Const id_Borreavisoscerrados = 1310              'CRM(13)
Public Const id_LlamadasClientes = 1311                  'CRM(13)

Public Const id_Ordenesproducción = 1401                 'Producción(14)
Public Const id_Ordenesenvasado = 1402                   'Producción(14)
Public Const id_CostesTasas = 1403                       'Producción(14)
Public Const id_Registrotrazabilidad = 1404              'Producción(14)
Public Const id_Parámetroscalidad = 1405                 'Producción(14)
Public Const id_DeclaraciónalcoholAEAT = 1406           'Producción(14)

Public Const id_PantalladeVenta = 1501                  'Punto de Venta(15)
Public Const id_Cierredecaja = 1502                     'Punto de Venta(15)
Public Const id_Etiquetasestantería = 1503               'Punto de Venta(15)
Public Const id_ParámetrosGenerales = 1504               'Punto de Venta(15)
Public Const id_ParámetrosTerminales = 1505              'Punto de Venta(15)

Public Const id_Accionesrealizadas = 2001                'Utilidades(20)
Public Const id_BorreFacturasyMovimientos = 2002       'Utilidades(20)
Public Const id_CambiodeCliente = 2003                  'Utilidades(20)
Public Const id_InfAlb_Pedanulados = 2004              'Utilidades(20)
Public Const id_ComprobarCCC_NIF = 2005                  'Utilidades(20)
Public Const id_ExportarAlbaranesservicio = 2006        'Utilidades(20)
Public Const id_ExportarmailCSV = 2007                  'Utilidades(20)
Public Const id_Lotesfitossubvencionados = 2008        'Utilidades(20)
Public Const id_DeclaracionROPO = 2009                   'Utilidades(20)



































'#####   Para mostrar u acoultar puntos de menus debido a la instalacion
'
'        Solo los ocultarPorInstalacion =0 seran visibles
'           Si valor=1 es ocultado por que no lleva telfonia, no es Herbelca, ses euler.....
'           Si el valor es 2 lo ocultamos pq ariadna quiere que no sea visible
Public Sub ActualizarTablasMenusArigesNuevo()
Dim B As Boolean
Dim SQL As String


    'Ponemos todos los que puedan ser visibles a 0
    SQL = "UPDATE menus SET ocultarPorInstalacion =0 where aplicacion='ariges' and ocultarPorInstalacion =1"
    conn.Execute SQL
    
    SQL = "" 'reestablezco
    

    'ALBARAN INTERNO
    B = DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "ALI", "T") <> ""
    
    If Not B Then SQL = SQL & ", " & id_AlbaranesInternos
    If Not (B And vParamAplic.NumeroInstalacion <> vbEuler) Then SQL = SQL & ", " & id_FacturaciónInternos
    If Not B Then SQL = SQL & ", " & id_InfAlbaranesInternos
    
    'SERVICIOS
    If Not vParamAplic.Servicios Then SQL = SQL & ", " & id_AlbaranesServicios
    If Not (vParamAplic.Servicios And vParamAplic.NumeroInstalacion <> vbEuler) Then SQL = SQL & ", " & id_FacturaciónServicios  'euler, si hay no los factura
    

    'Gasolinera
    If Not vParamAplic.TieneGasolinera Then
        SQL = SQL & ", " & id_AlbaranesGasolinera & ", " & id_AlbaranesTienda & ", " & id_ImportarficheroGasolinera & ", " & id_CambiarAlbaranes_Facturas
        SQL = SQL & ", " & id_Previsión & ", " & id_Combustible & ", " & id_Tienda & ", " & id_AjusteFormasdePago
    End If
    
    
    'REPARACIONES
    If Not vParamAplic.Reparaciones Then
        SQL = SQL & ", " & id_NúmerosdeSerie & ", " & id_Motivosbaja & ", " & id_MotivosPendienteRep & ", " & id_Tiposaveria
        SQL = SQL & ", " & id_Trabajosrealizados & ", " & id_Serviciosasistenciatécnica & ", " & id_EntradaReparación & ", " & id_ControlReparación
        SQL = SQL & ", " & id_HcoReparaciones & ", " & id_Infpordía & ", " & id_InfporClienteRepar & ", " & id_FrecuenciadeReparación
        SQL = SQL & ", " & id_InfporTécnico & ", " & id_InfReparacionesefectuadas & ", " & id_InfGarantíaproveedor & ", " & id_AlbaránOrdendetrabajo
        SQL = SQL & ", " & id_InfporTécnico & ", " & id_InfReparacionesefectuadas & ", " & id_InfGarantíaproveedor
        SQL = SQL & ", " & id_AlbaránOrdendetrabajo & ", " & id_AlbaránServExterior & ", " & id_AlbaránReparación
        SQL = SQL & ", " & id_Proyectos & ", " & id_Frecuencias & ", " & id_AlbaránReparaciónRepara
        SQL = SQL & ", " & id_PrevisiónFActuraciónRepara & ", " & id_Facturación
        
    Else
        'Si que tiene reparacon, pero si no es eulertaxco NO iran ciertos puntos
        'Si es EULER
        If Not vParamAplic.NumeroInstalacion = InstalacionEsEulerTaxco Then SQL = SQL & ", " & id_AlbaránOrdendetrabajo & ", " & id_AlbaránServExterior
        
    End If
    
    'Mantenimientos
    If Not vParamAplic.Mantenimientos Then
        SQL = SQL & ", " & id_TiposdeContrato & ", " & id_Mantenimientos & ", " & id_InfMantenimientos & ", " & id_InfRevisiones
        SQL = SQL & ", " & id_FichaMantenimientos & ", " & id_InfAltas & ", " & id_InfTeórico & ", " & id_Etiquetas
        SQL = SQL & ", " & id_Etiquetas & ", " & id_Cartasrenovación & ", " & id_Traspasosiguienteaactual & ", " & id_HcoMantenimientos
        SQL = SQL & ", " & id_InfAnulados & ", " & id_PrevisiónMantenimientos & ", " & id_FacturaciónMantenimientos & ", " & id_PrevisiónRenting
        SQL = SQL & ", " & id_FacturaciónRenting
    End If

    'Telefonia
     If vParamAplic.TieneTelefonia2 = 0 Then
        SQL = SQL & ", " & id_AlbaranesTelefonía & ", " & id_ImportarficheroTelefonía & ", " & id_Datospendientesfacturar & ", " & id_CargosVarios
        SQL = SQL & ", " & id_Modificaciónmasivacuotas & ", " & id_Comparativadescuentos & ", " & id_Facturaciónporsoporte & ", " & id_Resumenporsoporte
        SQL = SQL & ", " & id_Datosimportaciónfichero & ", " & id_Conceptosconsumos & ", " & id_Descuentosconsumos & ", " & id_Conceptoscuotas
        SQL = SQL & ", " & id_Descuentoscuotas & ", " & id_Cuotaspropias
        
        'Y dos listados
        SQL = SQL & ", " & id_InfTeléfonosxCliente & ", " & id_InfCuotastelefonía
        
        
    End If
    
    
    'Tickets agrupados
    If Not vParamAplic.ContabilizarTicketAgrupados Then SQL = SQL & ", " & id_Contabilizarticketsagrupados & ", " & id_InfTicketsfacturados
    
    'Produccion
    B = vParamAplic.Produccion
    If Not B Then B = vParamAplic.TieneComponentes_y_Produccion
    If Not B Then
        'PuntoDeMenuVisible Me.mnproduccion, B
        SQL = SQL & ", " & id_Ordenesproducción & ", " & id_Ordenesenvasado & ", " & id_CostesTasas & ", " & id_Registrotrazabilidad
        SQL = SQL & ", " & id_Parámetroscalidad & ", " & id_DeclaraciónalcoholAEAT
    Else
        'Si tiene produccion pero no es fontenas, no ve la declaracion de alcoho
        If vParamAplic.NumeroInstalacion <> vbFontenas Then SQL = SQL & ", " & id_DeclaraciónalcoholAEAT
    End If
    
    'CRM
    If Not vParamAplic.TieneCRM Then
        SQL = SQL & ", " & id_TiposAcciones & ", " & id_Conceptosllamadas
        SQL = SQL & ", " & id_Accionescomerciales & ", " & id_Generaracciones
        SQL = SQL & ", " & id_InfMasivo & ", " & id_InfresumenCRM
        SQL = SQL & ", " & id_InfClientesporacción & ", " & id_InfClientesporacción
        SQL = SQL & ", " & id_InfAvisospendientes & ", " & id_Borreavisoscerrados
        SQL = SQL & ", " & id_LlamadasClientes
    End If

    'FLOTAS
    If Not vParamAplic.GestionFlotas Then
        SQL = SQL & ", " & id_RegistroGastos & ", " & id_FlotasAdm
        SQL = SQL & ", " & id_Conceptos & ", " & id_Generaracciones
    End If
    
    'OBRAS
    If Not (vParamAplic.HayDeparNuevo = 2) Then
        SQL = SQL & ", " & id_Capítulos & ", " & id_Actuaciones & ", " & id_Partesdetrabajo2
        SQL = SQL & ", " & id_Tiposordenes & ", " & id_Reloj & ", " & id_InfCompras_Ventasactuación & ", " & id_ImpresiónCertificación
    End If

    'PEDIDOS DE DEVOLUCION
    B = False
    If DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "PEW", "T") <> "" Then B = True
    If Not B Then SQL = SQL & ", " & id_Devoluciónmaterial
        
    'PEdidos B
    'B = False
    'If vParamAplic.NumeroInstalacion = 2 Then
    '    'Tiene almacen B
    '    If DevuelveDesdeBD(conAri, "codtipom", "stipom", "codtipom", "PEZ", "T") <> "" Then
    '        If vUsu.AlmacenPorDefecto2 = CStr(vParamAplic.AlmacenB) Then B = True
    '    End If
    'End If
    'PuntoDeMenuVisible Me.mnFacPedidos(13), B
   


    'Facturacion electronica
    If vParamAplic.PathFacturaE = "" Then SQL = SQL & ", " & id_EnvíoFacturasweb
    
    
    'TRATAMIENTOS ADV
    If Not vParamAplic.LlevaADV Then
        SQL = SQL & ", " & id_Materiasactivas & ", " & id_ADR & ", " & id_Plagas
        SQL = SQL & ", " & id_Flotas & ", " & id_Tratamientos & ", " & id_Tratamientos & ", " & id_Partesdetrabajo
        SQL = SQL & ", " & id_InfFitosanitarios_Campos & ", " & id_Ajustecomprastrat
    End If
    
    'Renting y servicios
    If Not vParamAplic.Renting Then SQL = SQL & ", " & id_PrevisiónRenting & ", " & id_FacturaciónRenting
    
    If Not vParamAplic.AguasPotables = "" Then
        SQL = SQL & ", " & id_Parámetros & ", " & id_Calibres & ", " & id_Contadores & ", " & id_Importarfichero
        SQL = SQL & ", " & id_Facturaciónagua & ", " & id_Resumenfacturación & ", " & id_InfFacturaciónxperiodo & ", " & id_InfContadoresexportación
        SQL = SQL & ", " & id_Modificarcuotavarios & ", " & id_Declaracióndetalladaejercicio & ", " & id_InfTasaspendcobro
    End If
    
    
    If Not vParamAplic.Huertos Then SQL = SQL & ", " & id_InfHuertos_Hanegadas & ", " & id_Facturaciónderrama
    
    If Not vParamAplic.ClientesPotenciales Then SQL = SQL & ", " & id_ClientesPotenciales
        

    
    If InstalacionEsEulerTaxco Then
        'HAY que  ocultar "ciertos puntos" de menui
        If vParamAplic.NumeroInstalacion = vbTaxco Then
            SQL = SQL & ", " & id_Proyectos & ", " & id_AlbaránReparación & ", " & id_InfCostes   'Solo visible en euler
        Else
            'EULER NO VE
            SQL = SQL & ", " & id_TarifasTaxímetros & ", " & id_CambiodeCliente
        End If
        SQL = SQL & ", " & id_NúmerosdeSerie & ", " & id_Motivosbaja & ", " & id_MotivosPendienteRep & ", " & id_Tiposaveria
        SQL = SQL & ", " & id_Trabajosrealizados & ", " & id_Serviciosasistenciatécnica & ", " & id_EntradaReparación & ", " & id_ControlReparación
        SQL = SQL & ", " & id_HcoReparaciones
        
        SQL = SQL & ", " & id_Infpordía & ", " & id_InfporClienteRepar & ", " & id_FrecuenciadeReparación
        SQL = SQL & ", " & id_InfporTécnico & ", " & id_InfReparacionesefectuadas & ", " & id_InfGarantíaproveedor & ", " & id_AlbaránOrdendetrabajo
        SQL = SQL & ", " & id_InfporTécnico & ", " & id_InfReparacionesefectuadas & ", " & id_InfGarantíaproveedor
        
        SQL = SQL & ", " & id_Frecuencias & ", "     ' & id_AlbaránReparaciónRepara
        SQL = SQL & ", " & id_PrevisiónFActuraciónRepara & ", " & id_Facturación
    
       
       
        
    Else
        'No es eler taxco
        SQL = SQL & ", " & id_AlbaránOrdendetrabajo & ", " & id_AlbaránServExterior & ", " & id_AlbaránReparación
        SQL = SQL & ", " & id_Proyectos & ", " & id_InfCostes & ", " & id_TarifasTaxímetros & ", " & id_CambiodeCliente
        
    End If
    
    
    If Not vParamAplic.LotesGeneralitat Then SQL = SQL & ", " & id_Lotesfitossubvencionados
    If Not vParamAplic.ManipuladorFitosanitarios2 Then SQL = SQL & ", " & id_DeclaracionROPO


    
    
    
    If SQL <> "" Then
        SQL = Mid(SQL, 2) 'quitamos la primera coma
        SQL = "UPDATE menus SET ocultarPorInstalacion =1 WHERE aplicacion='ariges' AND ocultarPorInstalacion <=1 AND codigo IN (" & SQL & ")"
        ejecutar SQL, False
    End If
    
    
        
    'falta revisar. No se han añadido puntos de menu
    ' PuntoDeMenuVisible mnUtilidadesVarias(4), B 'vParamAplic.NumeroInstalacion = vbEuler   Configurar PDFS   YA NO ESTA
    ' PuntoDeMenuVisible mnUtilidadesVarias(8), vParamAplic.ImportacionesCoarval
        
        


End Sub





