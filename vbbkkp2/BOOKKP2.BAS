OPTION EXPLICIT
DEFINT A-Z
CONST CELL_WIDTH = 20
CONST COL_DATE = 0
CONST COL_DESCRIPTION = 1
CONST COL_EXPENSE = 3
CONST COL_INCOME = 2
CONST FALSE = 0
CONST FILE_NAME_FILTER$ = "*/<>?[\]| "
CONST MAXIMUM_DATA_LENGTH = 18
CONST MAXIMUM_NUMBER_LENGTH = 13
CONST MAXIMUM_WHOLE_NR_LENGTH = 10
CONST NO_PAGE = 255
CONST NO_ROW = 255
CONST NUMBER_MASK$ = "#############0.00"
CONST ROWS_PER_PAGE = 21
CONST TRUE = -1

TYPE LossGainBlockStr
 Page AS INTEGER
 Row AS INTEGER
END TYPE

TYPE SelectionStr
 Column AS INTEGER
 Page AS INTEGER
 Row AS INTEGER
END TYPE

DECLARE SUB CalculateLossOrGain (Table() AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DECLARE SUB CalculateLossOrGainSelection (Table() AS STRING)
DECLARE SUB Center (x AS INTEGER, Text AS STRING)
DECLARE SUB ClearLossGainBlock (Table() AS STRING, LossGainBlock AS LossGainBlockStr)
DECLARE SUB CreateFile (FileName AS STRING)
DECLARE SUB DisplayCell (Table() AS STRING, Row AS INTEGER, Column AS INTEGER, Page AS INTEGER)
DECLARE SUB DisplayHelp ()
DECLARE SUB DisplayInformation ()
DECLARE SUB DisplayPage (Table() AS STRING, Selection AS SelectionStr)
DECLARE SUB DisplayStatusBar (TableFileName AS STRING, Selection AS SelectionStr)
DECLARE SUB DisplayTableList ()
DECLARE SUB DrawBox (x AS INTEGER, y AS INTEGER, BoxWidth AS INTEGER, BoxHeight AS INTEGER)
DECLARE SUB Export (Table() AS STRING)
DECLARE SUB FormatCell (Table() AS STRING, Previous AS SelectionStr, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DECLARE SUB Import (Table() AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DECLARE SUB Initialize (Table() AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DECLARE SUB InsertRow (Table() AS STRING, AtRow AS INTEGER, AtPage AS INTEGER, LossGainBlock AS LossGainBlockStr)
DECLARE SUB Main ()
DECLARE SUB MoveRow (Table() AS STRING, SourceRow AS INTEGER, SourcePage AS INTEGER, TargetRow AS INTEGER, TargetPage AS INTEGER)
DECLARE SUB PrintSelection (Table() AS STRING)
DECLARE SUB PrintTable (Table() AS STRING, Selection AS SelectionStr)
DECLARE SUB PullDownMenus (Table() AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DECLARE SUB Quit ()
DECLARE SUB RemoveRow (Table() AS STRING, AtRow AS INTEGER, AtPage AS INTEGER, LossGainBlock AS LossGainBlockStr)
DECLARE SUB RemoveTable (FileName AS STRING, TableFileName AS STRING)
DECLARE SUB SaveTable (Table() AS STRING, FileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DECLARE SUB SetLossGainBlock (Table() AS STRING, Selection AS SelectionStr, Page AS INTEGER, Row AS INTEGER, LossGainBlock AS LossGainBlockStr, Income AS DOUBLE, Expenses AS DOUBLE)
DECLARE SUB Sort (Table() AS STRING, SortColumn AS INTEGER, LossGainBlock AS LossGainBlockStr)
DECLARE FUNCTION AddTableToList (FileName AS STRING) AS STRING
DECLARE FUNCTION Choose (x AS INTEGER, y AS INTEGER, Prompt AS STRING, Choices AS STRING) AS STRING
DECLARE FUNCTION FileExists (FileName AS STRING) AS INTEGER
DECLARE FUNCTION FileNamePrompt (FileName AS STRING) AS STRING
DECLARE FUNCTION FormatTableFileName (FileName AS STRING) AS STRING
DECLARE FUNCTION GetDate () AS STRING
DECLARE FUNCTION GetDescription () AS STRING
DECLARE FUNCTION GetInput (x AS INTEGER, y AS INTEGER, Prompt AS STRING, MaximumLength AS INTEGER, IsFilePrompt AS INTEGER, IsNumeric AS INTEGER, Default AS STRING) AS STRING
DECLARE FUNCTION GetTime () AS STRING
DECLARE FUNCTION InLossGainBlock (Row AS INTEGER, Column AS INTEGER, Page AS INTEGER, LossGainBlock AS LossGainBlockStr) AS INTEGER
DECLARE FUNCTION LoadTable (Table() AS STRING, FileName AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr) AS STRING
DECLARE FUNCTION Pad (Text AS STRING, Padding AS INTEGER, RightAligned AS INTEGER) AS STRING
DECLARE FUNCTION RowEmpty (Table() AS STRING, Row AS INTEGER, Page AS INTEGER) AS INTEGER
DECLARE FUNCTION SaveTableAs (Table() AS STRING, FileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr) AS STRING
DECLARE FUNCTION TableBaseFileName (FileName AS STRING) AS STRING
DECLARE FUNCTION TableHeader () AS STRING
DECLARE FUNCTION TableListed (FileName AS STRING) AS INTEGER

ON ERROR GOTO ErrorTrap
DIM SHARED IgnoreError AS INTEGER

 CALL Main
 END

MenuSizes:
 DATA 14,7,25,3,14,2

MenuItems:
 DATA "New File", "Load File", "Save File", "Save File As", "Remove File", "Print Table", "Quit"
 DATA "Calculate Loss/Gain", "Loss/Gain for Selection", "Print Selection"
 DATA "Display Help", "Information"

ErrorTrap:
DIM ErrorCode AS INTEGER

 ErrorCode = ERR
 PCOPY 0, 2: COLOR 0, 7
 Center 6, " " + ERROR$(ErrorCode) + ". "
 COLOR 7, 0: Center 0, " Spacebar = Ignore error  Enter = Retry  Escape = Quit  "
 IgnoreError = FALSE
 DO
  SELECT CASE INPUT$(1)
   CASE CHR$(13)
    PCOPY 2, 0: RESUME
   CASE IS = CHR$(27)
    COLOR 7, 0: CLS : SYSTEM
   CASE " "
    IgnoreError = TRUE: PCOPY 2, 0: RESUME NEXT
  END SELECT
 LOOP

FUNCTION AddTableToList (FileName AS STRING) AS STRING
DIM FileDate AS STRING
DIM FileH AS INTEGER
 
 FileH = FREEFILE
 OPEN "Names.lst" FOR APPEND LOCK READ WRITE AS FileH
  PRINT #FileH, TableBaseFileName(FileName); ","; GetTime(); ","; GetDate()
 CLOSE FileH

 AddTableToList = FileName
END FUNCTION

SUB CalculateLossOrGain (Table() AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DIM Expenses AS DOUBLE
DIM Income AS DOUBLE
DIM Page AS INTEGER
DIM Row AS INTEGER

 Income = 0
 Expenses = 0
 FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3)
  FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
   IF Row >= Selection.Row AND Page >= Selection.Page THEN EXIT FOR
   Income = Income + VAL(Table(Row, COL_INCOME, Page))
   Expenses = Expenses + VAL(Table(Row, COL_EXPENSE, Page))
  NEXT Row
  IF Row >= Selection.Row AND Page >= Selection.Page THEN EXIT FOR
 NEXT Page

 SetLossGainBlock Table(), Selection, Page, Row, LossGainBlock, Income, Expenses
END SUB

SUB CalculateLossOrGainSelection (Table() AS STRING)
DIM Description AS STRING
DIM Expenses AS DOUBLE
DIM Income AS DOUBLE
DIM IsLoss AS INTEGER
DIM LossGain AS STRING
DIM Page AS INTEGER
DIM Row AS INTEGER

 Description = GetDescription()
 IF NOT Description = "" THEN
  Description = LTRIM$(RTRIM$(UCASE$(Description)))

  Income = 0
  Expenses = 0
  FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3)
   FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
    IF LTRIM$(RTRIM$(UCASE$(Table(Row, COL_DESCRIPTION, Page)))) = Description THEN
     Income = Income + VAL(Table(Row, COL_INCOME, Page))
     Expenses = Expenses + VAL(Table(Row, COL_EXPENSE, Page))
    END IF
   NEXT Row
  NEXT Page

  IsLoss = (Income - Expenses < 0)
  LossGain = STR$(ABS(Income - Expenses))

  Description = LCASE$(Description)
  MID$(Description, 1, 1) = UCASE$(LEFT$(Description, 1))

  DrawBox 7, 19, 40, 4

  LOCATE 8, 21: PRINT "  Description: "; Description
  LOCATE , 21: PRINT "       Income: $ "; Pad$(FORMAT$(Income, NUMBER_MASK$), CELL_WIDTH, TRUE)
  LOCATE , 21: PRINT "     Expenses: $ "; Pad$(FORMAT$(Expenses, NUMBER_MASK$), CELL_WIDTH, TRUE)
  LOCATE , 21: PRINT SPACE$(5);
  IF IsLoss THEN PRINT "    Loss";  ELSE PRINT "    Gain";
  PRINT ": $ "; Pad$(FORMAT$(VAL(LossGain), NUMBER_MASK$), CELL_WIDTH, TRUE)

  DO WHILE INKEY$ = "": LOOP
 END IF
END SUB

SUB Center (x AS INTEGER, Text AS STRING)
 IF x > 0 THEN LOCATE x
 LOCATE , INT(40 - (LEN(Text) / 2))
 PRINT Text
END SUB

FUNCTION Choose (x AS INTEGER, y AS INTEGER, Prompt AS STRING, Choices AS STRING) AS STRING
DIM Choice AS STRING

 LOCATE x, y: PRINT Prompt
 DO
  Choice = UCASE$(INPUT$(1))
 LOOP UNTIL INSTR(Choices, Choice) > 0
 LOCATE x, y: PRINT SPACE$(LEN(Prompt))

 Choose = Choice
END FUNCTION

SUB ClearLossGainBlock (Table() AS STRING, LossGainBlock AS LossGainBlockStr)
DIM Row AS INTEGER

 IF NOT LossGainBlock.Row = NO_ROW AND NOT LossGainBlock.Page = NO_PAGE THEN
  FOR Row = LossGainBlock.Row TO LossGainBlock.Row + 3
   Table(Row, COL_INCOME, LossGainBlock.Page) = ""
   Table(Row, COL_EXPENSE, LossGainBlock.Page) = ""
  NEXT Row
  LossGainBlock.Row = NO_ROW
  LossGainBlock.Page = NO_PAGE
 END IF
END SUB

SUB CreateFile (FileName AS STRING)
DIM FileH AS INTEGER

 FileH = FREEFILE
 OPEN FileName FOR APPEND LOCK READ WRITE AS FileH
 CLOSE FileH
END SUB

SUB DisplayCell (Table() AS STRING, Row AS INTEGER, Column AS INTEGER, Page AS INTEGER)
DIM PadRight AS INTEGER

 PadRight = Column = COL_INCOME OR Column = COL_EXPENSE

 LOCATE Row + 2, (Column * CELL_WIDTH) + 1
 PRINT " "; Pad(Table(Row, Column, Page), MAXIMUM_DATA_LENGTH, PadRight); " ";
END SUB

SUB DisplayHelp ()
DIM FileH AS INTEGER
DIM HelpText AS STRING

 DrawBox 5, 16, 46, 13
 Center 5, "´ Help Ã"

 IgnoreError = FALSE
 FileH = FREEFILE
 OPEN "Bookkp2.hlp" FOR INPUT LOCK READ WRITE AS FileH
  DO UNTIL EOF(FileH)
   IF IgnoreError THEN EXIT DO
   LINE INPUT #FileH, HelpText
   LOCATE , 19: PRINT HelpText
  LOOP
 CLOSE FileH

 DO WHILE INKEY$ = "": LOOP
END SUB

SUB DisplayInformation ()
 DrawBox 5, 28, 22, 3

 Center 6, "Bookkeeping 2"
 Center 0, "By: Peter Swinkels"
 Center 0, "***1997***"

 DO WHILE INKEY$ = "": LOOP
END SUB

SUB DisplayPage (Table() AS STRING, Selection AS SelectionStr)
DIM Column AS INTEGER
DIM Row AS INTEGER

 COLOR 0, 7: LOCATE 1, 1: PRINT TableHeader()

 COLOR 7, 0
 FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
  FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
   DisplayCell Table(), Row, Column, Selection.Page
  NEXT Row
 NEXT Column
END SUB

SUB DisplayStatusBar (TableFileName AS STRING, Selection AS SelectionStr)
 COLOR 0, 7: LOCATE 23, 1

 PRINT USING "   Row: ##   Column: ##   Page: ###     Filename: "; Selection.Row + 1; Selection.Column + 1; Selection.Page + 1;

 IF TableFileName = "" THEN
  PRINT "None " + SPACE$(25); ""
 ELSE
  PRINT USING "\          \" + SPACE$(17); TableFileName;
 END IF
END SUB

SUB DisplayTableList ()
DIM FileH AS INTEGER
DIM KeyStroke AS STRING
DIM ListFile AS STRING
DIM Row AS INTEGER
DIM SaveDate AS STRING
DIM SaveTime AS STRING

 PCOPY 0, 1
 
 IF NOT FileExists("Names.lst") THEN CreateFile "Names.lst"

 IgnoreError = FALSE
 FileH = FREEFILE
 OPEN "Names.lst" FOR INPUT LOCK READ WRITE AS FileH
  IF LOF(FileH) = 0 THEN
   COLOR 7, 0: CLS
   PRINT "There are no saved files."
   DO: LOOP WHILE INKEY$ = ""
  ELSE
   DO UNTIL EOF(FileH) OR IgnoreError OR KeyStroke = CHR$(27)
    COLOR , 0: CLS
    COLOR 0, 7: PRINT "  File Name:  Time:    Date:"; SPACE$(52)
    COLOR 7, 0
    FOR Row = 1 TO 21
     IF EOF(FileH) THEN EXIT FOR
     INPUT #FileH, ListFile, SaveTime, SaveDate
     PRINT USING " \        \ \        \ \        \"; ListFile; SaveTime; SaveDate
    NEXT Row
    COLOR 0, 7: LOCATE 23, 2: PRINT " Press any key to continue... "
    KeyStroke = INPUT$(1)
   LOOP
  END IF
 CLOSE FileH
 COLOR , 0: CLS : PCOPY 1, 0
END SUB

SUB DrawBox (x AS INTEGER, y AS INTEGER, BoxWidth AS INTEGER, BoxHeight AS INTEGER)
DIM BoxRow AS INTEGER

 COLOR 0, 7
 LOCATE x, y: PRINT "Ú"; STRING$(BoxWidth, "Ä"); "¿"
 FOR BoxRow = 1 TO BoxHeight
  LOCATE , y: PRINT "³"; SPACE$(BoxWidth); "³°°"
 NEXT BoxRow
 LOCATE , y: PRINT "À"; STRING$(BoxWidth, "Ä"); "Ù°°"
 LOCATE , y + 2: PRINT STRING$(BoxWidth + 2, "°")
END SUB

SUB Export (Table() AS STRING)
DIM Column AS INTEGER
DIM FileH AS INTEGER
DIM Page AS INTEGER
DIM Row AS INTEGER

 IgnoreError = FALSE
 FileH = FREEFILE
 OPEN "Export.txt" FOR OUTPUT LOCK READ WRITE AS FileH
  FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3)
   IF IgnoreError THEN EXIT FOR
    FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
    IF IgnoreError THEN EXIT FOR
    FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
     IF IgnoreError THEN EXIT FOR
     PRINT #FileH, Table(Row, Column, Page);
     IF NOT Column = UBOUND(Table, 2) THEN PRINT #FileH, CHR$(9);
    NEXT Column
    PRINT #FileH,
   NEXT Row
  NEXT Page
 CLOSE FileH
END SUB

FUNCTION FileExists (FileName AS STRING) AS INTEGER
 FileExists = NOT (DIR$(FileName) = "")
END FUNCTION

FUNCTION FileNamePrompt (FileName AS STRING) AS STRING
 DrawBox 5, 28, 22, 1
 FileNamePrompt = GetInput(6, 29, " File Name: ", 8, TRUE, FALSE, FileName)
END FUNCTION

SUB FormatCell (Table() AS STRING, Previous AS SelectionStr, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
 IF NOT InLossGainBlock(Previous.Row, Previous.Column, Previous.Page, LossGainBlock) THEN
  IF NOT (Previous.Row = Selection.Row AND Previous.Column = Selection.Column AND Previous.Page = Selection.Page) THEN
   IF NOT Table(Previous.Row, Previous.Column, Previous.Page) = "" THEN
    IF Previous.Column = COL_INCOME OR Previous.Column = COL_EXPENSE THEN
     IF Selection.Page = Previous.Page THEN
      Table(Previous.Row, Previous.Column, Previous.Page) = FORMAT$(VAL(Table(Previous.Row, Previous.Column, Previous.Page)), NUMBER_MASK$)
      DisplayCell Table(), Previous.Row, Previous.Column, Previous.Page
     END IF
    END IF
   END IF
  END IF
 END IF
END SUB

FUNCTION FormatTableFileName (FileName AS STRING) AS STRING
DIM TableFileName AS STRING

 TableFileName = LTRIM$(RTRIM$(LCASE$(FileName)))
 MID$(TableFileName, 1, 1) = UCASE$(LEFT$(TableBaseFileName(TableFileName), 1))

 FormatTableFileName = TableFileName + ".bk2"
END FUNCTION

FUNCTION GetDate () AS STRING
DIM CurrentDate AS SINGLE

 CurrentDate = DATEVALUE(DATE$)

 GetDate = LTRIM$(STR$(DAY(CurrentDate))) + "-" + LTRIM$(STR$(MONTH(CurrentDate))) + "-" + LTRIM$(STR$(YEAR(CurrentDate)))
END FUNCTION

FUNCTION GetDescription () AS STRING
 DrawBox 7, 22, 34, 1
 GetDescription = GetInput(8, 23, " Description: ", MAXIMUM_DATA_LENGTH, FALSE, FALSE, "")
END FUNCTION

FUNCTION GetInput (x AS INTEGER, y AS INTEGER, Prompt AS STRING, MaximumLength AS INTEGER, IsFilePrompt AS INTEGER, IsNumeric AS INTEGER, Default AS STRING) AS STRING
DIM KeyStroke AS STRING
DIM Text AS STRING

 Text = Default
 
 COLOR , 0: LOCATE x, y: PRINT SPACE$(LEN(Prompt) + MaximumLength + 2);

 DO
  COLOR 7: LOCATE x, y: PRINT Prompt; Text; : COLOR 23: PRINT "_ "
  DO
   KeyStroke = INKEY$
  LOOP WHILE KeyStroke = ""
  SELECT CASE KeyStroke
   CASE CHR$(8)
    IF LEN(Text) > 0 THEN Text = LEFT$(Text, LEN(Text) - 1)
   CASE CHR$(13)
    EXIT DO
   CASE CHR$(27)
    Text = ""
    EXIT DO
   CASE CHR$(0) + ";"
    IF IsFilePrompt THEN DisplayTableList
   CASE ELSE
    IF KeyStroke >= " " AND LEN(Text) < MaximumLength THEN
     IF IsFilePrompt THEN
      IF INSTR(FILE_NAME_FILTER$, KeyStroke) = 0 THEN Text = Text + KeyStroke
     ELSEIF IsNumeric THEN
      IF INSTR("0123456789", KeyStroke) > 0 THEN Text = Text + KeyStroke
     ELSE
      Text = Text + KeyStroke
     END IF
    END IF
  END SELECT
 LOOP
 LOCATE x, y: PRINT SPACE$(LEN(Prompt) + LEN(Text) + 2)

 GetInput = Text
END FUNCTION

FUNCTION GetTime () AS STRING
DIM CurrentTime AS SINGLE

 CurrentTime = TIMEVALUE(TIME$)

 GetTime = FORMAT$(HOUR(CurrentTime), "##") + ":" + FORMAT$(MINUTE(CurrentTime), "00") + ":" + FORMAT$(SECOND(CurrentTime), "00")
END FUNCTION

SUB Import (Table() AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DIM AbsoluteLossGainRow AS STRING
DIM Column AS INTEGER
DIM Delimiter AS STRING
DIM FileH AS INTEGER
DIM ImportData AS STRING
DIM Length AS INTEGER
DIM Page AS INTEGER
DIM Position AS INTEGER
DIM Row AS INTEGER

 IF FileExists("Import.txt") THEN
  FileH = FREEFILE
  OPEN "Import.txt" FOR BINARY LOCK READ WRITE AS FileH
   Initialize Table(), TableFileName, Selection, LossGainBlock
   ImportData = INPUT$(LOF(FileH), FileH)
  CLOSE FileH

  Position = INSTR(ImportData, CHR$(10))
  DO WHILE Position > 0
   ImportData = LEFT$(ImportData, Position - 1) + MID$(ImportData, Position + 1)
   Position = INSTR(ImportData, CHR$(10))
  LOOP
 
  DrawBox 7, 19, 40, 1
  AbsoluteLossGainRow = GetInput(8, 21, " Loss/Gain row: ", 5, FALSE, TRUE, "")
  IF NOT AbsoluteLossGainRow = "" THEN
   LossGainBlock.Page = VAL(AbsoluteLossGainRow) \ ROWS_PER_PAGE
   LossGainBlock.Row = VAL(AbsoluteLossGainRow) - (LossGainBlock.Page * ROWS_PER_PAGE)
  END IF

  FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3)
   IF ImportData = "" THEN EXIT FOR
   FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
    IF ImportData = "" THEN EXIT FOR
    FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
     IF ImportData = "" THEN EXIT FOR
     IF Column = UBOUND(Table, 2) THEN Delimiter = CHR$(13) ELSE Delimiter = CHR$(9)
     Length = INSTR(ImportData, Delimiter)
     IF Length = 0 THEN Length = LEN(ImportData)
     Table(Row, Column, Page) = LEFT$(ImportData, Length - 1)
     IF Length > MAXIMUM_DATA_LENGTH THEN Table(Row, Column, Page) = LEFT$(Table(Row, Column, Page), MAXIMUM_DATA_LENGTH)
     IF Column = COL_INCOME OR Column = COL_EXPENSE THEN
      IF NOT InLossGainBlock(Row, Column, Page, LossGainBlock) THEN
       IF NOT Table(Row, Column, Page) = "" THEN
        Table(Row, Column, Page) = FORMAT$(VAL(Table(Row, Column, Page)), NUMBER_MASK$)
       END IF
      END IF
     END IF
     ImportData = MID$(ImportData, Length + 1)
    NEXT Column
   NEXT Row
  NEXT Page
 ELSE
  ERROR 53
 END IF
END SUB

SUB Initialize (Table() AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DIM FileName AS STRING

 SCREEN 0: WIDTH 80, 25: COLOR 7, 0: CLS

 REDIM Table(0 TO 20, 0 TO 3, 0 TO 20) AS STRING

 IgnoreError = FALSE
 LossGainBlock.Page = NO_PAGE
 LossGainBlock.Row = NO_ROW
 Selection.Column = LBOUND(Table, 2)
 Selection.Page = LBOUND(Table, 3)
 Selection.Row = LBOUND(Table, 1)
 TableFileName = ""

 IF NOT COMMAND$ = "" THEN
  FileName = COMMAND$
  IF INSTR(FileName, ".") > 0 THEN FileName = LEFT$(FileName, INSTR(FileName, ".") - 1)
  TableFileName = LoadTable(Table(), FileName, TableFileName, Selection, LossGainBlock)
 END IF
END SUB

FUNCTION InLossGainBlock (Row AS INTEGER, Column AS INTEGER, Page AS INTEGER, LossGainBlock AS LossGainBlockStr) AS INTEGER
DIM InBlock AS INTEGER

 InBlock = FALSE

 IF Page = LossGainBlock.Page THEN
  IF Column = COL_INCOME OR Column = COL_EXPENSE THEN
   IF Row >= LossGainBlock.Row AND Row <= LossGainBlock.Row + 4 THEN
    InBlock = TRUE
   END IF
  END IF
 END IF

 InLossGainBlock = InBlock
END FUNCTION

SUB InsertRow (Table() AS STRING, AtRow AS INTEGER, AtPage AS INTEGER, LossGainBlock AS LossGainBlockStr)
DIM EndRow AS INTEGER
DIM Page AS INTEGER
DIM Row AS INTEGER

 IF NOT (LossGainBlock.Row = NO_PAGE OR LossGainBlock.Page = NO_PAGE) THEN
  IF LossGainBlock.Page > AtPage THEN
   IF LossGainBlock.Row = UBOUND(Table, 1) THEN
    LossGainBlock.Page = LossGainBlock.Page + 1
    LossGainBlock.Row = LBOUND(Table, 1)
   END IF
  ELSEIF LossGainBlock.Row > AtRow THEN
   LossGainBlock.Row = LossGainBlock.Row + 1
  END IF
 END IF

 FOR Page = UBOUND(Table, 3) - 1 TO AtPage STEP -1
  IF Page = AtPage THEN EndRow = AtRow ELSE EndRow = LBOUND(Table, 1)
  FOR Row = UBOUND(Table, 1) TO EndRow STEP -1
   IF Row = UBOUND(Table, 1) THEN
    MoveRow Table(), Row, Page, LBOUND(Table, 1), Page + 1
   ELSE
    MoveRow Table(), Row, Page, Row + 1, Page
   END IF
  NEXT Row
 NEXT Page
END SUB

FUNCTION LoadTable (Table() AS STRING, FileName AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr) AS STRING
DIM Column AS INTEGER
DIM FileH AS INTEGER
DIM Length AS INTEGER
DIM Page AS INTEGER
DIM Row AS INTEGER

 IF NOT FileName = "" THEN
  FileName = FormatTableFileName(FileName)
  IF FileExists(FileName) THEN
   IgnoreError = FALSE
   FileH = FREEFILE
   OPEN FileName FOR BINARY LOCK READ WRITE AS FileH
    Initialize Table(), TableFileName, Selection, LossGainBlock
    Selection.Row = ASC(INPUT$(1, FileH))
    Selection.Column = ASC(INPUT$(1, FileH))
    Selection.Page = ASC(INPUT$(1, FileH))
    LossGainBlock.Row = ASC(INPUT$(1, FileH))
    LossGainBlock.Page = ASC(INPUT$(1, FileH))
    FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3)
     IF IgnoreError THEN EXIT FOR
      FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
      IF IgnoreError THEN EXIT FOR
      FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
       IF IgnoreError THEN EXIT FOR
       Length = ASC(INPUT$(1, FileH)): Table(Row, Column, Page) = INPUT$(Length, FileH)
      NEXT Column
     NEXT Row
    NEXT Page
   CLOSE FileH
  ELSE
   FileName = TableFileName
   ERROR 53
  END IF
 END IF

 LoadTable = FileName
END FUNCTION

SUB Main ()
DIM KeyStroke AS STRING
DIM Length AS INTEGER
DIM LossGainBlock AS LossGainBlockStr
DIM Previous AS SelectionStr
DIM Selection AS SelectionStr
DIM Table(0 TO 0, 0 TO 0, 0 TO 0) AS STRING
DIM TableFileName AS STRING

 Initialize Table(), TableFileName, Selection, LossGainBlock

 DisplayPage Table(), Selection
 DO
  DisplayStatusBar TableFileName, Selection

  Length = LEN(Table(Selection.Row, Selection.Column, Selection.Page))
  DisplayCell Table(), Selection.Row, Selection.Column, Selection.Page
  DO
   KeyStroke = INKEY$
  LOOP WHILE KeyStroke = ""
  COLOR 7, 0: DisplayCell Table(), Selection.Row, Selection.Column, Selection.Page

  Previous = Selection

  SELECT CASE KeyStroke
   CASE CHR$(0) + "H"
    IF Selection.Row = LBOUND(Table, 1) THEN
     IF Selection.Page > LBOUND(Table, 1) THEN
      Selection.Page = Selection.Page - 1
      Selection.Row = UBOUND(Table, 1)
      DisplayPage Table(), Selection
     END IF
    ELSE
     Selection.Row = Selection.Row - 1
    END IF
   CASE CHR$(0) + "P"
    IF Selection.Row = UBOUND(Table, 1) THEN
     IF Selection.Page < UBOUND(Table, 1) THEN
      Selection.Page = Selection.Page + 1
      Selection.Row = LBOUND(Table, 1)
      DisplayPage Table(), Selection
     END IF
    ELSE
     Selection.Row = Selection.Row + 1
    END IF
   CASE CHR$(0) + "K"
    IF Selection.Column > LBOUND(Table, 2) THEN
     Selection.Column = Selection.Column - 1
    ELSE
     Selection.Column = UBOUND(Table, 2)
    END IF
   CASE CHR$(0) + "M"
    IF Selection.Column < UBOUND(Table, 2) THEN
     Selection.Column = Selection.Column + 1
    ELSE
     Selection.Column = LBOUND(Table, 2)
    END IF
   CASE CHR$(0) + "I"
    IF Selection.Page > LBOUND(Table, 3) THEN
     Selection.Page = Selection.Page - 1
     Selection.Row = UBOUND(Table, 2)
     DisplayPage Table(), Selection
    END IF
   CASE CHR$(0) + "Q"
    IF Selection.Page < UBOUND(Table, 3) THEN
     Selection.Page = Selection.Page + 1
     Selection.Row = LBOUND(Table, 2)
     DisplayPage Table(), Selection
    END IF
   CASE CHR$(0) + "R"
    IF Selection.Page < UBOUND(Table, 3) THEN
     IF Length = 0 THEN Table(Selection.Row, Selection.Column, Selection.Page) = ClipBoard.GETTEXT()
    END IF
   CASE CHR$(0) + "S"
    IF InLossGainBlock(Selection.Row, Selection.Column, Selection.Page, LossGainBlock) THEN
     ClearLossGainBlock Table(), LossGainBlock
     DisplayPage Table(), Selection
    ELSE
     IF Length > 0 THEN
      ClipBoard.SETTEXT Table(Selection.Row, Selection.Column, Selection.Page)
      Table(Selection.Row, Selection.Column, Selection.Page) = ""
     END IF
    END IF
   CASE CHR$(0) + ";"
    DisplayTableList
   CASE CHR$(0) + "<"
    DrawBox 8, 32, 14, 1
    IF Choose(9, 34, "Export Y/N?", "YN") = "Y" THEN Export Table()
    DisplayPage Table(), Selection
   CASE CHR$(0) + "="
    DrawBox 8, 32, 14, 1
    IF Choose(9, 34, "Import Y/N?", "YN") = "Y" THEN
     Import Table(), TableFileName, Selection, LossGainBlock
     TableFileName = ""
    END IF
    DisplayPage Table(), Selection
   CASE CHR$(3)
    IF Length > 0 THEN ClipBoard.SETTEXT Table(Selection.Row, Selection.Column, Selection.Page)
   CASE CHR$(4)
    RemoveRow Table(), Selection.Row, Selection.Page, LossGainBlock
    DisplayPage Table(), Selection
   CASE CHR$(8)
    IF NOT InLossGainBlock(Selection.Row, Selection.Column, Selection.Page, LossGainBlock) THEN
     IF Length > 0 THEN Table(Selection.Row, Selection.Column, Selection.Page) = LEFT$(Table(Selection.Row, Selection.Column, Selection.Page), Length - 1)
    END IF
   CASE CHR$(9)
    IF RowEmpty(Table(), UBOUND(Table, 1), UBOUND(Table, 3) - 1) THEN
     InsertRow Table(), Selection.Row, Selection.Page, LossGainBlock
     DisplayPage Table(), Selection
    ELSE
     DrawBox 5, 27, 24, 1
     Center 6, "Last row is not empty."
     DO WHILE INKEY$ = "": LOOP
     DisplayPage Table(), Selection
    END IF
   CASE CHR$(13)
    IF Selection.Column < UBOUND(Table, 2) THEN
     Selection.Column = Selection.Column + 1
    ELSE
     IF Selection.Row = UBOUND(Table, 1) THEN
      IF Selection.Page < UBOUND(Table, 3) THEN
       Selection.Page = Selection.Page + 1
       Selection.Row = LBOUND(Table, 1)
       DisplayPage Table(), Selection
      END IF
     ELSE
      Selection.Row = Selection.Row + 1
      Selection.Column = LBOUND(Table, 2)
     END IF
    END IF
   CASE CHR$(19)
    Sort Table(), Selection.Column, LossGainBlock
    DisplayPage Table(), Selection
   CASE CHR$(27)
    PullDownMenus Table(), TableFileName, Selection, LossGainBlock
    DisplayPage Table(), Selection
   CASE ELSE
    IF Selection.Page < UBOUND(Table, 3) THEN
     IF NOT InLossGainBlock(Selection.Row, Selection.Column, Selection.Page, LossGainBlock) THEN
      IF Selection.Column = COL_DATE OR Selection.Column = COL_DESCRIPTION THEN
       IF KeyStroke >= " " AND Length < MAXIMUM_DATA_LENGTH THEN Table(Selection.Row, Selection.Column, Selection.Page) = Table(Selection.Row, Selection.Column, Selection.Page) + KeyStroke
      ELSE
       IF KeyStroke = "," THEN KeyStroke = "."
       IF INSTR("0123456789.", KeyStroke) > 0 THEN
        IF INSTR(Table(Selection.Row, Selection.Column, Selection.Page), ".") > 0 THEN
         IF NOT KeyStroke = "." AND Length < MAXIMUM_NUMBER_LENGTH THEN Table(Selection.Row, Selection.Column, Selection.Page) = Table(Selection.Row, Selection.Column, Selection.Page) + KeyStroke
        ELSE
         IF Length < MAXIMUM_WHOLE_NR_LENGTH OR (KeyStroke = "." AND Length < MAXIMUM_WHOLE_NR_LENGTH + 1) THEN Table(Selection.Row, Selection.Column, Selection.Page) = Table(Selection.Row, Selection.Column, Selection.Page) + KeyStroke
        END IF
       END IF
      END IF
     END IF
    ELSE
     DrawBox 5, 27, 24, 1
     Center 6, "This page is reserved."
     DO WHILE INKEY$ = "": LOOP
     DisplayPage Table(), Selection
    END IF
   END SELECT

   FormatCell Table(), Previous, Selection, LossGainBlock
 LOOP
END SUB

SUB MoveRow (Table() AS STRING, SourceRow AS INTEGER, SourcePage AS INTEGER, TargetRow AS INTEGER, TargetPage AS INTEGER)
DIM Column AS INTEGER

 FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
  Table(TargetRow, Column, TargetPage) = Table(SourceRow, Column, SourcePage)
  Table(SourceRow, Column, SourcePage) = ""
 NEXT Column
END SUB

FUNCTION Pad (Text AS STRING, Padding AS INTEGER, RightAligned AS INTEGER) AS STRING
DIM Padded AS STRING

 IF RightAligned THEN
  Padded = SPACE$(Padding - LEN(Text)) + Text
 ELSE
  Padded = Text + SPACE$(Padding - LEN(Text))
 END IF

 Pad = Padded
END FUNCTION

SUB PrintSelection (Table() AS STRING)
DIM Column AS INTEGER
DIM Description AS STRING
DIM Expenses AS DOUBLE
DIM Income AS DOUBLE
DIM IsLoss AS INTEGER
DIM LossGain AS STRING
DIM Page AS INTEGER
DIM Row AS INTEGER

 Description = GetDescription()
 IF NOT Description = "" THEN
  Description = LTRIM$(RTRIM$(UCASE$(Description)))
  Income = 0
  Expenses = 0
  FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3) - 1
   FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
    IF LTRIM$(RTRIM$(UCASE$(Table(Row, COL_DESCRIPTION, Page)))) = Description THEN
     FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
      IF Column = COL_INCOME OR Column = COL_EXPENSE THEN
       LPRINT " "; Pad(Table(Row, Column, Page), MAXIMUM_DATA_LENGTH, TRUE); " ";
      ELSE
       LPRINT " "; Pad(Table(Row, Column, Page), MAXIMUM_DATA_LENGTH, FALSE); " ";
      END IF
     NEXT Column
     Income = Income + VAL(Table(Row, COL_INCOME, Page))
     Expenses = Expenses + VAL(Table(Row, COL_EXPENSE, Page))
    END IF
   NEXT Row
  NEXT Page

  IsLoss = (Income - Expenses < 0)
  LossGain = STR$(ABS(Income - Expenses))
  
  LPRINT
  LPRINT "    Income: $ "; Pad$(FORMAT$(Income, NUMBER_MASK$), 20, TRUE)
  LPRINT "  Expenses: $ "; Pad$(FORMAT$(Expenses, NUMBER_MASK$), 20, TRUE)
  IF IsLoss THEN LPRINT "      Loss:";  ELSE LPRINT "      Gain:";
  LPRINT " $ "; Pad$(FORMAT$(VAL(LossGain), NUMBER_MASK$), 20, TRUE)
 END IF
END SUB

SUB PrintTable (Table() AS STRING, Selection AS SelectionStr)
DIM Column AS INTEGER
DIM EndPage AS INTEGER
DIM Page AS INTEGER
DIM PageNumber AS STRING
DIM Row AS INTEGER
DIM StartPage AS INTEGER

 DrawBox 8, 18, 42, 1

 EndPage = NO_PAGE
 StartPage = NO_PAGE
 
 SELECT CASE Choose(9, 20, "1 = Print all pages.  2 = Print a page.", "12" + CHR$(27))
  CASE "1"
   EndPage = UBOUND(Table, 3)
   StartPage = LBOUND(Table, 3)
  CASE "2"
   DisplayPage Table(), Selection
   DrawBox 8, 30, 18, 1
   PageNumber = GetInput(9, 31, " Page Number: ", 2, FALSE, TRUE, "")
   StartPage = VAL(PageNumber) - 1
   EndPage = StartPage
 END SELECT

 IF NOT (StartPage = NO_PAGE OR EndPage = NO_PAGE) THEN
  IF StartPage >= LBOUND(Table, 3) AND EndPage <= UBOUND(Table, 3) THEN
   LPRINT TableHeader()
   FOR Page = StartPage TO EndPage
    FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
     IF NOT (Table(Row, COL_DATE, Page) = "" AND Table(Row, COL_DESCRIPTION, Page) = "" AND Table(Row, COL_INCOME, Page) = "" AND Table(Row, COL_EXPENSE, Page) = "") THEN
      FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
       IF Column = COL_INCOME OR Column = COL_EXPENSE THEN
        LPRINT " "; Pad(Table(Row, Column, Page), MAXIMUM_DATA_LENGTH, TRUE); " ";
       ELSE
        LPRINT " "; Pad(Table(Row, Column, Page), MAXIMUM_DATA_LENGTH, FALSE); " ";
       END IF
      NEXT Column
     END IF
    NEXT Row
   NEXT Page
  END IF
 END IF
END SUB

SUB PullDownMenus (Table() AS STRING, TableFileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DIM FileName AS STRING
DIM KeyStroke AS STRING
DIM Menu AS INTEGER
DIM MenuBar AS STRING
DIM MenuHeight(0 TO 2)
DIM MenuItem AS INTEGER
DIM MenuText(0 TO 2, 0 TO 6) AS STRING
DIM MenuWidth(0 TO 2)
DIM MenuY AS INTEGER
DIM SelectedMenu AS INTEGER
DIM SelectedMenuItem AS INTEGER
 
 MenuBar = Pad("  File" + SPACE$(6) + "Other" + SPACE$(5) + "Help ", 80, FALSE)
 SelectedMenu = 0
 SelectedMenuItem = 0

 RESTORE MenuSizes
 FOR Menu = LBOUND(MenuText, 1) TO UBOUND(MenuText, 1)
  READ MenuWidth(Menu), MenuHeight(Menu)
 NEXT Menu

 RESTORE MenuItems
 FOR Menu = LBOUND(MenuText, 1) TO UBOUND(MenuText, 1)
  FOR MenuItem = 0 TO MenuHeight(Menu) - 1
   READ MenuText(Menu, MenuItem)
   MenuText(Menu, MenuItem) = " " + MenuText(Menu, MenuItem) + SPACE$((MenuWidth(Menu) - LEN(MenuText(Menu, MenuItem))) - 1)
  NEXT MenuItem
 NEXT Menu

 PCOPY 0, 1
 DO
  PCOPY 1, 0
  DO
   MenuY = (SelectedMenu * 10) + 1
   DrawBox 2, MenuY, MenuWidth(SelectedMenu), MenuHeight(SelectedMenu)
   LOCATE 1, 1: PRINT MenuBar;
   COLOR 7, 0
   LOCATE 1, MenuY: PRINT MID$(MenuBar, MenuY, 10);
   COLOR 0, 7
   FOR MenuItem = 0 TO MenuHeight(SelectedMenu) - 1
    IF SelectedMenuItem = MenuItem THEN COLOR 7, 0 ELSE COLOR 0, 7
    LOCATE MenuItem + 3, MenuY + 1: PRINT MenuText(SelectedMenu, MenuItem)
   NEXT MenuItem
   DO
    KeyStroke = INKEY$
   LOOP WHILE KeyStroke = ""
   IF KeyStroke = CHR$(0) + "K" OR KeyStroke = CHR$(0) + "M" THEN
    PCOPY 1, 0
    SelectedMenuItem = 0
   END IF
   SELECT CASE KeyStroke
    CASE CHR$(0) + "H"
     IF SelectedMenuItem > 0 THEN
      SelectedMenuItem = SelectedMenuItem - 1
     ELSE
      SelectedMenuItem = MenuHeight(SelectedMenu) - 1
     END IF
    CASE CHR$(0) + "P"
     IF SelectedMenuItem < MenuHeight(SelectedMenu) - 1 THEN
      SelectedMenuItem = SelectedMenuItem + 1
     ELSE
      SelectedMenuItem = 0
     END IF
    CASE CHR$(0) + "K"
     IF SelectedMenu = LBOUND(MenuText, 1) THEN
      SelectedMenu = UBOUND(MenuText, 1)
     ELSE
      MenuY = MenuY - 10
      SelectedMenu = SelectedMenu - 1
     END IF
    CASE CHR$(0) + "M"
     IF SelectedMenu = UBOUND(MenuText, 1) THEN
      SelectedMenu = LBOUND(MenuText, 1)
     ELSE
      SelectedMenu = SelectedMenu + 1
     END IF
    CASE CHR$(13)
     SELECT CASE SelectedMenu
      CASE 0
       SELECT CASE SelectedMenuItem
        CASE 0
         DrawBox 8, 25, 28, 1
         IF Choose(9, 27, "Start with new table Y/N?", "YN") = "Y" THEN
          Initialize Table(), TableFileName, Selection, LossGainBlock
          DisplayPage Table(), Selection
          DisplayStatusBar TableFileName, Selection
          PCOPY 0, 1
         END IF
        CASE 1
         FileName = LoadTable(Table(), FileNamePrompt(""), TableFileName, Selection, LossGainBlock)
         IF NOT FileName = "" THEN TableFileName = FileName
         DisplayPage Table(), Selection
         DisplayStatusBar TableFileName, Selection
         PCOPY 0, 1
        CASE 2
         IF TableFileName = "" THEN
          FileName = SaveTableAs(Table(), FileNamePrompt(TableBaseFileName(TableFileName)), Selection, LossGainBlock)
          IF NOT FileName = "" THEN TableFileName = FileName
          DisplayStatusBar TableFileName, Selection
         ELSE
          SaveTable Table(), TableFileName, Selection, LossGainBlock
          DisplayStatusBar TableFileName, Selection
         END IF
        CASE 3
         FileName = SaveTableAs(Table(), FileNamePrompt(TableBaseFileName(TableFileName)), Selection, LossGainBlock)
         IF NOT FileName = "" THEN TableFileName = FileName
         DisplayStatusBar TableFileName, Selection
        CASE 4
         RemoveTable FileNamePrompt(TableBaseFileName(TableFileName)), TableFileName
        CASE 5
         PrintTable Table(), Selection
        CASE 6
         Quit
       END SELECT
      CASE 1
       SELECT CASE SelectedMenuItem
        CASE 0
         ClearLossGainBlock Table(), LossGainBlock
         CalculateLossOrGain Table(), Selection, LossGainBlock
         DisplayPage Table(), Selection
         PCOPY 0, 1
        CASE 1
         CalculateLossOrGainSelection Table()
        CASE 2
         PrintSelection Table()
       END SELECT
      CASE 2
       SELECT CASE SelectedMenuItem
        CASE 0
         DisplayHelp
        CASE 1
         DisplayInformation
       END SELECT
     END SELECT
     EXIT DO
    CASE CHR$(27)
     EXIT DO
   END SELECT
  LOOP
 LOOP UNTIL KeyStroke = CHR$(27)
END SUB

SUB Quit ()
 DrawBox 8, 33, 12, 1
 IF Choose(9, 35, "Quit Y/N?", "YN") = "Y" THEN
  SCREEN 0: WIDTH 80, 25
  COLOR 7, 0: CLS
  SYSTEM
 END IF
END SUB

SUB RemoveRow (Table() AS STRING, AtRow AS INTEGER, AtPage AS INTEGER, LossGainBlock AS LossGainBlockStr)
DIM Page AS INTEGER
DIM Row AS INTEGER
DIM StartRow AS INTEGER

 IF InLossGainBlock(AtRow, COL_INCOME, AtPage, LossGainBlock) THEN ClearLossGainBlock Table(), LossGainBlock

 IF NOT (LossGainBlock.Row = NO_PAGE OR LossGainBlock.Page = NO_PAGE) THEN
  IF LossGainBlock.Page > AtPage THEN
   IF LossGainBlock.Row = LBOUND(Table, 1) THEN
    LossGainBlock.Page = LossGainBlock.Page - 1
    LossGainBlock.Row = UBOUND(Table, 1)
   END IF
  ELSEIF LossGainBlock.Row > AtRow THEN
   LossGainBlock.Row = LossGainBlock.Row - 1
  END IF
 END IF

 FOR Page = AtPage TO UBOUND(Table, 3) - 1
  IF Page = AtPage THEN StartRow = AtRow ELSE StartRow = LBOUND(Table, 1)
  FOR Row = StartRow TO UBOUND(Table, 1)
   IF Row = UBOUND(Table, 1) THEN
    MoveRow Table(), LBOUND(Table, 1), Page + 1, Row, Page
   ELSE
    MoveRow Table(), Row + 1, Page, Row, Page
   END IF
  NEXT Row
 NEXT Page
END SUB

SUB RemoveTable (FileName AS STRING, TableFileName AS STRING)
DIM DeletedFile AS STRING
DIM InFileH AS INTEGER
DIM OutFileH AS INTEGER
DIM SaveDate AS STRING
DIM SaveTime AS STRING

 IF NOT (FileName = "" OR INSTR(FileName, "*") > 0) THEN
  DeletedFile = FormatTableFileName(FileName)
  DrawBox 5, 28, 24, 1
  IF Choose(6, 30, "Remove " + CHR$(34) + TableBaseFileName(DeletedFile) + CHR$(34) + " Y/N?", "YN") = "Y" THEN
   IF NOT FileExists("Names.lst") THEN CreateFile "Names.lst"

   KILL DeletedFile
   IgnoreError = FALSE
   InFileH = FREEFILE
   OPEN "Names.lst" FOR INPUT LOCK READ WRITE AS InFileH
    OutFileH = FREEFILE
    OPEN "Names.tmp" FOR OUTPUT LOCK READ WRITE AS OutFileH
     DO UNTIL EOF(InFileH) OR IgnoreError
      INPUT #InFileH, TableFileName, SaveTime, SaveDate
      IF NOT TableBaseFileName$(DeletedFile) = TableFileName THEN PRINT #OutFileH, DeletedFile; ","; SaveTime; ","; SaveDate
     LOOP
    CLOSE OutFileH
   CLOSE InFileH
   KILL "Names.lst"
   NAME "Names.tmp" AS "Names.lst"
  END IF
 END IF
END SUB

FUNCTION RowEmpty (Table() AS STRING, Row AS INTEGER, Page AS INTEGER) AS INTEGER
DIM Column AS INTEGER
DIM Empty AS INTEGER

 Empty = TRUE
 FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
  IF NOT Table(Row, Column, Page) = "" THEN
   Empty = FALSE
   EXIT FOR
  END IF
 NEXT Column

 RowEmpty = Empty
END FUNCTION

SUB SaveTable (Table() AS STRING, FileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr)
DIM Column AS INTEGER
DIM FileH AS INTEGER
DIM Page AS INTEGER
DIM Row AS INTEGER

 IF NOT FileName = "" THEN
  IgnoreError = FALSE
  FileName = FormatTableFileName(TableBaseFileName(FileName))
  FileH = FREEFILE
  OPEN FileName FOR OUTPUT LOCK READ WRITE AS FileH
   PRINT #FileH, CHR$(Selection.Row);
   PRINT #FileH, CHR$(Selection.Column);
   PRINT #FileH, CHR$(Selection.Page);
   PRINT #FileH, CHR$(LossGainBlock.Row);
   PRINT #FileH, CHR$(LossGainBlock.Page);
   FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3)
    IF IgnoreError THEN EXIT FOR
     FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
     IF IgnoreError THEN EXIT FOR
      FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
       IF IgnoreError THEN EXIT FOR
       PRINT #FileH, CHR$(LEN(Table(Row, Column, Page))); Table(Row, Column, Page);
      NEXT Column
    NEXT Row
   NEXT Page
  CLOSE FileH
 END IF
END SUB

FUNCTION SaveTableAs (Table() AS STRING, FileName AS STRING, Selection AS SelectionStr, LossGainBlock AS LossGainBlockStr) AS STRING
 IF NOT FileName = "" THEN
  IF TableListed(FileName) THEN
   DrawBox 5, 20, 38, 1
   IF Choose(6, 22, "File already exists, overwrite Y/N?", "YN") = "Y" THEN
    SaveTable Table(), FileName, Selection, LossGainBlock
   END IF
  ELSE
   SaveTable Table(), FileName, Selection, LossGainBlock
   FileName = AddTableToList(FileName)
  END IF
 END IF

 SaveTableAs = FileName
END FUNCTION

SUB SetLossGainBlock (Table() AS STRING, Selection AS SelectionStr, Page AS INTEGER, Row AS INTEGER, LossGainBlock AS LossGainBlockStr, Income AS DOUBLE, Expenses AS DOUBLE)
DIM Column AS INTEGER
DIM IsLoss AS INTEGER
DIM LossGainResult AS DOUBLE

 LossGainResult = Income - Expenses
 IsLoss = (LossGainResult < 0)
 IF IsLoss THEN LossGainResult = -LossGainResult

 IF Row < UBOUND(Table, 1) - 2 THEN
  LossGainBlock.Row = Row
  LossGainBlock.Page = Page
 ELSE
  LossGainBlock.Row = 0
  IF Selection.Page = 21 THEN
   LossGainBlock.Page = Selection.Page
  ELSE
   LossGainBlock.Page = Page + 1
  END IF
 END IF

 IF IsLoss THEN Column = COL_EXPENSE ELSE Column = COL_INCOME

 Table(LossGainBlock.Row, Column, LossGainBlock.Page) = "$ " + FORMAT$(LossGainResult, NUMBER_MASK$)
 Table(LossGainBlock.Row + 1, COL_INCOME, LossGainBlock.Page) = STRING$(MAXIMUM_DATA_LENGTH, "-")
 Table(LossGainBlock.Row + 1, COL_EXPENSE, LossGainBlock.Page) = STRING$(MAXIMUM_DATA_LENGTH, "-")
 Table(LossGainBlock.Row + 2, COL_INCOME, LossGainBlock.Page) = "$ " + FORMAT$(Income, NUMBER_MASK$)
 Table(LossGainBlock.Row + 2, COL_EXPENSE, LossGainBlock.Page) = "$ " + FORMAT$(Expenses, NUMBER_MASK$)
 Table(LossGainBlock.Row + 3, COL_INCOME, LossGainBlock.Page) = STRING$(MAXIMUM_DATA_LENGTH, "=")
 Table(LossGainBlock.Row + 3, COL_EXPENSE, LossGainBlock.Page) = STRING$(MAXIMUM_DATA_LENGTH, "=")
END SUB

SUB Sort (Table() AS STRING, SortColumn AS INTEGER, LossGainBlock AS LossGainBlockStr)
DIM Column AS INTEGER
DIM OtherPage AS INTEGER
DIM OtherRow AS INTEGER
DIM Page AS INTEGER
DIM Row AS INTEGER

 FOR Page = LBOUND(Table, 3) TO UBOUND(Table, 3) - 1
  FOR Row = LBOUND(Table, 1) TO UBOUND(Table, 1)
   IF NOT RowEmpty(Table(), Row, Page) THEN
    IF NOT InLossGainBlock(Row, COL_INCOME, Page, LossGainBlock) THEN
     FOR OtherPage = LBOUND(Table, 3) TO UBOUND(Table, 3) - 1
      FOR OtherRow = LBOUND(Table, 1) TO UBOUND(Table, 1)
       IF NOT (OtherRow = Row AND OtherPage = Page) THEN
        IF NOT RowEmpty(Table(), OtherRow, OtherPage) THEN
         IF NOT InLossGainBlock(OtherRow, COL_INCOME, OtherPage, LossGainBlock) THEN
          IF Table(Row, SortColumn, Page) < Table(OtherRow, SortColumn, OtherPage) THEN
           FOR Column = LBOUND(Table, 2) TO UBOUND(Table, 2)
            SWAP Table(Row, Column, Page), Table(OtherRow, Column, OtherPage)
           NEXT Column
          END IF
         END IF
        END IF
       END IF
      NEXT OtherRow
     NEXT OtherPage
    END IF
   END IF
  NEXT Row
 NEXT Page
END SUB

FUNCTION TableBaseFileName (FileName AS STRING) AS STRING
DIM BaseFileName AS STRING

 BaseFileName = FileName
 IF INSTR(BaseFileName, ".") > 0 THEN BaseFileName = LEFT$(BaseFileName, INSTR(BaseFileName, ".") - 1)

 TableBaseFileName = BaseFileName
END FUNCTION

FUNCTION TableHeader () AS STRING
 TableHeader = " Date:" + SPACE$(15) + "Description:" + SPACE$(8) + "Income:" + SPACE$(13) + "Expenses:" + SPACE$(10)
END FUNCTION

FUNCTION TableListed (FileName AS STRING) AS INTEGER
DIM Exists AS INTEGER
DIM FileH AS INTEGER
DIM ListFile AS STRING
DIM SaveDate AS STRING
DIM SaveTime AS STRING

 IF NOT FileExists("Names.lst") THEN CreateFile "Names.lst"

 Exists = FALSE
 IgnoreError = FALSE
 FileH = FREEFILE
 OPEN "Names.lst" FOR INPUT LOCK READ WRITE AS FileH
  DO UNTIL EOF(FileH) OR Exists OR IgnoreError
   INPUT #FileH, ListFile, SaveTime, SaveDate
   Exists = (FormatTableFileName$(ListFile) = FormatTableFileName$(FileName))
  LOOP
 CLOSE FileH

 TableListed = Exists
END FUNCTION

