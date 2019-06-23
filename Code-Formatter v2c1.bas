'============================================================================
'  Format PowerBASIC source code
'
'Author Bruce Brown, August 2010
'Small Changes: Theo Gottwald
' v2  changes : Paul Elliott May 2011
' v2a changed : Paul Elliott June 2011
'                got Global/Register/Dim/Local/Static in 1 routine
' v2b Peter
' v2c Paul       some changes to format Class code a little better
'                split some CASE statements out for easier maintenance
'
'PBWIN 9
'Placed in public domain by author 8/26/10, all the usual disclaimers apply
'============================================================================
'Formats a Powerbasic source code file (BAS, INC) for consistent spacing and
'indentation. REM keywords are changed to a apostrophe (') and ASM keywords
'to !. Options provided to:
'   start all inline comments on a specific column
'     line-up the = sign of all declared equates on a specified column.
'     backup the source file prior to formatting
'     define the number of space for a tab
'     capitalize defined equates
'     split Global/Local variables to 1 per line
'Uses no global or static variables so easily could be converted to a DLL
'============================================================================
#COMPILE EXE
#DIM ALL
#TOOLS OFF

%UNICODE                          = 1
#IF %DEF(%UNICODE)
    MACRO XSTRING = WSTRING
    MACRO XASCIZ = WSTRINGZ
#ELSE
'MACRO XSTRING = STRING
'MACRO XASCIZ = ASCIZ
#ENDIF

'#INCLUDE "X_AU.inc"
#INCLUDE ONCE "Win32API.inc"
#INCLUDE ONCE "CommCtrl.inc"
TYPE fmtoptions
    capequates AS LONG                           'TRUE/FALSE capitalize equates
    remcol AS LONG                               'column position for inline remarks
    backup AS LONG                               'TRUE/FALSE backup source before formatting
    fspecin AS XASCIZ * %MAX_PATH                'source file spec
    fspecout AS XASCIZ * %MAX_PATH               'temporary output file spec
    lineupequates AS LONG                        'TRUE/FALSE position declared equates
    equatescolumn AS LONG                        'column pos of = in equates definitions
    fhin AS LONG                                 'file handle, source
    fhout AS LONG                                'file handle output
    tabsize AS LONG                              'size of tab in characters
    hprogress AS DWORD                           'handle, progress window
    indent AS LONG                               'left indenting helper
    splitvariables AS LONG                       ' split multiple variables in Global/Local to seperate lines
    ascolumn AS LONG                             ' column for AS  from multiple variables
    inClass AS LONG                              ' needed to keep track if within Class
END TYPE

MACRO S_TRIMBA(P1) = TRIM$(P1, ANY CHR$(0 TO 32))

MACRO Construct(P1, P2)
CASE P1
    IF Is_there(w(), P2, WordCount) THEN
        IF (WordNo = 1) THEN spacer = spacer + fo.tabsize
    END IF
END MACRO

MACRO m_ExitFunction(nresult)
FUNCTION = nresult
EXIT FUNCTION
END MACRO

%MAX_LINELEN                      = 400          'maximum source code line length

FUNCTION WINMAIN(BYVAL hInstance AS DWORD, BYVAL hPrevInst AS DWORD, _
    BYVAL lpszCmdLine AS XASCIZ PTR, BYVAL nCmdShow AS LONG) AS LONG
    '-------------------------------------------------------------------------------
    'Application entry
    '-------------------------------------------------------------------------------
   LOCAL fo                            AS FMTOPTIONS
   LOCAL ztext                         AS XASCIZ * %MAX_PATH
    
    'check the command line for a filespec
    IF LEN(@lpszcmdline) THEN
        IF ISFILE(@lpszcmdline) THEN
            fo.fspecin = @lpszcmdline            'assumes valid source file
        ELSE
            MessageBox 0, "File not found" + $CR + @lpszcmdline, "", %MB_ICONINFORMATION
            EXIT FUNCTION
        END IF
    ELSE
        'browse for a source file
        DISPLAY OPENFILE 0, 0, 0, "Format Powerbasic Code", "", _
        CHR$( "Basic Source [BAS, INC]", 0, "*.BAS;*.INC", 0), _
        "", "", %OFN_FILEMUSTEXIST TO fo.fspecin
        IF LEN(fo.fspecin) = 0 THEN
            EXIT FUNCTION
        END IF
    END IF
    '************************************************************
    'These values are hard coded for demo purposes. They could be
    'provided to the app with a intrinsinc dialog, INI file or
    'via the command line. See the FMTOPTION udt
    '************************************************************
    fo.capequates = 1
    fo.remcol = 50
    fo.backup = 1
    fo.lineupequates = 1
    fo.equatescolumn = 35
    fo.tabsize = 4
    fo.splitvariables = 1                        ' added v2
    fo.ascolumn = 40                             ' added v2
    '************************************************************
    IF fo.backup THEN
        CopyFile fo.fspecin, fo.fspecin + ".bak", 0
    END IF
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Get a temporary output filespec and open both input and output
    'files
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    GetTempPath %MAX_PATH, ztext
    GetTempFileName ztext, "fmt", 0, fo.fspecout
    fo.fhin = FREEFILE
    OPEN fo.fspecin FOR INPUT ACCESS READ AS fo.fhin
    fo.fhout = FREEFILE
    OPEN fo.fspecout FOR OUTPUT AS fo.fhout
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Create a simple progress control
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    fo.hprogress = CaptionProgressWindow(0, "Format Source Code")
    DoFormat BYVAL VARPTR(fo)
    CLOSE fo.fhin, fo.fhout
    DestroyWindow fo.hprogress
    CopyFile fo.fspecout, fo.fspecin, 0
END FUNCTION


SUB DoFormat(BYVAL fo AS FMTOPTIONS PTR)
    '----------------------------------------------------------------------------
    'Format the source file. The w(1-x) array holds words, tokens, strings
    'of parsed source line data. w(0) element is reserved for any in-line comment
    'that doesn't fit to be placed above the affected line.
    '----------------------------------------------------------------------------
   LOCAL Wordcount                     AS LONG                       '# tokens in source line
   LOCAL CurrentLine                   AS LONG                     '0 based current edit line number
   LOCAL TotalLines                    AS LONG                      'total # lines in control
   LOCAL ztext                         AS ASCIZ * %MAX_LINELEN           'source code line
   DIM w(100)                          AS LOCAL STRING                    'words/tokens in line
   LOCAL build                         AS STRING                         'rebuilt source line
    
    ' 05/06/2011 PDE expand GLOBAL/LOCAL to 1 variable per line
    LOCAL str1                         AS STRING
   LOCAL str2                          AS STRING
   LOCAL str3                          AS STRING
    LOCAL strBgn                       AS STRING
   LOCAL strAs                         AS STRING
    
    DIM vars(25)                       AS STRING
   LOCAL varstr                        AS STRING
   LOCAL varfixsize                    AS STRING
   LOCAL varcnt                        AS LONG
   LOCAL varDeclare                    AS STRING
   LOCAL varDeclareSize                AS LONG
    
   LOCAL a                             AS LONG
   LOCAL b                             AS LONG
   LOCAL x                             AS LONG
   LOCAL y                             AS LONG
   LOCAL z                             AS LONG
   LOCAL preAS                         AS LONG
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Start formatting, fetch each source line, parse into
    'tokens/wo
    'rds, rebuild it according to user preferences,
    'and add to temp file.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    FILESCAN @fo.fhin, RECORDS TO TotalLines     'total source code lines
    FOR currentline = 1 TO totallines
        LINE INPUT #@fo.fhin, ztext
        
      IF INSTR(UCASE$(ztext), "#PBFORMS") THEN
        GOTO NoGlobalLocal1
    END IF
    
    IF @fo.splitvariables = 0 THEN
        GOTO NoGlobalLocal
    END IF
    
    ' 06/21/2011 PDE split up Global/Local to 1 variable per line
    str1 = ztext
    
    str2 = TRIM$(str1)
    str2 = UCASE$(str2)
    
    preAS = @fo.ascolumn - 2
    
    str2 = TRIM$(str1)
    str2 = UCASE$(str2)
    varDeclare = ""
    varDeclareSize = 0
    IF LEFT$(str2, 7) = "GLOBAL " THEN
        varDeclareSize = 7
    END IF
    IF LEFT$(str2, 9) = "REGISTER " THEN
        varDeclareSize = 9
    END IF
    IF LEFT$(str2, 4) = "DIM " THEN
        varDeclareSize = 4
    END IF
    IF LEFT$(str2, 6) = "LOCAL " THEN
        varDeclareSize = 6
    END IF
    IF LEFT$(str2, 7) = "STATIC " THEN
        varDeclareSize = 7
    END IF
    IF varDeclareSize = 0 THEN GOTO NoGlobalLocal
    varDeclare = LEFT$(str2, varDeclareSize)
    
    z = INSTR(str2, " AS ")
    IF z <= 0 THEN GOTO NoGlobalLocal            ' old style?
    
    str3 = UCASE$(str1)
    
    y = INSTR(str2, ",")
    IF y <= 0 THEN                               ' ignore only 1 variable
        z = INSTR(str3, " AS ")
        IF z < preAS THEN
            b = preAS - z
            INCR b
            PRINT #@fo.fhout, LEFT$(str1, z - 1); SPACE$(b); MID$(str1, z)
        ELSE
            str2 = LEFT$(str1, z - 1)
            str2 = RTRIM$(str2)
            a = LEN(str2)
            IF a < preAS THEN
                b = preAS - a
                PRINT #@fo.fhout, LEFT$(str1, a); SPACE$(b); MID$(str1, z)
            ELSE
                PRINT #@fo.fhout, str1
            END IF
        END IF
        ITERATE FOR
    END IF                                       ' ignore only 1 variable
    
    str2 = UCASE$(str1)
    
    x = INSTR(str2, varDeclare)
    strBgn = LEFT$(str1, x + varDeclareSize - 1)
    z = INSTR(str2, " AS ")
    strAs = MID$(str1, z)
    
    x = x + varDeclareSize
    
    varstr = MID$(str1, x, z - x)
    y = preAS - LEN(strBgn)
    varfixsize = SPACE$(y)
    varcnt = PARSECOUNT(varstr, ",")
    RESET vars()
    PARSE varstr, vars(), ","
    FOR y = 0 TO varcnt - 1
        IF LEN(varfixsize) < LEN(TRIM$(vars(y))) THEN
            PRINT #@fo.fhout, "*ERR "; varDeclare; STR$(LEN(varfixsize)); " but need "; STR$(LEN(TRIM$(vars(y))))
        END IF
        LSET varfixsize = TRIM$(vars(y))
        PRINT #@fo.fhout, strBgn; varfixsize; strAs
    NEXT
    
    ITERATE FOR                                  ' done with
    '  end of splitvariable code 06/21/2011 PDE
    
NoGlobalLocal:
    IF (currentline MOD 10) = 0 THEN             'update the progress control
        SendMessage @fo.hProgress, %PBM_SETPOS, 100 * (CurrentLine / TotalLines), 0
    END IF
    WordCount = Line2Words(ztext, w(), @fo)
    build = RebuildLine(w(), WordCount, @fo)
    IF LEN(w(0)) THEN                            'inline comment doesn't fit
        PRINT #@fo.fhout, w(0)                   'this increases the output file line count
        w(0) = ""
    END IF
    PRINT #@fo.fhout, build
    ITERATE FOR
NoGlobalLocal1:
    PRINT #@fo.fhout, ztext
NEXT
END SUB


FUNCTION Line2Words(BYREF work AS ASCIZ * %MAX_LINELEN, BYREF w() AS STRING, _
    BYREF fo AS FMTOPTIONS) AS LONG
    '-------------------------------------------------------------------
    'Parse source line into word/token array w(). Handle special cases for
    'REM, DATA, string literals, ASM, !, '.
    '  work     [in/out] source code line text
    '  w()      [in/out] array of words/tokens
    '  fo       [in/out] udt of formatting options
    'Returns count of words/tokens found
    '-------------------------------------------------------------------
   LOCAL stemp                         AS STRING
   LOCAL ncount                        AS LONG                          'word/token count
   LOCAL pchr                          AS BYTE PTR
   LOCAL s1                            AS LONG
    
    ' REPLACE $TAB WITH SPACE$(fo.tabsize) IN work *** replaced with following 05/06/2011 PDE
    work = TAB$(work, fo.tabsize)
    
    work = S_TRIMBA(work)
    'X_AU "-----------------------------------"+$CRLF+TRIM$(work)+$CRLF+"--------------------------"
    stemp = UCASE$(TRIM$(work))                  'standardize incomming text
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'These need no or limited processing
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    IF LEFT$(stemp, 1) = "'" OR _                'remark
    LEN(stemp) = 0 OR _                          'blank line
    LEFT$(stemp, 5) = "DATA " THEN               'DATA statments ignored
        w(1) = TRIM$(work)
        m_ExitFunction(1)
    ELSEIF LEFT$(stemp, 4) = "REM " THEN
        w(1) = "'" + TRIM$(MID$(work, 5))        'replace REM with '
        m_ExitFunction(1)
    ELSEIF LEFT$(stemp, 4) = "ASM " THEN         'replace ASM with !
        w(1) = "!" + RTRIM$(EXTRACT$(MID$(work, 5), ANY "';"))
        w(2) = ";" + TRIM$(REMAIN$(work, ANY ";'"))
        m_ExitFunction(2)
    ELSEIF LEFT$(stemp, 1) = "!" THEN
        w(1) = TRIM$(EXTRACT$(work, ANY "';"))   'up to any comment
        w(2) = ";" + TRIM$(REMAIN$(work, ANY ";'"))
        m_ExitFunction(2)                        'two words/tokens
    END IF
    stemp = ""                                   'clear for reuse
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Parse remainder
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    pchr = VARPTR(work)                          'point to first byte
    DO
        IF @pchr = 34 THEN                       'quote start?
            GOSUB SaveWord                       'save any current word/token
            s1 = pchr                            'save string start
            DO
                INCR pchr                        'next string character
            LOOP UNTIL @pchr = 34                'quoted string finished
            stemp = PEEK$(s1, pchr - s1 + 1)
            GOSUB SaveWord
            INCR pchr                            'next byte/character
            ITERATE DO
        END IF
        IF @pchr = 39 THEN                       'inline "'" remark
            GOSUB SaveWord
            'balance of line
            stemp = MID$(work, pchr - VARPTR(work) + 1)
            EXIT DO                              'line done
        END IF
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Check for one of the standard delimters
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        IF INSTR( " ,\=+-/^*)(:[];><.", CHR$(@pchr)) THEN
            GOSUB SaveWord
            IF @pchr <> 32 THEN
                stemp = CHR$(@pchr)
                GOSUB SaveWord
            END IF
        ELSE
            stemp = stemp + CHR$(@pchr)          'add char to current token
        END IF
        INCR pchr
    LOOP WHILE @pchr <> 0                        'check for line end
    GOSUB SaveWord
    m_ExitFunction(ncount)                       'return token count
    
SaveWord:
    IF LEN(stemp) THEN                           'anything to do?
        INCR ncount                              'next array element
        w(ncount) = stemp                        'save it
        '   X_AU STR$(ncount)+"-"+stemp
        stemp = ""                               'erase working string
    END IF
    RETURN
    
END FUNCTION

'----------------------------------------------------------------

FUNCTION Is_there(BYREF w() AS STRING, BYVAL S01 AS STRING, BYVAL wc AS LONG) AS LONG
   REGISTER R01                        AS LONG, R02 AS LONG
    R02 = 0
    FOR R01 = 2 TO wc
        IF TRIM$(UCASE$(w(R01))) = UCASE$(S01) THEN R02 = 1
    NEXT
    FUNCTION = 1 - R02                           ' Inverted for easier usage
END FUNCTION

'----------------------------------------------------------------

FUNCTION RebuildLine(BYREF w() AS STRING, BYREF wordcount AS LONG, _
    BYREF fo AS FMTOPTIONS) AS STRING
    '----------------------------------------------------------------
    'Reassemble a source line from the individual tokens in string
    'array Words() standardize puncuation, spacing and indentation.
    '  w()         [in/out] array of source words/tokens
    '  wordcount   [in/out] # words/tokens in array
    '  fo          [in/out] udt of formatting options
    'Returns the formatted source code line
    '----------------------------------------------------------------
    LOCAL i&                                     ' NASTY HABIT
   LOCAL stemp                         AS STRING
   LOCAL comment                       AS STRING
   LOCAL schr                          AS STRING
   LOCAL spacer                        AS LONG                          'helper with indenting
   LOCAL WordNo                        AS LONG
    
    schr = LEFT$(w(1), 1)
    IF schr = "'" THEN
        m_ExitFunction(SPACE$(fo.indent) + w(1)) 'return full comment line
    ELSEIF UCASE$(LEFT$(w(1), 5)) = "DATA" THEN  'line is data
        m_ExitFunction(w(1))                     'return unmodified
    ELSEIF schr = "!" THEN
        IF LEN(w(2)) > 1 THEN                    'inline assembler comment
            stemp = LSET$(SPACE$(fo.indent) + w(1), fo.remcol - 1) + w(2)
        ELSE
            stemp = SPACE$(fo.indent) + w(1)
        END IF
        m_ExitFunction(stemp)
    END IF
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'If last item in token array is a inline comment, save it to
    'be inserted later.
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    schr = LEFT$(w(WordCount), 1)
    IF schr = "'" OR schr = ";" THEN
        comment = w(WordCount)
        IF LEN(comment) <= 1 THEN
            comment = ""                         'make blank line
        END IF
        DECR WordCount
    END IF
    IF fo.lineupequates THEN
        schr = LEFT$(w(1), 1)
        IF (schr = "%" OR schr = "$") AND w(2) = "=" THEN
            IF fo.capequates THEN
                w(1) = UCASE$(w(1))
            END IF
            i& = MAX&(LEN(w(1)) + 2, fo.equatescolumn)
            stemp = LSET$(w(1), i& - 2) + " = "
            FOR i& = 3 TO wordcount
                stemp = stemp + TRIM$(w(i&))
            NEXT
            IF LEN(comment) AND LEN(stemp) < fo.remcol - 1 THEN
                stemp = LSET$(stemp, fo.remcol - 1) + TRIM$(comment)
            ELSE
                w(0) = comment
                comment = ""
            END IF
            m_ExitFunction(stemp)
        END IF
    END IF
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'process each token in array w()
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    stemp = ""
    DO
        INCR WordNo                              'first/next word
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'Is word abreviated print statement ?. Expand and capitilize
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        schr = LEFT$(w(WordNo), 1)
        IF w(WordNo) = "?" THEN                  'expand to PRINT (PBCC)
            w(WordNo) = "PRINT"
        ELSEIF schr = $DQ THEN                   'string literal
            stemp = stemp + w(WordNo) + " "      'simply add it
            ITERATE DO
        ELSEIF schr = "%" OR schr = "$" THEN     'equates
            IF fo.capequates THEN
                w(WordNo) = UCASE$(w(WordNo))
            END IF
        ELSEIF schr = "!" THEN
            stemp = stemp + w(WordNo)            'assembler line
            spacer = fo.indent
            fo.indent = 0
            EXIT DO
        END IF
        
        SELECT CASE UCASE$(w(WordNo))            'process the token
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'Check to see if token is a label (wordcount = 2), only thing
                'allowed on a line with a label is a comment and that was
                'removed above. Should it fail this test its assumed to be
                'part of a multi-statement line. Labels always start in col 1
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE ":"                             'check for label
                stemp = RTRIM$(stemp)
                IF WordCount = 2 THEN            'we assume a label (name + :)
                    spacer = fo.indent           'save indent value
                    fo.indent = 0                'move label to left margin
                END IF
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'Try to distinguish between array and something else as to
                'whether or not there should be a space preecding these. Need
                'to test for a PB intrinsic function like ATTRIB(xxx) vs OR
                '(xx + yy).
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE "(", "["
                IF INSTR( "+-*/\=><", w(WordNo - 1)) = 0 THEN
                    stemp = RTRIM$(stemp)
                END IF
                SELECT CASE UCASE$(w(WordNo - 1))
                    CASE "IF", "ELSEIF", "AND", "OR", "NOT", "ISFALSE", "ISTRUE", "XOR", "TO"
                        stemp = stemp + " "      'add spacing
                END SELECT
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'We never want any space to proceed these delimiters
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE ")", ",", "]", "[", ";", "."
                stemp = RTRIM$(stemp)            'no preceding space
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'Catch <>, >=, <= combinations to eliminate < >, > =,. etc.
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE "<", ">", "="
                ' 05/06/2011 PDE added + - * not sure how to handle OR AND
                IF INSTR( "<>=+-*ORAND", w(WordNo - 1)) THEN
                    stemp = RTRIM$(stemp)
                END IF
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'Compare last token to THEN, nothing else on line means started
                'a multi-line IF block. Any comment already removed above.
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE "THEN"
                IF WordNo = WordCount THEN
                    spacer = spacer + fo.tabsize 'change for next line
                END IF
            CASE "TYPE", "UNION", "TRY", "FUNCTION", "SUB", "INTERFACE", "PROPERTY"
                IF (WordNo = 1) AND (w(WordNo + 1) <> "=") THEN spacer = spacer + fo.tabsize
                
            CASE "CLASS"
                IF WordNo = 1 AND UCASE$(w(WordNo + 1)) <> "METHOD" THEN
                    INCR fo.inClass
                END IF
                IF WordNo = 1 AND w(WordNo + 1) <> "=" THEN spacer = spacer + fo.tabsize
                
            CASE "METHOD"
                IF WordNo = 1 AND fo.inClass > 0 THEN
                    spacer = spacer + fo.tabsize ' next line
                END IF
                
                '         case "FOR"
                '          if Is_there(w(),"NEXT",WordCount) then
                '             IF (WordNo = 1) THEN spacer = spacer + fo.tabsize
                '          end if
                Construct( "FOR", "NEXT")
                Construct( "WHILE", "WEND")
                Construct( "DO", "LOOP")
                
            CASE "MACRO"
                IF WordNo = 1 AND UCASE$(w(2)) = "FUNCTION" THEN
                    spacer = spacer + fo.tabsize
                END IF
            CASE "CALLBACK"                      ' 05/09/2011 PDE need to indent same as normal FUNCTION
                IF WordNo = 1 AND UCASE$(w(2)) = "FUNCTION" THEN
                    spacer = spacer + fo.tabsize
                END IF
            CASE "SELECT"
                IF WordNo = 1 THEN spacer = spacer + (fo.tabsize * 2)
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'Handle ELSEIF seperately since it appears on same line with THEN
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE "#IF"                           'these require no line ending THEN
                spacer = spacer + fo.tabsize
            CASE "ELSEIF", "#ELSEIF", "#ELSE"
                'change for this line
                fo.indent = fo.indent - fo.tabsize
            CASE "#ENDIF"
                fo.indent = fo.indent - fo.tabsize
                spacer = 0
            CASE "ELSE", "CASE"
                IF WordNo = 1 THEN               'first word?
                    spacer = spacer + fo.tabsize 'change for next line
                    'change for this line
                    fo.indent = fo.indent - fo.tabsize
                END IF
            CASE "END"
                SELECT CASE UCASE$(w(WordNo + 1))
                    CASE "TYPE", "IF", "UNION", "SUB", "FUNCTION", "INTERFACE", "TRY", "MACRO", "METHOD", "PROPERTY"
                        fo.indent = fo.indent - fo.tabsize
                    CASE "CLASS"
                        DECR fo.inClass
                        spacer = 0               ' fo.tabsize ' next line
                        ' fo.indent - fo.tabsize ' this line
                        fo.indent = fo.indent - fo.tabsize
                    CASE "SELECT"
                        fo.indent = fo.indent - (fo.tabsize * 2)
                END SELECT
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                'Check for the close of a looping block. If its on the same
                'line as loop start, keep same indent.
                '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            CASE "WEND", "NEXT", "LOOP"
                IF wordno = 1 THEN
                    schr = UCASE$(stemp)         'standardize
                    spacer = 0                   'be sure this = 0
                    'default is close loop
                    fo.indent = fo.indent - fo.tabsize
                    IF INSTR(schr, "DO ") THEN
                        fo.indent = fo.indent + fo.tabsize
                    ELSEIF INSTR(schr, "WHILE ") THEN
                        fo.indent = fo.indent + fo.tabsize
                    ELSEIF INSTR(schr, "ITERATE") THEN
                        fo.indent = fo.indent + fo.tabsize
                    END IF
                END IF
        END SELECT
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        'No space following "([]"
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        IF w(wordno) = "." THEN
            stemp = RTRIM$(stemp) + "."
        ELSEIF INSTR( "(][", w(WordNo - 1)) THEN
            stemp = RTRIM$(stemp) + w(WordNo) + " "
        ELSE
            stemp = stemp + w(WordNo) + " "
        END IF
    LOOP WHILE WordNo < WordCount
    IF fo.indent < 0 THEN fo.indent = 0          'stay at left margin
    stemp = SPACE$(fo.indent) + RTRIM$(stemp)    'final formated line
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Add any line comment preserved above. Fit failure returns
    'inline comment in w(0)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    w(0) = ""
    IF LEN(comment) THEN
        IF LEN(stemp) < fo.remcol THEN
            stemp = LEFT$(stemp + SPACE$(fo.RemCol - 1), fo.RemCol - 1) + comment
        ELSE
            w(0) = SPACE$(fo.indent) + comment
        END IF
    END IF
    fo.indent = fo.indent + spacer               'new indent value
    FUNCTION = stemp                             'return formatted source line
END FUNCTION

FUNCTION CaptionProgressWindow(BYVAL hparent AS DWORD, BYVAL PBCAPTION AS XSTRING) AS DWORD
    '--------------------------------------------------------------------------
    'Create a centered progress control with caption and return handle
    '--------------------------------------------------------------------------
   LOCAL rc                            AS RECT
   LOCAL hwin                          AS DWORD                           'progress control handle
   LOCAL ht                            AS LONG                          'window height & width
   LOCAL wd                            AS LONG                          'window height & width
    'fetch desktop size
    SystemParametersInfo %SPI_GETWORKAREA, 0, rc, 0
    ht = GetSystemMetrics(%SM_CYCAPTION)         'height of window caption
    wd = rc.nright \ 2                           '1/2 of screen width
    IF hparent = 0 THEN
        hparent = GetDesktopWindow
    END IF
    hwin = CreateWindowEx(%WS_EX_TOOLWINDOW OR %WS_EX_DLGMODALFRAME, "msctls_progress32", _
    BYVAL STRPTR(PBCAPTION), %WS_VISIBLE OR %WS_CAPTION OR %WS_POPUP OR %PBS_SMOOTH, _
    (rc.nright \ 2) - (wd \ 2), _                'horiz pos start
    (rc.nbottom \ 2) - ht, _                     'vert  pos start
    wd, _                                        'width, 1/2 screen
    ht * 2, _                                    'height, 2 * caption
    hparent, 0, GetModuleHandle(BYVAL %NULL), BYVAL %NULL)
    FUNCTION = hwin                              'return progress handle
END FUNCTION
