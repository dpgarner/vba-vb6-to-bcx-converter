' ===================================================================
' vb2bcx_lib.bas - shared library for the vb2bcx converter.
'
' No top-level executable code of its own -- pulled into another file
' with '$INCLUDE: 'vb2bcx_lib.bas'  (see vb2bcx.bas / vb2bcx_test.bas).
' If your BCX rejects that metacommand form, paste this file's
' contents in place of the include line instead.
'
' NOTE ON TESTING: written and hand-traced, but NOT compiled -- BCX is
' a Windows-targeted tool (BCX.exe + a C compiler) and neither is
' available in the sandbox this was written in. Please compile it and
' report anything the compiler complains about.
' ===================================================================

DECLARE FUNCTION IsAlNumChar(ch$) AS INTEGER
DECLARE FUNCTION IsRemHere(ln$, i) AS INTEGER
DECLARE FUNCTION MaskLine$(ln$)
DECLARE FUNCTION FindKeyword(hay$, startPos, word$) AS INTEGER
DECLARE FUNCTION KeywordPresent(masked$, word$) AS INTEGER
DECLARE FUNCTION KeywordFollowedByChar(masked$, word$, ch$) AS INTEGER
DECLARE FUNCTION PhraseFollows(masked$, w1$, w2$) AS INTEGER
DECLARE FUNCTION SkipSpacesFrom(s$, pos) AS INTEGER
DECLARE FUNCTION LeadingWhitespace$(ln$)
DECLARE FUNCTION DefaultOutputName$(p$)
DECLARE FUNCTION PrecededByAs(ln$, matchPos) AS INTEGER
DECLARE FUNCTION IsPlainInteger(s$) AS INTEGER
DECLARE FUNCTION IsEnumStart(masked$, enumName$) AS INTEGER
DECLARE FUNCTION IsEnumEnd(masked$) AS INTEGER
DECLARE FUNCTION IsAttributeLine(masked$) AS INTEGER
DECLARE FUNCTION IsClassModuleHeader(masked$) AS INTEGER
DECLARE FUNCTION HasOnErrorResumeNext(masked$) AS INTEGER
DECLARE FUNCTION HasErrDot(masked$) AS INTEGER
DECLARE FUNCTION HasIsNothing(masked$) AS INTEGER
DECLARE FUNCTION HasEqualsNothing(masked$) AS INTEGER
DECLARE FUNCTION HasCreateObjectCall(masked$) AS INTEGER
DECLARE FUNCTION HasSetNew(masked$) AS INTEGER
DECLARE FUNCTION HasPropertyGetLetSet(masked$) AS INTEGER
DECLARE FUNCTION IsConditionalCompilation(ln$) AS INTEGER
DECLARE SUB AddLine(arr$(), cnt, val$)
DECLARE SUB ParseArgs(cmd$, args$(), argCount)
DECLARE SUB ParseEnumMember(ln$, memberName$, memberVal$, hasVal)
DECLARE SUB EmitFlag(outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, reason$)
DECLARE SUB StripVisibility(ln$)
DECLARE SUB ReplaceAsBoolean(ln$, lineNo, warnings$(), warnCount)
DECLARE SUB ReplaceEndif(ln$)
DECLARE SUB ReplaceVbConstants(ln$, lineNo, warnings$(), warnCount)
DECLARE SUB RenameReserved(ln$, lineNo, warnings$(), warnCount)
DECLARE SUB FlagUnsupported(masked$, ln$, lineNo, warnings$(), warnCount, outlines$(), outCount, indent$)
DECLARE SUB ConvertProgram(lines$(), lineCount, outlines$(), outCount, warnings$(), warnCount)

' ===================================================================
' low-level character / masking helpers (same design as qb2bcx_lib.bas)
' ===================================================================

FUNCTION IsAlNumChar(ch$) AS INTEGER
    DIM c$
    IF ch$ = "" THEN
        IsAlNumChar = 0
        EXIT FUNCTION
    END IF
    c$ = UCASE$(ch$)
    IF (c$ >= "A" AND c$ <= "Z") OR (ch$ >= "0" AND ch$ <= "9") OR ch$ = "_" THEN
        IsAlNumChar = 1
    ELSE
        IsAlNumChar = 0
    END IF
END FUNCTION

' VBA/VB6 comments only use ' (REM is legal too, but vanishingly rare in
' modern VB6 code) -- kept for parity with the QB port's masking style.
FUNCTION IsRemHere(ln$, i) AS INTEGER
    DIM n, beforeOK, afterOK
    n = LEN(ln$)
    IF i + 2 > n THEN
        IsRemHere = 0
        EXIT FUNCTION
    END IF
    IF UCASE$(MID$(ln$, i, 3)) <> "REM" THEN
        IsRemHere = 0
        EXIT FUNCTION
    END IF
    IF i = 1 THEN
        beforeOK = 1
    ELSE
        beforeOK = (IsAlNumChar(MID$(ln$, i - 1, 1)) = 0)
    END IF
    IF i + 3 > n THEN
        afterOK = 1
    ELSE
        afterOK = (IsAlNumChar(MID$(ln$, i + 3, 1)) = 0)
    END IF
    IF beforeOK = 1 AND afterOK = 1 THEN
        IsRemHere = 1
    ELSE
        IsRemHere = 0
    END IF
END FUNCTION

FUNCTION MaskLine$(ln$)
    DIM result$, i, n, ch$, inString
    n = LEN(ln$)
    result$ = ""
    inString = 0
    i = 1
    DO WHILE i <= n
        ch$ = MID$(ln$, i, 1)
        IF inString = 1 THEN
            IF ch$ = CHR$(34) THEN
                result$ = result$ + CHR$(34)
                inString = 0
            ELSE
                result$ = result$ + "x"
            END IF
            i = i + 1
        ELSE
            IF ch$ = CHR$(34) THEN
                inString = 1
                result$ = result$ + CHR$(34)
                i = i + 1
            ELSEIF ch$ = "'" THEN
                result$ = result$ + SPACE$(n - i + 1)
                i = n + 1
            ELSEIF (UCASE$(ch$) = "R") AND (IsRemHere(ln$, i) = 1) THEN
                result$ = result$ + SPACE$(n - i + 1)
                i = n + 1
            ELSE
                result$ = result$ + ch$
                i = i + 1
            END IF
        END IF
    LOOP
    MaskLine$ = result$
END FUNCTION

FUNCTION FindKeyword(hay$, startPos, word$) AS INTEGER
    DIM upHay$, upWord$, p, wlen, hlen, beforeOK, afterOK
    upHay$ = UCASE$(hay$)
    upWord$ = UCASE$(word$)
    wlen = LEN(word$)
    hlen = LEN(hay$)
    p = startPos
    DO WHILE p <= hlen
        p = INSTR(p, upHay$, upWord$)
        IF p = 0 THEN
            FindKeyword = 0
            EXIT FUNCTION
        END IF
        IF p = 1 THEN
            beforeOK = 1
        ELSE
            beforeOK = (IsAlNumChar(MID$(hay$, p - 1, 1)) = 0)
        END IF
        IF p + wlen > hlen THEN
            afterOK = 1
        ELSE
            afterOK = (IsAlNumChar(MID$(hay$, p + wlen, 1)) = 0)
        END IF
        IF beforeOK = 1 AND afterOK = 1 THEN
            FindKeyword = p
            EXIT FUNCTION
        END IF
        p = p + 1
    LOOP
    FindKeyword = 0
END FUNCTION

FUNCTION KeywordPresent(masked$, word$) AS INTEGER
    IF FindKeyword(masked$, 1, word$) <> 0 THEN
        KeywordPresent = 1
    ELSE
        KeywordPresent = 0
    END IF
END FUNCTION

FUNCTION KeywordFollowedByChar(masked$, word$, ch$) AS INTEGER
    DIM p, q, hlen
    hlen = LEN(masked$)
    p = FindKeyword(masked$, 1, word$)
    DO WHILE p <> 0
        q = p + LEN(word$)
        DO WHILE q <= hlen AND MID$(masked$, q, 1) = " "
            q = q + 1
        LOOP
        IF q <= hlen AND MID$(masked$, q, 1) = ch$ THEN
            KeywordFollowedByChar = 1
            EXIT FUNCTION
        END IF
        p = FindKeyword(masked$, p + 1, word$)
    LOOP
    KeywordFollowedByChar = 0
END FUNCTION

' True if w1$ occurs (as a whole word) directly followed by whitespace
' then w2$ -- used for "As Boolean", "As Variant", "Property Get", etc.
FUNCTION PhraseFollows(masked$, w1$, w2$) AS INTEGER
    DIM p, q, hlen
    hlen = LEN(masked$)
    p = FindKeyword(masked$, 1, w1$)
    DO WHILE p <> 0
        q = p + LEN(w1$)
        DO WHILE q <= hlen AND MID$(masked$, q, 1) = " "
            q = q + 1
        LOOP
        IF FindKeyword(masked$, q, w2$) = q THEN
            PhraseFollows = 1
            EXIT FUNCTION
        END IF
        p = FindKeyword(masked$, p + 1, w1$)
    LOOP
    PhraseFollows = 0
END FUNCTION

FUNCTION SkipSpacesFrom(s$, pos) AS INTEGER
    DIM p, hlen
    p = pos
    hlen = LEN(s$)
    DO WHILE p <= hlen AND MID$(s$, p, 1) = " "
        p = p + 1
    LOOP
    SkipSpacesFrom = p
END FUNCTION

FUNCTION LeadingWhitespace$(ln$)
    DIM i, n
    n = LEN(ln$)
    i = 1
    DO WHILE i <= n AND (MID$(ln$, i, 1) = " " OR MID$(ln$, i, 1) = CHR$(9))
        i = i + 1
    LOOP
    LeadingWhitespace$ = LEFT$(ln$, i - 1)
END FUNCTION

FUNCTION DefaultOutputName$(p$)
    DIM i, dotPos, slashPos, base$, ext$
    dotPos = 0
    slashPos = 0
    FOR i = 1 TO LEN(p$)
        IF MID$(p$, i, 1) = "." THEN dotPos = i
        IF MID$(p$, i, 1) = "\" OR MID$(p$, i, 1) = "/" THEN slashPos = i
    NEXT i
    IF dotPos > slashPos THEN
        base$ = LEFT$(p$, dotPos - 1)
        ext$ = MID$(p$, dotPos)
    ELSE
        base$ = p$
        ext$ = ".bas"
    END IF
    DefaultOutputName$ = base$ + "_bcx" + ext$
END FUNCTION

SUB AddLine(arr$(), cnt, val$)
    cnt = cnt + 1
    REDIM PRESERVE arr$(cnt)
    arr$(cnt) = val$
END SUB

SUB ParseArgs(cmd$, args$(), argCount)
    DIM i, n, ch$, cur$, inQuote
    argCount = 0
    n = LEN(cmd$)
    cur$ = ""
    inQuote = 0
    i = 1
    DO WHILE i <= n
        ch$ = MID$(cmd$, i, 1)
        IF inQuote = 1 THEN
            IF ch$ = CHR$(34) THEN
                inQuote = 0
            ELSE
                cur$ = cur$ + ch$
            END IF
        ELSE
            IF ch$ = CHR$(34) THEN
                inQuote = 1
            ELSEIF ch$ = " " THEN
                IF LEN(cur$) > 0 THEN
                    argCount = argCount + 1
                    REDIM PRESERVE args$(argCount)
                    args$(argCount) = cur$
                    cur$ = ""
                END IF
            ELSE
                cur$ = cur$ + ch$
            END IF
        END IF
        i = i + 1
    LOOP
    IF LEN(cur$) > 0 THEN
        argCount = argCount + 1
        REDIM PRESERVE args$(argCount)
        args$(argCount) = cur$
    END IF
END SUB

' ===================================================================
' VB6/VBA-specific helpers
' ===================================================================

' True if the nearest preceding word (ignoring whitespace) before
' matchPos is "As" -- meaning this occurrence is a VB type name in a
' declaration ("As Double"), not a variable/parameter identifier, and
' must NOT be renamed as a C-keyword collision.
FUNCTION PrecededByAs(ln$, matchPos) AS INTEGER
    DIM i, wordStart, wordEnd, word$
    i = matchPos - 1
    DO WHILE i >= 1 AND MID$(ln$, i, 1) = " "
        i = i - 1
    LOOP
    IF i < 1 THEN
        PrecededByAs = 0
        EXIT FUNCTION
    END IF
    wordEnd = i
    wordStart = i
    DO WHILE wordStart >= 1 AND IsAlNumChar(MID$(ln$, wordStart, 1)) = 1
        wordStart = wordStart - 1
    LOOP
    wordStart = wordStart + 1
    IF wordStart > wordEnd THEN
        PrecededByAs = 0
        EXIT FUNCTION
    END IF
    word$ = MID$(ln$, wordStart, wordEnd - wordStart + 1)
    IF UCASE$(word$) = "AS" THEN
        PrecededByAs = 1
    ELSE
        PrecededByAs = 0
    END IF
END FUNCTION

' True if s$ (trimmed) is an optionally-negative run of digits, e.g. an
' Enum member's explicit "= 10" value -- as opposed to an expression
' like "= 1 + OFFSET" that we can't safely fold into a running count.
FUNCTION IsPlainInteger(s$) AS INTEGER
    DIM t$, i, hlen, startPos
    t$ = LTRIM$(RTRIM$(s$))
    hlen = LEN(t$)
    IF hlen = 0 THEN
        IsPlainInteger = 0
        EXIT FUNCTION
    END IF
    startPos = 1
    IF MID$(t$, 1, 1) = "-" THEN startPos = 2
    IF startPos > hlen THEN
        IsPlainInteger = 0
        EXIT FUNCTION
    END IF
    FOR i = startPos TO hlen
        IF MID$(t$, i, 1) < "0" OR MID$(t$, i, 1) > "9" THEN
            IsPlainInteger = 0
            EXIT FUNCTION
        END IF
    NEXT i
    IsPlainInteger = 1
END FUNCTION

' "[Public|Private] Enum <name>" -- returns the name in enumName$.
FUNCTION IsEnumStart(masked$, enumName$) AS INTEGER
    DIM start, hlen, p, q, j
    hlen = LEN(masked$)
    start = SkipSpacesFrom(masked$, 1)
    p = start
    IF FindKeyword(masked$, start, "PUBLIC") = start THEN
        p = start + 6
    ELSEIF FindKeyword(masked$, start, "PRIVATE") = start THEN
        p = start + 7
    END IF
    p = SkipSpacesFrom(masked$, p)
    IF FindKeyword(masked$, p, "ENUM") = p THEN
        q = SkipSpacesFrom(masked$, p + 4)
        j = q
        DO WHILE j <= hlen AND IsAlNumChar(MID$(masked$, j, 1)) = 1
            j = j + 1
        LOOP
        IF j > q THEN
            enumName$ = MID$(masked$, q, j - q)
            IsEnumStart = 1
            EXIT FUNCTION
        END IF
    END IF
    enumName$ = ""
    IsEnumStart = 0
END FUNCTION

FUNCTION IsEnumEnd(masked$) AS INTEGER
    DIM start, p
    start = SkipSpacesFrom(masked$, 1)
    IF FindKeyword(masked$, start, "END") = start THEN
        p = SkipSpacesFrom(masked$, start + 3)
        IF FindKeyword(masked$, p, "ENUM") = p THEN
            IsEnumEnd = 1
            EXIT FUNCTION
        END IF
    END IF
    IsEnumEnd = 0
END FUNCTION

' Parses a trimmed Enum member line: "Name" or "Name = expr".
SUB ParseEnumMember(ln$, memberName$, memberVal$, hasVal)
    DIM t$, j, hlen, rest$
    memberName$ = ""
    memberVal$ = ""
    hasVal = 0
    t$ = LTRIM$(RTRIM$(ln$))
    hlen = LEN(t$)
    j = 1
    DO WHILE j <= hlen AND IsAlNumChar(MID$(t$, j, 1)) = 1
        j = j + 1
    LOOP
    IF j = 1 THEN EXIT SUB
    memberName$ = LEFT$(t$, j - 1)
    rest$ = LTRIM$(MID$(t$, j))
    IF LEFT$(rest$, 1) = "=" THEN
        memberVal$ = LTRIM$(MID$(rest$, 2))
        hasVal = 1
    END IF
END SUB

' Drops VB6 IDE export noise like  Attribute VB_Name = "Module1"
FUNCTION IsAttributeLine(masked$) AS INTEGER
    DIM start, hlen, p
    hlen = LEN(masked$)
    start = SkipSpacesFrom(masked$, 1)
    IF FindKeyword(masked$, start, "ATTRIBUTE") = start THEN
        p = SkipSpacesFrom(masked$, start + 9)
        IF p + 2 <= hlen THEN
            IF UCASE$(MID$(masked$, p, 3)) = "VB_" THEN
                IsAttributeLine = 1
                EXIT FUNCTION
            END IF
        END IF
    END IF
    IsAttributeLine = 0
END FUNCTION

' "VERSION 1.0 CLASS" -- the .cls file header VB6 writes on export.
FUNCTION IsClassModuleHeader(masked$) AS INTEGER
    DIM start, hlen, p, q
    hlen = LEN(masked$)
    start = SkipSpacesFrom(masked$, 1)
    IF FindKeyword(masked$, start, "VERSION") = start THEN
        p = SkipSpacesFrom(masked$, start + 7)
        IF p + 2 <= hlen THEN
            IF MID$(masked$, p, 3) = "1.0" THEN
                q = SkipSpacesFrom(masked$, p + 3)
                IF FindKeyword(masked$, q, "CLASS") = q THEN
                    IsClassModuleHeader = 1
                    EXIT FUNCTION
                END IF
            END IF
        END IF
    END IF
    IsClassModuleHeader = 0
END FUNCTION

FUNCTION HasOnErrorResumeNext(masked$) AS INTEGER
    DIM p, q
    p = FindKeyword(masked$, 1, "ON")
    DO WHILE p <> 0
        q = SkipSpacesFrom(masked$, p + 2)
        IF FindKeyword(masked$, q, "ERROR") = q THEN
            q = SkipSpacesFrom(masked$, q + 5)
            IF FindKeyword(masked$, q, "RESUME") = q THEN
                q = SkipSpacesFrom(masked$, q + 6)
                IF FindKeyword(masked$, q, "NEXT") = q THEN
                    HasOnErrorResumeNext = 1
                    EXIT FUNCTION
                END IF
            END IF
        END IF
        p = FindKeyword(masked$, p + 1, "ON")
    LOOP
    HasOnErrorResumeNext = 0
END FUNCTION

FUNCTION HasErrDot(masked$) AS INTEGER
    DIM p, q, hlen
    hlen = LEN(masked$)
    p = FindKeyword(masked$, 1, "ERR")
    DO WHILE p <> 0
        q = SkipSpacesFrom(masked$, p + 3)
        IF q <= hlen AND MID$(masked$, q, 1) = "." THEN
            HasErrDot = 1
            EXIT FUNCTION
        END IF
        p = FindKeyword(masked$, p + 1, "ERR")
    LOOP
    HasErrDot = 0
END FUNCTION

FUNCTION HasIsNothing(masked$) AS INTEGER
    DIM p, q
    p = FindKeyword(masked$, 1, "IS")
    DO WHILE p <> 0
        q = SkipSpacesFrom(masked$, p + 2)
        IF FindKeyword(masked$, q, "NOTHING") = q THEN
            HasIsNothing = 1
            EXIT FUNCTION
        END IF
        p = FindKeyword(masked$, p + 1, "IS")
    LOOP
    HasIsNothing = 0
END FUNCTION

FUNCTION HasEqualsNothing(masked$) AS INTEGER
    DIM p, q
    p = 1
    DO
        p = INSTR(p, masked$, "=")
        IF p = 0 THEN EXIT DO
        q = SkipSpacesFrom(masked$, p + 1)
        IF FindKeyword(masked$, q, "NOTHING") = q THEN
            HasEqualsNothing = 1
            EXIT FUNCTION
        END IF
        p = p + 1
    LOOP
    HasEqualsNothing = 0
END FUNCTION

FUNCTION HasCreateObjectCall(masked$) AS INTEGER
    HasCreateObjectCall = KeywordFollowedByChar(masked$, "CREATEOBJECT", "(")
END FUNCTION

' "Set <ident> = New" -- object instantiation.
FUNCTION HasSetNew(masked$) AS INTEGER
    DIM p, q, j, r, hlen
    hlen = LEN(masked$)
    p = FindKeyword(masked$, 1, "SET")
    DO WHILE p <> 0
        q = SkipSpacesFrom(masked$, p + 3)
        j = q
        DO WHILE j <= hlen AND IsAlNumChar(MID$(masked$, j, 1)) = 1
            j = j + 1
        LOOP
        IF j > q THEN
            r = SkipSpacesFrom(masked$, j)
            IF r <= hlen AND MID$(masked$, r, 1) = "=" THEN
                r = SkipSpacesFrom(masked$, r + 1)
                IF FindKeyword(masked$, r, "NEW") = r THEN
                    HasSetNew = 1
                    EXIT FUNCTION
                END IF
            END IF
        END IF
        p = FindKeyword(masked$, p + 1, "SET")
    LOOP
    HasSetNew = 0
END FUNCTION

FUNCTION HasPropertyGetLetSet(masked$) AS INTEGER
    IF PhraseFollows(masked$, "PROPERTY", "GET") = 1 THEN
        HasPropertyGetLetSet = 1
        EXIT FUNCTION
    END IF
    IF PhraseFollows(masked$, "PROPERTY", "LET") = 1 THEN
        HasPropertyGetLetSet = 1
        EXIT FUNCTION
    END IF
    IF PhraseFollows(masked$, "PROPERTY", "SET") = 1 THEN
        HasPropertyGetLetSet = 1
        EXIT FUNCTION
    END IF
    HasPropertyGetLetSet = 0
END FUNCTION

FUNCTION IsConditionalCompilation(ln$) AS INTEGER
    IF LEFT$(LTRIM$(ln$), 1) = "#" THEN
        IsConditionalCompilation = 1
    ELSE
        IsConditionalCompilation = 0
    END IF
END FUNCTION

' Appends both a ' TODO(BCX): comment (to the output) and a matching
' warning (to the summary) in one call.
SUB EmitFlag(outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, reason$)
    AddLine outlines$(), outCount, indent$ + "' TODO(BCX): " + reason$
    AddLine warnings$(), warnCount, "line " + LTRIM$(RTRIM$(STR$(lineNo))) + ": " + reason$
END SUB

' Strips a leading Public/Private/Friend modifier when it's directly in
' front of Sub/Function/Dim/Const/Type/Enum -- BCX has no module-level
' visibility model and wouldn't recognize the keyword.
SUB StripVisibility(ln$)
    DIM masked$, start, hlen, kwLen, p
    masked$ = MaskLine$(ln$)
    hlen = LEN(masked$)
    start = SkipSpacesFrom(masked$, 1)
    kwLen = 0
    IF FindKeyword(masked$, start, "PUBLIC") = start THEN
        kwLen = 6
    ELSEIF FindKeyword(masked$, start, "PRIVATE") = start THEN
        kwLen = 7
    ELSEIF FindKeyword(masked$, start, "FRIEND") = start THEN
        kwLen = 6
    END IF
    IF kwLen > 0 THEN
        p = SkipSpacesFrom(masked$, start + kwLen)
        IF FindKeyword(masked$, p, "SUB") = p OR FindKeyword(masked$, p, "FUNCTION") = p OR FindKeyword(masked$, p, "DIM") = p OR FindKeyword(masked$, p, "CONST") = p OR FindKeyword(masked$, p, "TYPE") = p OR FindKeyword(masked$, p, "ENUM") = p THEN
            ln$ = LEFT$(ln$, start - 1) + MID$(ln$, p)
        END IF
    END IF
END SUB

' "As Boolean" -> "As Integer" (BCX has no boolean type).
SUB ReplaceAsBoolean(ln$, lineNo, warnings$(), warnCount)
    DIM masked$, p, q, newln$, lastPos, changed
    masked$ = MaskLine$(ln$)
    newln$ = ""
    lastPos = 1
    changed = 0
    p = FindKeyword(masked$, 1, "AS")
    DO WHILE p <> 0
        q = SkipSpacesFrom(masked$, p + 2)
        IF FindKeyword(masked$, q, "BOOLEAN") = q THEN
            newln$ = newln$ + MID$(ln$, lastPos, p - lastPos) + "As Integer"
            lastPos = q + 7
            changed = 1
            p = FindKeyword(masked$, lastPos, "AS")
        ELSE
            p = FindKeyword(masked$, p + 2, "AS")
        END IF
    LOOP
    IF changed = 1 THEN
        newln$ = newln$ + MID$(ln$, lastPos)
        ln$ = newln$
        AddLine warnings$(), warnCount, "line " + LTRIM$(RTRIM$(STR$(lineNo))) + ": converted " + CHR$(34) + "As Boolean" + CHR$(34) + " to " + CHR$(34) + "As Integer" + CHR$(34) + " (no BCX boolean type)"
    END IF
END SUB

SUB ReplaceEndif(ln$)
    DIM masked$, p, newln$, lastPos
    masked$ = MaskLine$(ln$)
    p = FindKeyword(masked$, 1, "ENDIF")
    IF p = 0 THEN EXIT SUB
    newln$ = ""
    lastPos = 1
    DO WHILE p <> 0
        newln$ = newln$ + MID$(ln$, lastPos, p - lastPos) + "END IF"
        lastPos = p + 5
        p = FindKeyword(masked$, lastPos, "ENDIF")
    LOOP
    newln$ = newln$ + MID$(ln$, lastPos)
    ln$ = newln$
END SUB

' Replaces built-in vbXxx constants (vbCrLf, vbTab, ...) with their
' literal CHR$()/"" equivalents, since BCX has no VB intrinsic constants.
SUB ReplaceVbConstants(ln$, lineNo, warnings$(), warnCount)
    DIM names$(9), repl$(9), w, masked$, p, newln$, lastPos, changed
    names$(1) = "vbCrLf" : repl$(1) = "CHR$(13) + CHR$(10)"
    names$(2) = "vbNewLine" : repl$(2) = "CHR$(13) + CHR$(10)"
    names$(3) = "vbCr" : repl$(3) = "CHR$(13)"
    names$(4) = "vbLf" : repl$(4) = "CHR$(10)"
    names$(5) = "vbTab" : repl$(5) = "CHR$(9)"
    names$(6) = "vbNullString" : repl$(6) = CHR$(34) + CHR$(34)
    names$(7) = "vbNullChar" : repl$(7) = "CHR$(0)"
    names$(8) = "vbBack" : repl$(8) = "CHR$(8)"
    names$(9) = "vbFormFeed" : repl$(9) = "CHR$(12)"

    FOR w = 1 TO 9
        masked$ = MaskLine$(ln$)
        newln$ = ""
        lastPos = 1
        p = 1
        changed = 0
        DO
            p = FindKeyword(masked$, p, names$(w))
            IF p = 0 THEN EXIT DO
            newln$ = newln$ + MID$(ln$, lastPos, p - lastPos) + repl$(w)
            lastPos = p + LEN(names$(w))
            p = lastPos
            changed = 1
        LOOP
        IF changed = 1 THEN
            newln$ = newln$ + MID$(ln$, lastPos)
            ln$ = newln$
            AddLine warnings$(), warnCount, "line " + LTRIM$(RTRIM$(STR$(lineNo))) + ": replaced constant " + CHR$(34) + names$(w) + CHR$(34) + " with literal " + repl$(w)
        END IF
    NEXT w
END SUB

' Renames identifiers that would collide with C keywords once BCX
' compiles the code to C, skipping "As <Type>" usages (PrecededByAs)
' and the leading Const/Static statement keyword itself.
SUB RenameReserved(ln$, lineNo, warnings$(), warnCount)
    DIM words$(25), repl$(25), overlap(25), n, w
    DIM masked$, p, newln$, lastPos, changed, leadingWs, skip

    words$(1) = "int" : repl$(1) = "int_var" : overlap(1) = 0
    words$(2) = "char" : repl$(2) = "char_var" : overlap(2) = 0
    words$(3) = "float" : repl$(3) = "float_var" : overlap(3) = 0
    words$(4) = "double" : repl$(4) = "double_var" : overlap(4) = 0
    words$(5) = "void" : repl$(5) = "void_var" : overlap(5) = 0
    words$(6) = "struct" : repl$(6) = "struct_var" : overlap(6) = 0
    words$(7) = "class" : repl$(7) = "class_var" : overlap(7) = 0
    words$(8) = "union" : repl$(8) = "union_var" : overlap(8) = 0
    words$(9) = "enum" : repl$(9) = "enum_var" : overlap(9) = 0
    words$(10) = "static" : repl$(10) = "static_var" : overlap(10) = 1
    words$(11) = "auto" : repl$(11) = "auto_var" : overlap(11) = 0
    words$(12) = "register" : repl$(12) = "register_var" : overlap(12) = 0
    words$(13) = "const" : repl$(13) = "const_var" : overlap(13) = 1
    words$(14) = "signed" : repl$(14) = "signed_var" : overlap(14) = 0
    words$(15) = "unsigned" : repl$(15) = "unsigned_var" : overlap(15) = 0
    words$(16) = "sizeof" : repl$(16) = "sizeof_var" : overlap(16) = 0
    words$(17) = "typedef" : repl$(17) = "typedef_var" : overlap(17) = 0
    words$(18) = "extern" : repl$(18) = "extern_var" : overlap(18) = 0
    words$(19) = "volatile" : repl$(19) = "volatile_var" : overlap(19) = 0
    words$(20) = "switch" : repl$(20) = "switch_var" : overlap(20) = 0
    words$(21) = "default" : repl$(21) = "default_var" : overlap(21) = 0
    words$(22) = "break" : repl$(22) = "break_var" : overlap(22) = 0
    words$(23) = "continue" : repl$(23) = "continue_var" : overlap(23) = 0
    words$(24) = "return" : repl$(24) = "return_var" : overlap(24) = 0
    words$(25) = "short" : repl$(25) = "short_var" : overlap(25) = 0
    n = 25

    FOR w = 1 TO n
        masked$ = MaskLine$(ln$)
        leadingWs = SkipSpacesFrom(ln$, 1)
        newln$ = ""
        lastPos = 1
        p = 1
        changed = 0
        DO
            p = FindKeyword(masked$, p, words$(w))
            IF p = 0 THEN EXIT DO
            skip = 0
            IF PrecededByAs(ln$, p) = 1 THEN skip = 1
            IF overlap(w) = 1 AND p = leadingWs THEN skip = 1
            IF skip = 0 THEN
                newln$ = newln$ + MID$(ln$, lastPos, p - lastPos) + repl$(w)
                lastPos = p + LEN(words$(w))
                changed = 1
            END IF
            p = p + LEN(words$(w))
        LOOP
        IF changed = 1 THEN
            newln$ = newln$ + MID$(ln$, lastPos)
            ln$ = newln$
            AddLine warnings$(), warnCount, "line " + LTRIM$(RTRIM$(STR$(lineNo))) + ": renamed identifier " + CHR$(34) + words$(w) + CHR$(34) + " -> " + CHR$(34) + repl$(w) + CHR$(34) + " (reserved in C, would break BCX-generated code)"
        END IF
    NEXT w
END SUB

' Checks every object-model / class-only / no-BCX-equivalent pattern.
' Unlike the QB port's FlagUnsupported$ (first match wins), a VB6 line
' can trip several of these at once (e.g. a Property Get returning a
' Variant), so every match gets its own TODO comment and warning.
SUB FlagUnsupported(masked$, ln$, lineNo, warnings$(), warnCount, outlines$(), outCount, indent$)
    IF IsClassModuleHeader(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "this looks like a VB6/VBA class module; BCX has no class model at all -- consider converting it by hand into a TYPE plus a set of free SUB/FUNCTION routines that take that TYPE as a parameter"
    END IF
    IF HasPropertyGetLetSet(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Property Get/Let/Set only exist inside class modules; BCX has no property syntax -- rewrite as a plain FUNCTION/SUB"
    END IF
    IF KeywordPresent(masked$, "IMPLEMENTS") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Implements (interface inheritance) has no BCX equivalent"
    END IF
    IF KeywordPresent(masked$, "WITH") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "With...End With has no BCX equivalent; expand every .Member reference inside the block by hand"
    END IF
    IF HasOnErrorResumeNext(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "On Error Resume Next has no BCX equivalent; use ON ERROR GOTO <label> and check for the specific failure instead"
    END IF
    IF HasErrDot(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "the Err object has no BCX equivalent"
    END IF
    IF KeywordPresent(masked$, "OPTIONAL") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Optional parameters aren't supported; BCX SUB/FUNCTION parameter lists are fixed-arity -- give the caller explicit overload(s) instead"
    END IF
    IF HasCreateObjectCall(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "CreateObject (COM late binding) is not supported by BCX"
    END IF
    IF HasSetNew(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Set ... = New (object instantiation) is not supported by BCX"
    END IF
    IF HasIsNothing(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "object references (Is Nothing) are not supported by BCX"
    END IF
    IF HasEqualsNothing(masked$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "object references (= Nothing) are not supported by BCX"
    END IF
    IF PhraseFollows(masked$, "AS", "VARIANT") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Variant has no BCX equivalent; give the variable a concrete type"
    END IF
    IF PhraseFollows(masked$, "AS", "OBJECT") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Object has no BCX equivalent -- BCX has no object references"
    END IF
    IF PhraseFollows(masked$, "AS", "CURRENCY") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Currency has no BCX equivalent; consider LONG (scaled) or DOUBLE"
    END IF
    IF PhraseFollows(masked$, "AS", "DATE") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Date has no BCX equivalent; store as a string or a numeric serial value"
    END IF
    IF PhraseFollows(masked$, "AS", "COLLECTION") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Collection has no BCX equivalent; use an array"
    END IF
    IF PhraseFollows(masked$, "AS", "DICTIONARY") = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "Dictionary has no BCX equivalent; use parallel arrays or a TYPE array"
    END IF
    IF IsConditionalCompilation(ln$) = 1 THEN
        EmitFlag outlines$(), outCount, warnings$(), warnCount, lineNo, indent$, "conditional compilation (#If/#Else/#End If) needs BCX's own $IF/$ENDIF metacommands or should be resolved by hand"
    END IF
END SUB

' The full conversion pipeline. A single source of truth used by both
' vb2bcx.bas (file I/O) and vb2bcx_test.bas (in-memory tests).
SUB ConvertProgram(lines$(), lineCount, outlines$(), outCount, warnings$(), warnCount)
    DIM i, ln$, masked$
    DIM enumMode, enumNextValue, enumPrefix$, enumIndent$, enumName$
    DIM memberName$, memberVal$, hasVal
    DIM masked2$, indent2$

    outCount = 0
    warnCount = 0
    enumMode = 0

    FOR i = 1 TO lineCount
        ln$ = lines$(i)
        masked$ = MaskLine$(ln$)

        IF enumMode = 1 THEN
            IF IsEnumEnd(masked$) = 1 THEN
                enumMode = 0
                AddLine warnings$(), warnCount, "line " + LTRIM$(RTRIM$(STR$(i))) + ": converted Enum " + enumPrefix$ + " to a CONST list"
            ELSEIF LTRIM$(RTRIM$(ln$)) = "" THEN
                ' blank line inside the Enum block -- skip it
            ELSE
                ParseEnumMember ln$, memberName$, memberVal$, hasVal
                IF memberName$ <> "" THEN
                    IF hasVal = 1 AND IsPlainInteger(memberVal$) = 1 THEN
                        enumNextValue = VAL(memberVal$)
                        AddLine outlines$(), outCount, enumIndent$ + "CONST " + enumPrefix$ + "_" + memberName$ + " = " + LTRIM$(RTRIM$(STR$(enumNextValue)))
                        enumNextValue = enumNextValue + 1
                    ELSEIF hasVal = 1 THEN
                        AddLine outlines$(), outCount, enumIndent$ + "CONST " + enumPrefix$ + "_" + memberName$ + " = " + memberVal$
                        enumNextValue = enumNextValue + 1
                    ELSE
                        AddLine outlines$(), outCount, enumIndent$ + "CONST " + enumPrefix$ + "_" + memberName$ + " = " + LTRIM$(RTRIM$(STR$(enumNextValue)))
                        enumNextValue = enumNextValue + 1
                    END IF
                ELSE
                    AddLine outlines$(), outCount, ln$
                END IF
            END IF
        ELSE
            IF IsEnumStart(masked$, enumName$) = 1 THEN
                enumMode = 1
                enumNextValue = 0
                enumPrefix$ = enumName$
                enumIndent$ = LeadingWhitespace$(ln$)
            ELSEIF IsAttributeLine(masked$) = 1 THEN
                ' dropped -- pure VB6 IDE export noise
            ELSE
                StripVisibility ln$
                ReplaceAsBoolean ln$, i, warnings$(), warnCount
                ReplaceEndif ln$
                ReplaceVbConstants ln$, i, warnings$(), warnCount

                masked2$ = MaskLine$(ln$)
                indent2$ = LeadingWhitespace$(ln$)
                FlagUnsupported masked2$, ln$, i, warnings$(), warnCount, outlines$(), outCount, indent2$

                RenameReserved ln$, i, warnings$(), warnCount

                AddLine outlines$(), outCount, ln$
            END IF
        END IF
    NEXT i
END SUB
