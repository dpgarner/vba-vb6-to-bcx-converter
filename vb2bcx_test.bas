' ===================================================================
' vb2bcx_test.bas - test suite for the vb2bcx converter, in BCX.
'
' Mirrors the checks in test_vb2bcx.py / vb2bcx.feature, run against
' the exact same code (via ConvertProgram and the other routines in
' vb2bcx_lib.bas -- nothing here is a re-implementation).
'
' Build & run:
'   bcx vb2bcx_test.bas
'   gcc vb2bcx_test.c -o vb2bcx_test.exe -O2
'   vb2bcx_test.exe
'
' Also NOT compiled/verified in this sandbox -- same caveat as
' vb2bcx.bas. Please compile it and send back whatever the compiler
' complains about.
' ===================================================================

DEFINT A-Z

'$INCLUDE: 'vb2bcx_lib.bas'

DECLARE SUB Check(cond, desc$)
DECLARE SUB CheckStr(actual$, expected$, desc$)
DECLARE SUB CheckContains(haystack$, needle$, desc$)
DECLARE SUB CheckNotContains(haystack$, needle$, desc$)
DECLARE FUNCTION JoinLines$(arr$(), cnt)
DECLARE SUB TestAttributeStripping
DECLARE SUB TestVisibilityStripping
DECLARE SUB TestBooleanConversion
DECLARE SUB TestVbConstants
DECLARE SUB TestEnumConversion
DECLARE SUB TestReservedWordRenaming
DECLARE SUB TestUnsupportedFlags
DECLARE SUB TestEndToEnd

DIM SHARED totalPass, totalFail, totalTests
totalPass = 0
totalFail = 0
totalTests = 0

PRINT "vb2bcx.bas test suite"
PRINT "======================"

TestAttributeStripping
TestVisibilityStripping
TestBooleanConversion
TestVbConstants
TestEnumConversion
TestReservedWordRenaming
TestUnsupportedFlags
TestEndToEnd

PRINT ""
PRINT "======================"
PRINT LTRIM$(RTRIM$(STR$(totalPass))) + "/" + LTRIM$(RTRIM$(STR$(totalTests))) + " checks passed"
IF totalFail > 0 THEN
    PRINT LTRIM$(RTRIM$(STR$(totalFail))) + " FAILED"
ELSE
    PRINT "ALL PASSED"
END IF

END

' ===================================================================
' tiny assertion framework
' ===================================================================

SUB Check(cond, desc$)
    totalTests = totalTests + 1
    IF cond = 1 THEN
        totalPass = totalPass + 1
        PRINT "  [PASS] " + desc$
    ELSE
        totalFail = totalFail + 1
        PRINT "  [FAIL] " + desc$
    END IF
END SUB

SUB CheckStr(actual$, expected$, desc$)
    IF actual$ = expected$ THEN
        Check 1, desc$
    ELSE
        Check 0, desc$ + " (expected " + CHR$(34) + expected$ + CHR$(34) + ", got " + CHR$(34) + actual$ + CHR$(34) + ")"
    END IF
END SUB

SUB CheckContains(haystack$, needle$, desc$)
    IF INSTR(haystack$, needle$) > 0 THEN
        Check 1, desc$
    ELSE
        Check 0, desc$ + " (expected to find " + CHR$(34) + needle$ + CHR$(34) + " in output)"
    END IF
END SUB

SUB CheckNotContains(haystack$, needle$, desc$)
    IF INSTR(haystack$, needle$) = 0 THEN
        Check 1, desc$
    ELSE
        Check 0, desc$ + " (did not expect to find " + CHR$(34) + needle$ + CHR$(34) + ")"
    END IF
END SUB

FUNCTION JoinLines$(arr$(), cnt)
    DIM r$, i
    r$ = ""
    FOR i = 1 TO cnt
        IF i > 1 THEN r$ = r$ + CHR$(10)
        r$ = r$ + arr$(i)
    NEXT i
    JoinLines$ = r$
END FUNCTION

' ===================================================================
' Attribute VB_... stripping -- mirrors TestAttributeStripping
' ===================================================================

SUB TestAttributeStripping
    PRINT ""
    PRINT "Attribute VB_... stripping"
    DIM src$(1), sc, out$(1), oc, w$(1), wc

    REDIM src$(1) : sc = 1
    src$(1) = "Attribute VB_Name = " + CHR$(34) + "Module1" + CHR$(34)
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    Check (oc = 0), "Attribute VB_Name line is dropped"

    REDIM src$(1) : sc = 1
    src$(1) = "attribute VB_Exposed = False"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    Check (oc = 0), "Attribute matching is case-insensitive"

    REDIM src$(1) : sc = 1
    src$(1) = "AttributeStyle = 1"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    Check (oc = 1), "a line merely starting with 'Attribute' as an identifier is kept"
END SUB

' ===================================================================
' Public/Private/Friend stripping -- mirrors TestVisibilityStripping
' ===================================================================

SUB TestVisibilityStripping
    PRINT ""
    PRINT "Public/Private/Friend visibility stripping"
    DIM ln$

    ln$ = "Public Sub Foo()"
    StripVisibility ln$
    CheckStr ln$, "Sub Foo()", "strips Public from Sub"

    ln$ = "Private Function Bar() As Long"
    StripVisibility ln$
    CheckStr ln$, "Function Bar() As Long", "strips Private from Function"

    ln$ = "Friend Dim x As Integer"
    StripVisibility ln$
    CheckStr ln$, "Dim x As Integer", "strips Friend from Dim"

    ln$ = "Public Const Pi As Double = 3.14"
    StripVisibility ln$
    CheckStr ln$, "Const Pi As Double = 3.14", "strips Public from Const"

    ln$ = "Dim x As Integer"
    StripVisibility ln$
    CheckStr ln$, "Dim x As Integer", "leaves a bare Dim unchanged"

    ln$ = "Privateer = 1"
    StripVisibility ln$
    CheckStr ln$, "Privateer = 1", "does not eat 'Private' out of a longer identifier"
END SUB

' ===================================================================
' As Boolean -> As Integer -- mirrors TestBooleanConversion
' ===================================================================

SUB TestBooleanConversion
    PRINT ""
    PRINT "As Boolean -> As Integer"
    DIM ln$, w$(1), wc

    ln$ = "Dim flag As Boolean" : wc = 0
    ReplaceAsBoolean ln$, 1, w$(), wc
    CheckStr ln$, "Dim flag As Integer", "As Boolean becomes As Integer"
    Check (wc = 1), "emits exactly one warning"

    ln$ = "Dim flag as boolean" : wc = 0
    ReplaceAsBoolean ln$, 1, w$(), wc
    CheckStr ln$, "Dim flag As Integer", "match is case-insensitive"
END SUB

' ===================================================================
' vbXxx constant replacement -- mirrors TestVbConstants
' ===================================================================

SUB TestVbConstants
    PRINT ""
    PRINT "vbXxx constant replacement"
    DIM ln$, w$(1), wc

    ln$ = "msg = " + CHR$(34) + "a" + CHR$(34) + " & vbCrLf & " + CHR$(34) + "b" + CHR$(34)
    wc = 0
    ReplaceVbConstants ln$, 1, w$(), wc
    CheckStr ln$, "msg = " + CHR$(34) + "a" + CHR$(34) + " & CHR$(13) + CHR$(10) & " + CHR$(34) + "b" + CHR$(34), "replaces vbCrLf"
    Check (wc = 1), "vbCrLf replacement emits one warning"

    ln$ = "x = vbTab" : wc = 0
    ReplaceVbConstants ln$, 1, w$(), wc
    CheckStr ln$, "x = CHR$(9)", "replaces vbTab"

    ln$ = "x = vbNullString" : wc = 0
    ReplaceVbConstants ln$, 1, w$(), wc
    CheckStr ln$, "x = " + CHR$(34) + CHR$(34), "replaces vbNullString with an empty string literal"

    ln$ = "x = " + CHR$(34) + "vbCrLf" + CHR$(34) : wc = 0
    ReplaceVbConstants ln$, 1, w$(), wc
    CheckStr ln$, "x = " + CHR$(34) + "vbCrLf" + CHR$(34), "does not replace inside a string literal"
END SUB

' ===================================================================
' Enum -> CONST list -- mirrors TestEnumConversion
' ===================================================================

SUB TestEnumConversion
    PRINT ""
    PRINT "Enum -> CONST list"
    DIM src$(4), sc, out$(1), oc, w$(1), wc, joined$

    REDIM src$(4) : sc = 4
    src$(1) = "Enum Color"
    src$(2) = "    Red"
    src$(3) = "    Green"
    src$(4) = "End Enum"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    Check (oc = 2), "two enum members produce two CONST lines"
    CheckStr out$(1), "    CONST Color_Red = 0", "first member gets value 0"
    CheckStr out$(2), "    CONST Color_Green = 1", "second member auto-increments"

    REDIM src$(4) : sc = 4
    src$(1) = "Enum Color"
    src$(2) = "    Red"
    src$(3) = "    Blue = 10"
    src$(4) = "End Enum"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(out$(), oc)
    CheckContains joined$, "CONST Color_Blue = 10", "explicit value is honored"

    REDIM src$(3) : sc = 3
    src$(1) = "Public Enum Direction"
    src$(2) = "    North"
    src$(3) = "End Enum"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    CheckStr out$(1), "CONST Direction_North = 0", "Public Enum is recognized"
END SUB

' ===================================================================
' RenameReserved -- mirrors TestReservedWordRenaming
' ===================================================================

SUB TestReservedWordRenaming
    PRINT ""
    PRINT "RenameReserved - C keyword identifier renaming"
    DIM ln$, w$(1), wc

    ln$ = "int = 5" : wc = 0
    RenameReserved ln$, 1, w$(), wc
    CheckStr ln$, "int_var = 5", "renames int variable"

    ln$ = "Dim x As Double" : wc = 0
    RenameReserved ln$, 1, w$(), wc
    CheckStr ln$, "Dim x As Double", "does not rename a type name after As"

    ln$ = "double = 5" : wc = 0
    RenameReserved ln$, 1, w$(), wc
    CheckStr ln$, "double_var = 5", "still renames a variable literally named double"

    ln$ = "x = " + CHR$(34) + "int" + CHR$(34) : wc = 0
    RenameReserved ln$, 1, w$(), wc
    CheckStr ln$, "x = " + CHR$(34) + "int" + CHR$(34), "does not rename inside a string literal"
END SUB

' ===================================================================
' FlagUnsupported -- mirrors TestUnsupportedFlags
' ===================================================================

SUB TestUnsupportedFlags
    PRINT ""
    PRINT "FlagUnsupported - object-model / class-only constructs"
    DIM src$(1), sc, out$(1), oc, w$(1), wc, joined$

    REDIM src$(1) : sc = 1 : src$(1) = "VERSION 1.0 CLASS"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "class module", "flags a class module header"

    REDIM src$(1) : sc = 1 : src$(1) = "Public Property Get Name() As String"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "Property", "flags Property Get"

    REDIM src$(3) : sc = 3
    src$(1) = "With Application"
    src$(2) = ".Visible = True"
    src$(3) = "End With"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(out$(), oc)
    Check (INSTR(joined$, "TODO(BCX)") > 0), "flags a With block"

    REDIM src$(1) : sc = 1 : src$(1) = "On Error Resume Next"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "Resume Next", "flags On Error Resume Next"

    REDIM src$(1) : sc = 1 : src$(1) = "If Err.Number <> 0 Then"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "Err object", "flags Err.Number"

    REDIM src$(1) : sc = 1 : src$(1) = "Sub Foo(Optional x As Integer = 1)"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "Optional", "flags Optional parameters"

    REDIM src$(1) : sc = 1
    src$(1) = "Set fso = CreateObject(" + CHR$(34) + "Scripting.Dictionary" + CHR$(34) + ")"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "CreateObject", "flags CreateObject"

    REDIM src$(1) : sc = 1 : src$(1) = "Set c = New Collection"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "New", "flags Set ... = New"

    REDIM src$(1) : sc = 1 : src$(1) = "If obj Is Nothing Then"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "Nothing", "flags Is Nothing"

    REDIM src$(1) : sc = 1 : src$(1) = "Dim x As Variant"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(w$(), wc)
    CheckContains joined$, "Variant", "flags As Variant"

    REDIM src$(1) : sc = 1 : src$(1) = "x = x + 1"
    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    Check (wc = 0), "a clean line produces no warnings"
END SUB

' ===================================================================
' ConvertProgram end-to-end -- mirrors TestEndToEnd
' ===================================================================

SUB TestEndToEnd
    PRINT ""
    PRINT "ConvertProgram - full module"
    DIM src$(10), sc, out$(1), oc, w$(1), wc, joined$

    REDIM src$(10) : sc = 10
    src$(1) = "Attribute VB_Name = " + CHR$(34) + "Module1" + CHR$(34)
    src$(2) = "Public Enum Color"
    src$(3) = "    Red"
    src$(4) = "    Green"
    src$(5) = "End Enum"
    src$(6) = ""
    src$(7) = "Public Sub Greet(name As String)"
    src$(8) = "    Dim int As Integer"
    src$(9) = "    msg = " + CHR$(34) + "hi" + CHR$(34) + " & vbCrLf & name"
    src$(10) = "End Sub"

    ConvertProgram src$(), sc, out$(), oc, w$(), wc
    joined$ = JoinLines$(out$(), oc)

    CheckContains joined$, "CONST Color_Red = 0", "full module: Enum converted"
    CheckContains joined$, "CONST Color_Green = 1", "full module: second Enum member auto-increments"
    CheckContains joined$, "Sub Greet(name As String)", "full module: Public stripped from Sub"
    CheckContains joined$, "int_var As Integer", "full module: reserved word renamed"
    CheckContains joined$, "CHR$(13) + CHR$(10)", "full module: vbCrLf replaced"
    CheckNotContains joined$, "Attribute", "full module: Attribute line dropped"
    Check (wc >= 3), "full module reports at least 3 warnings"
END SUB
