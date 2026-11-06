' ===================================================================
' vb2bcx.bas - VBA/VB6 -> BCX BASIC source converter
' BCX BASIC port of the original Python vb2bcx.py tool.
'
' NOTE ON TESTING: written and hand-traced but NOT compiled -- BCX is
' a Windows-targeted tool (BCX.exe + a C compiler such as MinGW) and
' neither is available in the sandbox this was written in. Compile it
' yourself and report anything that breaks:
'
'   bcx vb2bcx.bas
'   gcc vb2bcx.c -o vb2bcx.exe -O2
'
' Then run:
'   vb2bcx.exe input.bas [output.bas]
'
' All the actual conversion logic lives in vb2bcx_lib.bas (included
' below) so vb2bcx_test.bas can exercise the exact same code path via
' ConvertProgram, rather than a hand-copied twin.
'
' WHAT IT DOES AUTOMATICALLY (see vb2bcx_lib.bas header for the full
' list of what gets flagged instead):
'   1. Strips "Attribute VB_..." IDE export noise.
'   2. Strips Public/Private/Friend visibility modifiers.
'   3. Converts "As Boolean" to "As Integer".
'   4. Replaces built-in vbXxx constants with literal CHR$()/"" values.
'   5. Converts "Enum ... End Enum" into a sequential CONST list.
'   6. Renames identifiers colliding with C keywords.
'   7. Normalizes ENDIF -> END IF.
' ===================================================================

DEFINT A-Z

'$INCLUDE: 'vb2bcx_lib.bas'

' -------------------------------------------------------------------
' MAIN
' -------------------------------------------------------------------

DIM args$(1), argCount
ParseArgs COMMAND$, args$(), argCount

IF argCount < 1 THEN
    PRINT "Usage: vb2bcx.exe input.bas [output.bas]"
    END
END IF

DIM inPath$, outPath$
inPath$ = args$(1)
IF argCount >= 2 THEN
    outPath$ = args$(2)
ELSE
    outPath$ = DefaultOutputName$(inPath$)
END IF

ON ERROR GOTO OpenFailed
OPEN inPath$ FOR INPUT AS #1
ON ERROR GOTO 0

DIM lines$(1), lineCount, oneLine$
lineCount = 0
DO WHILE NOT EOF(1)
    LINE INPUT #1, oneLine$
    lineCount = lineCount + 1
    REDIM PRESERVE lines$(lineCount)
    lines$(lineCount) = oneLine$
LOOP
CLOSE #1

DIM outlines$(1), outCount, warnings$(1), warnCount, i
ConvertProgram lines$(), lineCount, outlines$(), outCount, warnings$(), warnCount

OPEN outPath$ FOR OUTPUT AS #2
PRINT #2, "' ---------------------------------------------------------------"
PRINT #2, "' Converted from VBA/VB6 to BCX BASIC by vb2bcx.bas"
PRINT #2, "' Review every '  TODO(BCX): comment below before compiling with BCX."
PRINT #2, "' BCX has no class/object model, so anything class-, property-,"
PRINT #2, "' With-, or Err-related below is flagged rather than guessed at."
PRINT #2, "' Everything else has been rewritten automatically."
PRINT #2, "' ---------------------------------------------------------------"
PRINT #2, ""
FOR i = 1 TO outCount
    PRINT #2, outlines$(i)
NEXT i
CLOSE #2

PRINT "Wrote " + outPath$
IF warnCount > 0 THEN
    PRINT ""
    PRINT LTRIM$(RTRIM$(STR$(warnCount))) + " item(s) need your attention:"
    FOR i = 1 TO warnCount
        PRINT "  - " + warnings$(i)
    NEXT i
ELSE
    PRINT "No flagged constructs -- looks like a clean, direct port."
END IF

END

OpenFailed:
    PRINT "error: no such file: " + inPath$
    END
