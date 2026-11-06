#!/usr/bin/env python3
"""
vb2bcx.py - VBA / VB6 -> BCX BASIC source converter

BCX BASIC is a procedural, QuickBASIC-flavored dialect with no object
model, so most of the interesting work here is telling apart what
translates mechanically from what simply has no BCX equivalent
(classes, properties, error objects, With blocks, ...) and flagging
the latter clearly instead of guessing.

What this pass does automatically:
  1. Strips VB6/VBA IDE-generated "Attribute VB_..." metadata lines --
     pure export noise, meaningless to BCX.
  2. Strips Public/Private/Friend visibility modifiers from Sub,
     Function, Dim, Const, Type and Enum declarations -- BCX has no
     module-level visibility model, and these keywords would just be
     unrecognized noise to it.
  3. Converts "As Boolean" to "As Integer" (BCX has no boolean type;
     True/False still work as the usual -1/0).
  4. Replaces built-in vbXxx constants (vbCrLf, vbTab, vbNullString,
     ...) with their literal CHR$()/"" equivalents.
  5. Converts "Enum ... End Enum" blocks into a plain CONST list with
     auto-incrementing values, since BCX has no Enum statement.
  6. Renames identifiers that collide with C keywords (int, char,
     class, static, const, ...), since BCX transpiles to C.
  7. Normalizes ENDIF -> END IF (rare in VB6, but cheap to handle).

What it flags with a ' TODO(BCX): comment instead of guessing:
  - Class module files (VERSION 1.0 CLASS header) -- BCX is purely
    procedural; there's no automatic path from a class to a BCX TYPE
    + free functions.
  - Property Get/Let/Set, Implements -- class-only constructs.
  - With ... End With blocks -- BCX has no With statement; every
    .Member reference inside the block needs to be expanded by hand.
  - On Error Resume Next, Err.* -- BCX only has ON ERROR GOTO <label>;
    there's no Err object to inspect afterwards.
  - Optional parameters -- BCX SUB/FUNCTION parameters are fixed-arity.
  - Set ... = New / CreateObject(...), "Is Nothing" / "= Nothing" --
    BCX has no object references or late binding.
  - As Variant / As Object / As Currency / As Date / As Collection /
    As Dictionary -- no equivalent dynamic, COM, or container type.
  - #If / #Else / #End If conditional compilation.

All string/comment content is masked out before any keyword or
identifier search runs, so nothing inside "..." strings or after a
' comment is ever matched or rewritten.

This is a best-effort source-to-source pass, not a full VB parser.
Always diff/review the output.

Usage:
    python3 vb2bcx.py input.bas [output.bas]
"""

import re
import sys
import os

# ---------------------------------------------------------------------------
# Constructs with no automatic BCX translation. Flag, don't rewrite.
# Checked against the masked (string/comment-free) line.
# ---------------------------------------------------------------------------
UNSUPPORTED_PATTERNS = [
    (r'^\s*VERSION\s+1\.0\s+CLASS\b',
     "this looks like a VB6/VBA class module; BCX has no class model at "
     "all -- consider converting it by hand into a TYPE plus a set of "
     "free SUB/FUNCTION routines that take that TYPE as a parameter"),
    (r'\bPROPERTY\s+(GET|LET|SET)\b',
     "Property Get/Let/Set only exist inside class modules; BCX has no "
     "property syntax -- rewrite as a plain FUNCTION/SUB"),
    (r'\bIMPLEMENTS\b',
     "Implements (interface inheritance) has no BCX equivalent"),
    (r'\bWITH\b',
     "With...End With has no BCX equivalent; expand every .Member "
     "reference inside the block by hand"),
    (r'\bON\s+ERROR\s+RESUME\s+NEXT\b',
     "On Error Resume Next has no BCX equivalent; use ON ERROR GOTO "
     "<label> and check for the specific failure instead"),
    (r'\bERR\s*\.', "the Err object has no BCX equivalent"),
    (r'\bOPTIONAL\b',
     "Optional parameters aren't supported; BCX SUB/FUNCTION parameter "
     "lists are fixed-arity -- give the caller explicit overload(s) instead"),
    (r'\bCREATEOBJECT\s*\(', "CreateObject (COM late binding) is not supported by BCX"),
    (r'\bSET\s+\w+\s*=\s*NEW\b', "Set ... = New (object instantiation) is not supported by BCX"),
    (r'\bIS\s+NOTHING\b', "object references (Is Nothing) are not supported by BCX"),
    (r'=\s*NOTHING\b', "object references (= Nothing) are not supported by BCX"),
    (r'\bAS\s+VARIANT\b', "Variant has no BCX equivalent; give the variable a concrete type"),
    (r'\bAS\s+OBJECT\b', "Object has no BCX equivalent -- BCX has no object references"),
    (r'\bAS\s+CURRENCY\b', "Currency has no BCX equivalent; consider LONG (scaled) or DOUBLE"),
    (r'\bAS\s+DATE\b', "Date has no BCX equivalent; store as a string or a numeric serial value"),
    (r'\bAS\s+COLLECTION\b', "Collection has no BCX equivalent; use an array"),
    (r'\bAS\s+DICTIONARY\b', "Dictionary has no BCX equivalent; use parallel arrays or a TYPE array"),
    (r'^\s*#\s*IF\b|^\s*#\s*ELSE\b|^\s*#\s*END\s*IF\b',
     "conditional compilation (#If/#Else/#End If) needs BCX's own "
     "$IF/$ENDIF metacommands or should be resolved by hand"),
]

# Identifiers that are safe in VB but clash with C keywords -- BCX
# transpiles to C, so a variable literally named "int" or "class"
# would break the generated C code.
C_RESERVED_RENAME = {
    'int': 'int_var', 'char': 'char_var', 'float': 'float_var',
    'double': 'double_var', 'void': 'void_var', 'struct': 'struct_var',
    'class': 'class_var', 'union': 'union_var', 'enum': 'enum_var',
    'static': 'static_var', 'auto': 'auto_var', 'register': 'register_var',
    'const': 'const_var', 'signed': 'signed_var', 'unsigned': 'unsigned_var',
    'sizeof': 'sizeof_var', 'typedef': 'typedef_var', 'extern': 'extern_var',
    'volatile': 'volatile_var', 'switch': 'switch_var', 'default': 'default_var',
    'break': 'break_var', 'continue': 'continue_var', 'return': 'return_var',
    'short': 'short_var',
}

# Built-in VB constants with no BCX equivalent -> literal replacement.
VB_CONST_MAP = {
    'vbCrLf': 'CHR$(13) + CHR$(10)',
    'vbNewLine': 'CHR$(13) + CHR$(10)',
    'vbCr': 'CHR$(13)',
    'vbLf': 'CHR$(10)',
    'vbTab': 'CHR$(9)',
    'vbNullString': '""',
    'vbNullChar': 'CHR$(0)',
    'vbBack': 'CHR$(8)',
    'vbFormFeed': 'CHR$(12)',
}

VISIBILITY_RE = re.compile(
    r'^(\s*)(PUBLIC|PRIVATE|FRIEND)\s+(?=(SUB|FUNCTION|DIM|CONST|TYPE|ENUM)\b)',
    re.IGNORECASE
)
ATTRIBUTE_RE = re.compile(r'^\s*ATTRIBUTE\s+VB_', re.IGNORECASE)
AS_BOOLEAN_RE = re.compile(r'\bAS\s+BOOLEAN\b', re.IGNORECASE)
ENDIF_RE = re.compile(r'\bENDIF\b', re.IGNORECASE)
ENUM_START_RE = re.compile(r'^(\s*)(?:PUBLIC\s+|PRIVATE\s+)?ENUM\s+(\w+)', re.IGNORECASE)
ENUM_END_RE = re.compile(r'^\s*END\s+ENUM\b', re.IGNORECASE)
ENUM_MEMBER_RE = re.compile(r'^\s*(\w+)\s*(?:=\s*(.+))?$')


def strip_strings_and_comments(line):
    """Return line with string-literal and comment contents blanked out,
    so keyword/identifier regexes don't match inside them. Same length
    as input so offsets stay usable."""
    out = []
    in_string = False
    i = 0
    n = len(line)
    while i < n:
        ch = line[i]
        if in_string:
            out.append('x' if ch != '"' else '"')
            if ch == '"':
                in_string = False
        else:
            if ch == '"':
                in_string = True
                out.append('"')
            elif ch == "'":
                out.append(' ' * (n - i))
                break
            else:
                out.append(ch)
        i += 1
    return ''.join(out)


def convert_vb_constants(line, lineno, warnings):
    masked = strip_strings_and_comments(line)
    for name, replacement in VB_CONST_MAP.items():
        pattern = re.compile(r'\b' + re.escape(name) + r'\b', re.IGNORECASE)
        new_line = []
        pos = 0
        changed = False
        for m in pattern.finditer(masked):
            if line[m.start():m.end()].lower() == name.lower():
                new_line.append(line[pos:m.start()])
                new_line.append(replacement)
                pos = m.end()
                changed = True
        if changed:
            new_line.append(line[pos:])
            line = ''.join(new_line)
            warnings.append(
                f'line {lineno}: replaced constant "{name}" with literal {replacement}'
            )
            masked = strip_strings_and_comments(line)
    return line


PRECEDING_WORD_RE = re.compile(r'(\w+)\s*$')

# Words that collide with C keywords but are ALSO real VB/BASIC statement
# keywords in their own right (Const declarations, Static locals). When
# one of these is the very first token on the line, it's being used as
# that BASIC keyword, not as an identifier, and must not be renamed.
BASIC_KEYWORD_OVERLAP = {'const', 'static'}


def _preceded_by_as(text_before_match):
    """True if the nearest preceding word (ignoring whitespace) is the
    keyword As -- meaning this occurrence is a VB type name in a
    declaration ("As Double"), not a variable/parameter identifier, and
    must NOT be renamed."""
    m = PRECEDING_WORD_RE.search(text_before_match)
    return bool(m) and m.group(1).upper() == 'AS'


def rename_reserved_words(line, lineno, warnings):
    for word, replacement in C_RESERVED_RENAME.items():
        pattern = re.compile(r'\b' + re.escape(word) + r'\b', re.IGNORECASE)
        masked = strip_strings_and_comments(line)
        leading_ws = len(line) - len(line.lstrip(' \t'))
        new_line = []
        pos = 0
        changed = False
        for m in pattern.finditer(masked):
            if line[m.start():m.end()].lower() != word:
                continue
            if _preceded_by_as(line[:m.start()]):
                # "As Double" etc -- a type name, not an identifier
                continue
            if word in BASIC_KEYWORD_OVERLAP and m.start() == leading_ws:
                # leading "Const"/"Static" statement keyword, not a
                # variable named const/static
                continue
            new_line.append(line[pos:m.start()])
            new_line.append(replacement)
            pos = m.end()
            changed = True
        if changed:
            new_line.append(line[pos:])
            line = ''.join(new_line)
            warnings.append(
                f'line {lineno}: renamed identifier "{word}" -> "{replacement}" '
                f'(reserved in C, would break BCX-generated code)'
            )
    return line


def flag_unsupported(line, lineno, warnings):
    """Returns a list of (reason) strings for every unsupported pattern
    matched on this line (a line can trip more than one, e.g. a Property
    Get inside a class module)."""
    masked = strip_strings_and_comments(line)
    reasons = []
    for pattern, reason in UNSUPPORTED_PATTERNS:
        if re.search(pattern, masked, re.IGNORECASE):
            reasons.append(reason)
            warnings.append(f'line {lineno}: {reason}')
    return reasons


def convert(source_lines):
    warnings = []
    out_lines = []
    enum_mode = False
    enum_next_value = 0
    enum_prefix = ''
    enum_indent = ''

    for lineno, raw in enumerate(source_lines, start=1):
        line = raw.rstrip('\n')
        masked = strip_strings_and_comments(line)

        # --- Enum ... End Enum -> CONST list -------------------------------
        if enum_mode:
            if ENUM_END_RE.match(masked):
                enum_mode = False
                warnings.append(
                    f'line {lineno}: converted Enum {enum_prefix} to a CONST list'
                )
                continue
            member = ENUM_MEMBER_RE.match(line.strip())
            if line.strip() == '':
                continue
            if member:
                name = member.group(1)
                explicit = member.group(2)
                if explicit:
                    try:
                        enum_next_value = int(explicit.strip())
                    except ValueError:
                        # non-numeric initializer (rare); keep the raw
                        # expression rather than guessing
                        out_lines.append(
                            f'{enum_indent}CONST {enum_prefix}_{name} = {explicit.strip()}'
                        )
                        enum_next_value += 1
                        continue
                out_lines.append(f'{enum_indent}CONST {enum_prefix}_{name} = {enum_next_value}')
                enum_next_value += 1
                continue
            # unrecognized line inside an Enum block -- keep it verbatim
            # rather than silently dropping it
            out_lines.append(line)
            continue

        enum_start = ENUM_START_RE.match(masked)
        if enum_start:
            enum_mode = True
            enum_next_value = 0
            enum_prefix = enum_start.group(2)
            enum_indent = enum_start.group(1)
            continue

        # --- strip VB6 IDE "Attribute VB_..." export noise ------------------
        if ATTRIBUTE_RE.match(masked):
            continue

        # --- strip Public/Private/Friend visibility modifiers --------------
        vis = VISIBILITY_RE.match(masked)
        if vis:
            line = line[:vis.start(2)] + line[vis.end():]
            masked = strip_strings_and_comments(line)

        # --- As Boolean -> As Integer ---------------------------------------
        if AS_BOOLEAN_RE.search(masked):
            line = AS_BOOLEAN_RE.sub('As Integer', line)
            warnings.append(f'line {lineno}: converted "As Boolean" to "As Integer" (no BCX boolean type)')
            masked = strip_strings_and_comments(line)

        # --- ENDIF -> END IF --------------------------------------------------
        if ENDIF_RE.search(masked):
            line = ENDIF_RE.sub('END IF', line)

        # --- vbXxx constants -> literals -------------------------------------
        line = convert_vb_constants(line, lineno, warnings)

        # --- flag unsupported / needs-manual-porting constructs ---------------
        # (checked BEFORE renaming, since e.g. "VERSION 1.0 CLASS" or a
        # "Property" line must be matched on its original keywords, not
        # after "class"/other reserved words have already been renamed)
        reasons = flag_unsupported(line, lineno, warnings)
        if reasons:
            indent = line[:len(line) - len(line.lstrip(' \t'))]
            for reason in reasons:
                out_lines.append(f"{indent}' TODO(BCX): {reason}")

        # --- rename identifiers colliding with C keywords ---------------------
        line = rename_reserved_words(line, lineno, warnings)

        out_lines.append(line)

    return out_lines, warnings


HEADER = """' ---------------------------------------------------------------
' Converted from VBA/VB6 to BCX BASIC by vb2bcx.py
' Review every '  TODO(BCX): comment below before compiling with BCX.
' BCX has no class/object model, so anything class-, property-, With-,
' or Err-related below is flagged rather than guessed at. Everything
' else (visibility modifiers, vbXxx constants, Enum blocks, Boolean
' types, C-keyword identifier clashes) has been rewritten automatically.
' ---------------------------------------------------------------

"""


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    in_path = sys.argv[1]
    if not os.path.isfile(in_path):
        print(f"error: no such file: {in_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        base, ext = os.path.splitext(in_path)
        out_path = f'{base}_bcx{ext or ".bas"}'

    with open(in_path, 'r', errors='replace') as f:
        source_lines = f.readlines()

    converted, warnings = convert(source_lines)

    with open(out_path, 'w') as f:
        f.write(HEADER)
        f.write('\n'.join(converted))
        f.write('\n')

    print(f'Wrote {out_path}')
    if warnings:
        print(f'\n{len(warnings)} item(s) need your attention:')
        for w in warnings:
            print(f'  - {w}')
    else:
        print('No flagged constructs -- looks like a clean, direct port.')


if __name__ == '__main__':
    main()
