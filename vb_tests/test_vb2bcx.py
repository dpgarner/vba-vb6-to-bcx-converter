"""
Unit tests for vb2bcx.py (VBA/VB6 -> BCX BASIC converter).

Run with:  pytest tests/ -v
"""
import os
import sys
import subprocess

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import vb2bcx  # noqa: E402


# ---------------------------------------------------------------------------
# strip_strings_and_comments
# ---------------------------------------------------------------------------

class TestStripStringsAndComments:

    def test_blanks_out_double_quoted_string_contents(self):
        line = 'x = "vbCrLf is not a real constant here"'
        masked = vb2bcx.strip_strings_and_comments(line)
        assert 'vbCrLf' not in masked

    def test_blanks_out_apostrophe_comment_to_end_of_line(self):
        line = "x = 1 ' mentions int and With and Err"
        masked = vb2bcx.strip_strings_and_comments(line)
        assert 'int' not in masked
        assert 'With' not in masked
        assert 'Err' not in masked

    def test_preserves_line_length(self):
        line = 'x = "hi" \' trailing comment'
        masked = vb2bcx.strip_strings_and_comments(line)
        assert len(masked) == len(line)

    def test_code_outside_strings_and_comments_is_untouched(self):
        line = 'If x = 1 Then Debug.Print "yes"'
        masked = vb2bcx.strip_strings_and_comments(line)
        assert masked.startswith('If x = 1 Then Debug.Print "')


# ---------------------------------------------------------------------------
# Attribute VB_... stripping
# ---------------------------------------------------------------------------

class TestAttributeStripping:

    def test_attribute_vb_name_is_dropped(self):
        out, _ = vb2bcx.convert(['Attribute VB_Name = "Module1"', 'x = 1'])
        assert out == ['x = 1']

    def test_attribute_lines_are_case_insensitive(self):
        out, _ = vb2bcx.convert(['attribute VB_Exposed = False'])
        assert out == []

    def test_non_attribute_line_is_kept(self):
        out, _ = vb2bcx.convert(['AttributeStyle = 1'])
        assert out == ['AttributeStyle = 1']


# ---------------------------------------------------------------------------
# Public/Private/Friend visibility stripping
# ---------------------------------------------------------------------------

class TestVisibilityStripping:

    def test_strips_public_sub(self):
        out, _ = vb2bcx.convert(['Public Sub Foo()'])
        assert out == ['Sub Foo()']

    def test_strips_private_function(self):
        out, _ = vb2bcx.convert(['Private Function Bar() As Long'])
        assert out == ['Function Bar() As Long']

    def test_strips_friend_dim(self):
        out, _ = vb2bcx.convert(['Friend Dim x As Integer'])
        assert out == ['Dim x As Integer']

    def test_strips_public_const(self):
        out, _ = vb2bcx.convert(['Public Const Pi As Double = 3.14'])
        assert out == ['Const Pi As Double = 3.14']

    def test_leaves_bare_dim_unchanged(self):
        out, _ = vb2bcx.convert(['Dim x As Integer'])
        assert out == ['Dim x As Integer']

    def test_does_not_strip_private_used_as_identifier_prefix(self):
        # "Privateer" must not have its "Private" prefix eaten
        out, _ = vb2bcx.convert(['Privateer = 1'])
        assert out == ['Privateer = 1']


# ---------------------------------------------------------------------------
# As Boolean -> As Integer
# ---------------------------------------------------------------------------

class TestBooleanConversion:

    def test_as_boolean_becomes_as_integer(self):
        out, warnings = vb2bcx.convert(['Dim flag As Boolean'])
        assert out == ['Dim flag As Integer']
        assert any('Boolean' in w for w in warnings)

    def test_case_insensitive(self):
        out, _ = vb2bcx.convert(['Dim flag as boolean'])
        assert out == ['Dim flag As Integer']


# ---------------------------------------------------------------------------
# vbXxx constant replacement
# ---------------------------------------------------------------------------

class TestVbConstants:

    def test_replaces_vbcrlf(self):
        out, warnings = vb2bcx.convert(['msg = "a" & vbCrLf & "b"'])
        assert out == ['msg = "a" & CHR$(13) + CHR$(10) & "b"']
        assert any('vbCrLf' in w for w in warnings)

    def test_replaces_vbtab(self):
        out, _ = vb2bcx.convert(['x = vbTab'])
        assert out == ['x = CHR$(9)']

    def test_replaces_vbnullstring(self):
        out, _ = vb2bcx.convert(['x = vbNullString'])
        assert out == ['x = ""']

    def test_does_not_replace_inside_string_literal(self):
        out, _ = vb2bcx.convert(['x = "vbCrLf"'])
        assert out == ['x = "vbCrLf"']

    def test_multiple_constants_on_one_line(self):
        out, _ = vb2bcx.convert(['x = vbTab & vbCr & vbLf'])
        assert out == ['x = CHR$(9) & CHR$(13) & CHR$(10)']


# ---------------------------------------------------------------------------
# Enum -> CONST list
# ---------------------------------------------------------------------------

class TestEnumConversion:

    def test_simple_enum_gets_sequential_values(self):
        src = ['Enum Color', '    Red', '    Green', '    Blue', 'End Enum']
        out, warnings = vb2bcx.convert(src)
        assert out == [
            'CONST Color_Red = 0',
            'CONST Color_Green = 1',
            'CONST Color_Blue = 2',
        ]
        assert any('Enum Color' in w for w in warnings)

    def test_explicit_value_resets_the_counter(self):
        src = ['Enum Color', '    Red', '    Blue = 10', '    Yellow', 'End Enum']
        out, _ = vb2bcx.convert(src)
        assert out == [
            'CONST Color_Red = 0',
            'CONST Color_Blue = 10',
            'CONST Color_Yellow = 11',
        ]

    def test_public_enum_is_recognized(self):
        src = ['Public Enum Direction', '    North', 'End Enum']
        out, _ = vb2bcx.convert(src)
        assert out == ['CONST Direction_North = 0']

    def test_blank_lines_inside_enum_are_skipped(self):
        src = ['Enum Color', '    Red', '', '    Green', 'End Enum']
        out, _ = vb2bcx.convert(src)
        assert out == ['CONST Color_Red = 0', 'CONST Color_Green = 1']


# ---------------------------------------------------------------------------
# ENDIF normalization
# ---------------------------------------------------------------------------

class TestEndifNormalization:

    def test_endif_becomes_end_if(self):
        out, _ = vb2bcx.convert(['ENDIF'])
        assert out == ['END IF']

    def test_endif_inside_string_is_untouched(self):
        out, _ = vb2bcx.convert(['x = "ENDIF"'])
        assert out == ['x = "ENDIF"']


# ---------------------------------------------------------------------------
# C reserved word renaming
# ---------------------------------------------------------------------------

class TestReservedWordRenaming:

    def test_renames_int_variable(self):
        out, warnings = vb2bcx.convert(['int = 5'])
        assert out == ['int_var = 5']
        assert any('int' in w and 'int_var' in w for w in warnings)

    def test_does_not_rename_type_name_after_as(self):
        out, _ = vb2bcx.convert(['Dim x As Double'])
        assert out == ['Dim x As Double']

    def test_does_not_rename_type_name_in_function_return(self):
        out, _ = vb2bcx.convert(['Function Square(x As Double) As Double'])
        assert 'As double_var' not in ' '.join(out)
        assert 'As Double' in out[0]

    def test_still_renames_identifier_named_double(self):
        # a variable literally named "double" (not preceded by As) should
        # still be renamed -- only the type-name usage is protected
        out, _ = vb2bcx.convert(['double = 5'])
        assert out == ['double_var = 5']

    def test_does_not_rename_inside_string_literal(self):
        out, _ = vb2bcx.convert(['x = "int"'])
        assert out == ['x = "int"']


# ---------------------------------------------------------------------------
# Unsupported construct flagging
# ---------------------------------------------------------------------------

class TestUnsupportedFlags:

    def test_class_module_header_is_flagged(self):
        _, warnings = vb2bcx.convert(['VERSION 1.0 CLASS'])
        assert any('class module' in w for w in warnings)

    def test_property_get_is_flagged(self):
        out, warnings = vb2bcx.convert(['Public Property Get Name() As String'])
        assert any("TODO(BCX)" in line for line in out)
        assert any('Property' in w for w in warnings)

    def test_with_block_is_flagged(self):
        out, warnings = vb2bcx.convert(['With Application', '.Visible = True', 'End With'])
        assert sum('TODO(BCX)' in line for line in out) == 2
        assert any('With' in w for w in warnings)

    def test_on_error_resume_next_is_flagged(self):
        _, warnings = vb2bcx.convert(['On Error Resume Next'])
        assert any('On Error Resume Next' in w for w in warnings)

    def test_err_object_is_flagged(self):
        _, warnings = vb2bcx.convert(['If Err.Number <> 0 Then'])
        assert any('Err object' in w for w in warnings)

    def test_optional_parameter_is_flagged(self):
        _, warnings = vb2bcx.convert(['Sub Foo(Optional x As Integer = 1)'])
        assert any('Optional' in w for w in warnings)

    def test_createobject_is_flagged(self):
        _, warnings = vb2bcx.convert(['Set fso = CreateObject("Scripting.FileSystemObject")'])
        assert any('CreateObject' in w for w in warnings)

    def test_set_new_is_flagged(self):
        _, warnings = vb2bcx.convert(['Set c = New Collection'])
        assert any('New' in w for w in warnings)

    def test_is_nothing_is_flagged(self):
        _, warnings = vb2bcx.convert(['If obj Is Nothing Then'])
        assert any('Nothing' in w for w in warnings)

    def test_as_variant_is_flagged(self):
        _, warnings = vb2bcx.convert(['Dim x As Variant'])
        assert any('Variant' in w for w in warnings)

    def test_as_object_is_flagged(self):
        _, warnings = vb2bcx.convert(['Dim x As Object'])
        assert any('Object' in w for w in warnings)

    def test_conditional_compilation_is_flagged(self):
        _, warnings = vb2bcx.convert(['#If DEBUG Then'])
        assert any('conditional compilation' in w for w in warnings)

    def test_clean_line_produces_no_warnings(self):
        out, warnings = vb2bcx.convert(['x = x + 1'])
        assert warnings == []
        assert out == ['x = x + 1']


# ---------------------------------------------------------------------------
# End-to-end / CLI
# ---------------------------------------------------------------------------

class TestEndToEnd:

    def test_full_module_conversion(self):
        src = [
            'Attribute VB_Name = "Module1"',
            'Public Enum Color',
            '    Red',
            '    Green',
            'End Enum',
            '',
            'Public Sub Greet(name As String)',
            '    Dim int As Integer',
            '    msg = "hi" & vbCrLf & name',
            '    int = 1',
            'End Sub',
        ]
        out, warnings = vb2bcx.convert(src)
        joined = '\n'.join(out)

        assert 'CONST Color_Red = 0' in joined
        assert 'CONST Color_Green = 1' in joined
        assert joined.startswith('Sub Greet(name As String)') is False  # Attribute line gone, enum first
        assert 'Sub Greet(name As String)' in joined
        assert 'int_var As Integer' in joined
        assert 'CHR$(13) + CHR$(10)' in joined
        assert 'int_var = 1' in joined
        assert len(warnings) >= 3

    def test_cli_writes_default_output_file(self, tmp_path):
        script = os.path.join(os.path.dirname(__file__), '..', 'vb2bcx.py')
        in_file = tmp_path / "mod.bas"
        in_file.write_text('x = 1\n')

        result = subprocess.run(
            [sys.executable, script, str(in_file)],
            capture_output=True, text=True, cwd=tmp_path
        )
        out_file = tmp_path / "mod_bcx.bas"
        assert result.returncode == 0
        assert out_file.exists()
        assert 'x = 1' in out_file.read_text()

    def test_cli_writes_named_output_file(self, tmp_path):
        script = os.path.join(os.path.dirname(__file__), '..', 'vb2bcx.py')
        in_file = tmp_path / "mod.bas"
        in_file.write_text('On Error Resume Next\n')
        out_file = tmp_path / "custom.bas"

        result = subprocess.run(
            [sys.executable, script, str(in_file), str(out_file)],
            capture_output=True, text=True, cwd=tmp_path
        )
        assert result.returncode == 0
        assert out_file.exists()
        assert 'need your attention' in result.stdout

    def test_cli_errors_on_missing_input_file(self, tmp_path):
        script = os.path.join(os.path.dirname(__file__), '..', 'vb2bcx.py')
        result = subprocess.run(
            [sys.executable, script, str(tmp_path / "nope.bas")],
            capture_output=True, text=True
        )
        assert result.returncode == 1
        assert 'error' in result.stdout.lower()
