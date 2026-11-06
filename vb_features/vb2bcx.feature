Feature: Convert VBA/VB6 source to BCX BASIC
  As a developer porting VBA/VB6 code
  I want an automated converter
  So that mechanical syntax differences are handled for me and anything
  with no BCX equivalent -- classes, properties, With blocks, the Err
  object, and so on -- is clearly flagged rather than silently dropped
  or guessed at.

  Background:
    Given the vb2bcx converter

  Scenario: VB6 IDE export metadata is stripped
    Given the source line "Attribute VB_Name = \"Module1\""
    When I convert the program
    Then the output should be empty

  Scenario: Public/Private/Friend visibility modifiers are stripped
    Given the source line "Public Sub Foo()"
    When I convert the program
    Then the output should contain "Sub Foo()"
    And the output should not contain "Public"

  Scenario: A bare Const keyword is not mistaken for an identifier
    Given the source line "Public Const Pi As Double = 3.14"
    When I convert the program
    Then the output should contain "Const Pi As Double = 3.14"

  Scenario: Boolean is converted to Integer
    Given the source line "Dim flag As Boolean"
    When I convert the program
    Then the output should contain "Dim flag As Integer"
    And a warning mentioning "Boolean" should be reported

  Scenario: Built-in vb constants are replaced with literals
    Given the source line "msg = \"a\" & vbCrLf & \"b\""
    When I convert the program
    Then the output should contain "CHR$(13) + CHR$(10)"
    And a warning mentioning "vbCrLf" should be reported

  Scenario: An Enum block becomes a sequential CONST list
    Given the source program:
      """
      Enum Color
          Red
          Green
          Blue
      End Enum
      """
    When I convert the program
    Then the output should contain "CONST Color_Red = 0"
    And the output should contain "CONST Color_Green = 1"
    And the output should contain "CONST Color_Blue = 2"

  Scenario: An explicit Enum value resets the counter for later members
    Given the source program:
      """
      Enum Color
          Red
          Blue = 10
          Yellow
      End Enum
      """
    When I convert the program
    Then the output should contain "CONST Color_Blue = 10"
    And the output should contain "CONST Color_Yellow = 11"

  Scenario: A variable named after a C keyword is renamed
    Given the source line "int = 5"
    When I convert the program
    Then the output should contain "int_var = 5"
    And a warning mentioning "reserved in C" should be reported

  Scenario: A VB type name that collides with a C keyword is left alone
    Given the source line "Dim x As Double"
    When I convert the program
    Then the output should contain "Dim x As Double"
    And the output should not contain "double_var"

  Scenario Outline: Object-model and class-only constructs are flagged, not silently dropped
    Given the source line "<statement>"
    When I convert the program
    Then the output should contain a "TODO(BCX)" comment
    And a warning mentioning "<keyword>" should be reported

    Examples:
      | statement                                          | keyword       |
      | VERSION 1.0 CLASS                                  | class module  |
      | Public Property Get Name() As String                | Property      |
      | On Error Resume Next                                | Resume Next   |
      | If Err.Number <> 0 Then                             | Err object    |
      | Sub Foo(Optional x As Integer = 1)                  | Optional      |
      | Set fso = CreateObject("Scripting.Dictionary")       | CreateObject  |
      | Set c = New Collection                              | New           |
      | If obj Is Nothing Then                              | Nothing       |
      | Dim x As Variant                                    | Variant       |
      | Dim x As Object                                     | Object        |

  Scenario: A With block is flagged at both its start and end
    Given the source program:
      """
      With Application
      .Visible = True
      End With
      """
    When I convert the program
    Then at least 2 "TODO(BCX)" comments should appear in the output

  Scenario: Content inside string literals is never modified
    Given the source line "x = \"vbCrLf int Object\""
    When I convert the program
    Then the output should contain "x = \"vbCrLf int Object\""
    And no warning should be reported

  Scenario: Content inside comments is never modified
    Given the source line "' mentions int, vbCrLf, and With"
    When I convert the program
    Then the output should contain "mentions int, vbCrLf, and With"
    And no warning should be reported

  Scenario: A clean, already-compatible line produces no warnings
    Given the source line "x = x + 1"
    When I convert the program
    Then no warning should be reported

  Scenario: A full module converts and reports every issue found
    Given the source program:
      """
      Attribute VB_Name = "Module1"
      Public Enum Color
          Red
          Green
      End Enum

      Public Sub Greet(name As String)
          Dim int As Integer
          msg = "hi" & vbCrLf & name
          int = 1
      End Sub
      """
    When I convert the program
    Then the output should contain "CONST Color_Red = 0"
    And the output should contain "Sub Greet(name As String)"
    And the output should contain "int_var As Integer"
    And the output should contain "CHR$(13) + CHR$(10)"
    And at least 3 warnings should be reported
