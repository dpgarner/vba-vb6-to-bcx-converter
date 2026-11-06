import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', '..'))

from behave import given, when, then  # noqa: E402
import vb2bcx  # noqa: E402


@given('the vb2bcx converter')
def step_impl(context):
    context.source_lines = []


@given('the source line "{line}"')
def step_impl(context, line):
    context.source_lines = [line.replace('\\"', '"')]


@given('the source program:')
def step_impl(context):
    context.source_lines = context.text.strip('\n').split('\n')


@when('I convert the program')
def step_impl(context):
    context.output, context.warnings = vb2bcx.convert(context.source_lines)
    context.joined_output = '\n'.join(context.output)


@then('the output should be empty')
def step_impl(context):
    assert context.output == [], f'expected empty output, got: {context.output}'


@then('the output should contain "{text}"')
def step_impl(context, text):
    text = text.replace('\\"', '"')
    assert text in context.joined_output, (
        f'expected to find {text!r} in output:\n{context.joined_output}'
    )


@then('the output should not contain "{text}"')
def step_impl(context, text):
    text = text.replace('\\"', '"')
    assert text not in context.joined_output, (
        f'did not expect to find {text!r} in output:\n{context.joined_output}'
    )


@then('the output should contain a "{marker}" comment')
def step_impl(context, marker):
    assert any(marker in line for line in context.output), (
        f'expected a comment containing {marker!r} in output:\n{context.joined_output}'
    )


@then('at least {n:d} "{marker}" comments should appear in the output')
def step_impl(context, n, marker):
    count = sum(marker in line for line in context.output)
    assert count >= n, (
        f'expected at least {n} lines containing {marker!r}, found {count}:\n{context.joined_output}'
    )


@then('a warning mentioning "{text}" should be reported')
def step_impl(context, text):
    assert any(text in w for w in context.warnings), (
        f'expected a warning mentioning {text!r}, got: {context.warnings}'
    )


@then('no warning should be reported')
def step_impl(context):
    assert context.warnings == [], f'expected no warnings, got: {context.warnings}'


@then('at least {n:d} warnings should be reported')
def step_impl(context, n):
    assert len(context.warnings) >= n, (
        f'expected at least {n} warnings, got {len(context.warnings)}: {context.warnings}'
    )
