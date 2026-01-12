import pytest
import allure

@allure.title("Verify that 1 + 1 equals 2")
def test_sum_pass():
    """This test should pass."""
    assert 1 + 1 == 2


@allure.title("Verify that 2 + 2 equals 5 (intentional fail)")
def test_sum_fail():
    """This test will fail on purpose."""
    assert 2 + 2 == 5


@allure.title("Test with steps and attachments")
def test_with_steps_and_attachments():
    step_one()
    step_two()
    with open("sample_file.txt", "w") as f:
        f.write("Hello, Allure!")
    allure.attach.file("sample_file.txt", name="Sample file content", attachment_type=allure.attachment_type.TEXT)


@allure.step("Step One")
def step_one():
    # Imagine some logic here
    pass

@allure.step("Step Two")
def step_two():
    # Another piece of logic
    pass
