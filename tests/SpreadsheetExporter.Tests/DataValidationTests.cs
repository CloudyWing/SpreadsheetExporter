using CloudyWing.SpreadsheetExporter;

namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class DataValidationTests {
    [Test]
    public void Constructor_ShouldSetDefaultValues() {
        DataValidation validation = new();

        Assert.Multiple(() => {
            Assert.That(validation.IsDropdownShown, Is.True);
            Assert.That(validation.IsBlankAllowed, Is.True);
            Assert.That(validation.IsErrorAlertShown, Is.True);
            Assert.That(validation.IsInputPromptShown, Is.False);
        });
    }

    [Test]
    public void ValidationType_ShouldBeSettable() {
        DataValidation validation = new() {
            ValidationType = DataValidationType.List
        };

        Assert.That(validation.ValidationType, Is.EqualTo(DataValidationType.List));
    }

    [Test]
    public void Operator_ShouldBeSettable() {
        DataValidation validation = new() {
            Operator = DataValidationOperator.Between
        };

        Assert.That(validation.Operator, Is.EqualTo(DataValidationOperator.Between));
    }

    [Test]
    public void Value1AndValue2_ShouldBeSettable() {
        DataValidation validation = new() {
            Value1 = 1,
            Value2 = 100
        };

        Assert.Multiple(() => {
            Assert.That(validation.Value1, Is.EqualTo(1));
            Assert.That(validation.Value2, Is.EqualTo(100));
        });
    }

    [Test]
    public void ListItems_ShouldBeSettable() {
        string[] items = new[] { "Item1", "Item2", "Item3" };
        DataValidation validation = new() {
            ListItems = items
        };

        Assert.That(validation.ListItems, Is.EqualTo(items));
    }

    [Test]
    public void Formula_ShouldBeSettable() {
        DataValidation validation = new() {
            Formula = "LEN(A1)<=10"
        };

        Assert.That(validation.Formula, Is.EqualTo("LEN(A1)<=10"));
    }

    [Test]
    public void ErrorTitleAndMessage_ShouldBeSettable() {
        DataValidation validation = new() {
            ErrorTitle = "Error",
            ErrorMessage = "Invalid value"
        };

        Assert.Multiple(() => {
            Assert.That(validation.ErrorTitle, Is.EqualTo("Error"));
            Assert.That(validation.ErrorMessage, Is.EqualTo("Invalid value"));
        });
    }

    [Test]
    public void PromptTitleAndMessage_ShouldBeSettable() {
        DataValidation validation = new() {
            PromptTitle = "Prompt",
            PromptMessage = "Please enter a value"
        };

        Assert.Multiple(() => {
            Assert.That(validation.PromptTitle, Is.EqualTo("Prompt"));
            Assert.That(validation.PromptMessage, Is.EqualTo("Please enter a value"));
        });
    }

    [Test]
    public void BooleanProperties_ShouldBeSettable() {
        DataValidation validation = new() {
            IsDropdownShown = false,
            IsBlankAllowed = false,
            IsErrorAlertShown = false,
            IsInputPromptShown = true
        };

        Assert.Multiple(() => {
            Assert.That(validation.IsDropdownShown, Is.False);
            Assert.That(validation.IsBlankAllowed, Is.False);
            Assert.That(validation.IsErrorAlertShown, Is.False);
            Assert.That(validation.IsInputPromptShown, Is.True);
        });
    }
}
