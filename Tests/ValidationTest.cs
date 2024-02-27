using GenTimeSheet.Core;
using System.Reflection;

namespace Tests
{
    public class ValidationTest
    {
        [Fact]
        public void CheckDaysInMonthDoesNotContainValidationError()
        {
            var validation = new Validation("testFiles/1.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckDaysInMonth", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("DAYS_IN_MONTH_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.DoesNotContain(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckDaysInMonthContainsValidationError()
        {
            var validation = new Validation("testFiles/ContainsValidationErrors.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckDaysInMonth", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("DAYS_IN_MONTH_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.Contains(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public async Task CheckWeekendsColorDoesNotContainValidationError()
        {
            var validation = new Validation("testFiles/1.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckWeekendsColor", BindingFlags.NonPublic | BindingFlags.Instance);

            dynamic? task = methodInfo?.Invoke(validation, null);

            await task?.ConfigureAwait(false);

            var fieldInfo = typeof(Validation).
                GetField("WEEKENDS_COLOR_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.DoesNotContain(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public async Task CheckWeekendsColorContainsValidationError()
        {
            var validation = new Validation("testFiles/ContainsValidationErrors.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckWeekendsColor", BindingFlags.NonPublic | BindingFlags.Instance);

            dynamic? task = methodInfo?.Invoke(validation, null);

            await task?.ConfigureAwait(false);

            var fieldInfo = typeof(Validation).
                GetField("WEEKENDS_COLOR_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.Contains(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckXsCountDoesNotContainValidationError()
        {
            var validation = new Validation("testFiles/1.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckXsCount", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("X_COUNT_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            string? s = fieldInfo?.GetValue(validation)?.ToString();

            bool hasError = false;

            foreach (var error in validation.ValidationErrors)
            {
                if (s != null && error.Contains(s))
                {
                    hasError = true;

                    break;
                }
            }

            Assert.False(hasError);
        }

        [Fact]
        public void CheckXsCountContainsValidationError()
        {
            var validation = new Validation("testFiles/ContainsValidationErrors.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckXsCount", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("X_COUNT_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.Contains(fieldInfo?.GetValue(validation) + "7", validation.ValidationErrors);
            Assert.Contains(fieldInfo?.GetValue(validation) + "29", validation.ValidationErrors);
        }

        [Fact]
        public void CheckWeekendsWithEightsDoesNotContainValidationError()
        {
            var validation = new Validation("testFiles/1.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckWeekendsWithEights", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("WEEKENDS_WITH_EIGHTS_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.DoesNotContain(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckWeekendsWithEightsContainsValidationError()
        {
            var validation = new Validation("testFiles/ContainsValidationErrors.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckWeekendsWithEights", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("WEEKENDS_WITH_EIGHTS_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.Contains(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckFirstDayDoesNotContainValidationError()
        {
            var validation = new Validation("testFiles/1.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckFirstDay", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("FIRST_DAY_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.DoesNotContain(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckFirstDayContainsValidationError()
        {
            var validation = new Validation("testFiles/ContainsValidationErrors.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckFirstDay", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("FIRST_DAY_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.Contains(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckOrderXsAndEightsDoesNotContainValidationError()
        {
            var validation = new Validation("testFiles/1.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckOrderXsAndEights", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("EIGHTS_AFTER_X_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.DoesNotContain(fieldInfo?.GetValue(validation), validation.ValidationErrors);
        }

        [Fact]
        public void CheckOrderXsAndEightsContainsValidationError()
        {
            var validation = new Validation("testFiles/ContainsValidationErrors.docx");

            var methodInfo = typeof(Validation).
                GetMethod("CheckOrderXsAndEights", BindingFlags.NonPublic | BindingFlags.Instance);

            methodInfo?.Invoke(validation, null);

            var fieldInfo = typeof(Validation).
                GetField("EIGHTS_AFTER_X_ERROR", BindingFlags.NonPublic | BindingFlags.Static);

            Assert.Contains(fieldInfo?.GetValue(validation) + "3", validation.ValidationErrors);
            Assert.Contains(fieldInfo?.GetValue(validation) + "24", validation.ValidationErrors);
        }
    }
}
