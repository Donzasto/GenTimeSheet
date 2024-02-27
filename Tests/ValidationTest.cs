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
        public async void CheckWeekendsColorDoesNotContainValidationError()
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
    }
}
