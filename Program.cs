using System.Diagnostics;

WordValidation wordValidation = new();
Stopwatch stopwatch = new Stopwatch();
//засекаем время начала операции
stopwatch.Start();
wordValidation.CheckDaysInMonth();
wordValidation.CheckWeekendsColor();
wordValidation.CheckXsCount();
wordValidation.CheckWeekendsWithEights();
wordValidation.CheckFirstDay();
wordValidation.CheckOrderXand8();
stopwatch.Stop();
//смотрим сколько миллисекунд было затрачено на выполнение
Console.WriteLine(stopwatch.ElapsedMilliseconds);