```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]     : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method              | Mean       | Error     | StdDev    |
|-------------------- |-----------:|----------:|----------:|
| DateCustomFormatter |   5.968 ms | 0.0899 ms | 0.0841 ms |
| CustomFormatter     |  13.674 ms | 0.2596 ms | 0.2429 ms |
| DateDoubleFormatter |  75.714 ms | 0.9390 ms | 0.7841 ms |
| DoubleFormatter     | 123.004 ms | 1.4347 ms | 1.3420 ms |
