```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]     : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                    | Mean     | Error    | StdDev   | Gen0       | Allocated   |
|-------------------------- |---------:|---------:|---------:|-----------:|------------:|
| InterpolatedStringHandler | 684.9 ms | 13.39 ms | 14.32 ms |          - |      7.5 KB |
| RegularString             | 959.5 ms | 18.83 ms | 25.78 ms | 81000.0000 | 500007.3 KB |
