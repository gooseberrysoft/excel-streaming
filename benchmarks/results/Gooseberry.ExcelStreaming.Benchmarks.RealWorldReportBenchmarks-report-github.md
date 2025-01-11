```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]     : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method          | RowsCount | Mean           | Error        | StdDev       | Gen0   | Allocated |
|---------------- |---------- |---------------:|-------------:|-------------:|-------:|----------:|
| RealWorldReport | 100       |       847.8 μs |     11.35 μs |     10.62 μs | 1.9531 |  13.47 KB |
| RealWorldReport | 1000      |     6,577.2 μs |     35.67 μs |     29.79 μs |      - |  14.02 KB |
| RealWorldReport | 10000     |    60,202.0 μs |    740.03 μs |    692.22 μs |      - |  19.98 KB |
| RealWorldReport | 100000    |   584,652.5 μs | 11,666.83 μs | 12,967.65 μs |      - |  78.59 KB |
| RealWorldReport | 500000    | 2,892,435.0 μs | 56,399.40 μs | 69,263.54 μs |      - |  344.7 KB |
