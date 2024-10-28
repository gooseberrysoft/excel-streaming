```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5011/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.403
  [Host]     : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                        | RowsCount | Mean         | Error      | StdDev     | Version | Allocated |
|------------------------------ |---------- |-------------:|-----------:|-----------:|-------- |----------:|
| ExcelWriter_WithEscaping_Utf8 | 100       |     1.348 ms |  0.0268 ms |  0.0357 ms | 1.5.1   |  14.44 KB |
| ExcelWriter_WithEscaping      | 100       |     1.450 ms |  0.0274 ms |  0.0305 ms | 1.5.1   |  14.55 KB |
| ExcelWriter_WithEscaping_Utf8 | 1000      |     9.497 ms |  0.0610 ms |  0.0541 ms | 1.5.1   |  14.89 KB |
| ExcelWriter_WithEscaping      | 1000      |     9.775 ms |  0.0951 ms |  0.0794 ms | 1.5.1   |  14.96 KB |
| ExcelWriter_WithEscaping      | 10000     |    86.495 ms |  0.6249 ms |  0.5540 ms | 1.5.1   |  30.52 KB |
| ExcelWriter_WithEscaping_Utf8 | 10000     |    87.441 ms |  0.9526 ms |  0.8910 ms | 1.5.1   |  31.44 KB |
| ExcelWriter_WithEscaping      | 100000    |   870.672 ms | 14.7892 ms | 13.8338 ms | 1.5.1   | 184.49 KB |
| ExcelWriter_WithEscaping_Utf8 | 100000    |   876.530 ms | 16.0430 ms | 14.2217 ms | 1.5.1   | 194.11 KB |
| ExcelWriter_WithEscaping      | 500000    | 4,560.587 ms | 89.2765 ms | 99.2305 ms | 1.5.1   | 922.64 KB |
| ExcelWriter_WithEscaping_Utf8 | 500000    | 4,608.393 ms | 49.2215 ms | 41.1022 ms | 1.5.1   | 917.72 KB |
