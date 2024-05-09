```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                   | RowsCount | Mean          | Error       | StdDev      | Version | Allocated |
|------------------------- |---------- |--------------:|------------:|------------:|-------- |----------:|
| ExcelWriter_WithEscaping | 100       |      2.023 ms |   0.0402 ms |   0.0508 ms | 1.3.0   |   8.99 KB |
| ExcelWriter_WithEscaping | 1000      |     17.450 ms |   0.2556 ms |   0.2390 ms | 1.3.0   |   9.01 KB |
| ExcelWriter_WithEscaping | 10000     |    168.661 ms |   1.0603 ms |   0.9400 ms | 1.3.0   |   9.22 KB |
| ExcelWriter_WithEscaping | 100000    |  1,690.950 ms |  11.8483 ms |  10.5032 ms | 1.3.0   |   9.75 KB |
| ExcelWriter_WithEscaping | 500000    |  8,364.383 ms | 115.4255 ms |  96.3855 ms | 1.3.0   |   9.75 KB |
| ExcelWriter_WithEscaping | 1000000   | 16,534.636 ms | 140.1426 ms | 131.0895 ms | 1.3.0   |   9.75 KB |
