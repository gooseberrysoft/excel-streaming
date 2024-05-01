```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]   : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 6.0 : .NET 6.0.29 (6.0.2924.17105), X64 RyuJIT AVX2
  .NET 8.0 : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                           | Job      | Runtime  | Mean    | Error    | StdDev   | Median  | Allocated |
|--------------------------------- |--------- |--------- |--------:|---------:|---------:|--------:|----------:|
| ExcelWriter_WithEscaping         | .NET 7.0 | .NET 7.0 |      NA |       NA |       NA |      NA |        NA |
| ExcelWriter_WithoutEscaping      | .NET 7.0 | .NET 7.0 |      NA |       NA |       NA |      NA |        NA |
| ExcelWriter_WithEscaping_Utf8    | .NET 7.0 | .NET 7.0 |      NA |       NA |       NA |      NA |        NA |
| ExcelWriter_WithoutEscaping_Utf8 | .NET 7.0 | .NET 7.0 |      NA |       NA |       NA |      NA |        NA |
| ExcelWriter_WithoutEscaping_Utf8 | .NET 6.0 | .NET 6.0 | 1.476 s | 0.0121 s | 0.0101 s | 1.475 s |  16.09 KB |
| ExcelWriter_WithoutEscaping_Utf8 | .NET 8.0 | .NET 8.0 | 1.529 s | 0.0087 s | 0.0082 s | 1.528 s |  11.12 KB |
| ExcelWriter_WithoutEscaping      | .NET 8.0 | .NET 8.0 | 1.558 s | 0.0072 s | 0.0067 s | 1.558 s |  11.12 KB |
| ExcelWriter_WithEscaping_Utf8    | .NET 8.0 | .NET 8.0 | 1.563 s | 0.0128 s | 0.0114 s | 1.560 s |  11.12 KB |
| ExcelWriter_WithEscaping_Utf8    | .NET 6.0 | .NET 6.0 | 1.621 s | 0.0324 s | 0.0794 s | 1.592 s |  11.33 KB |
| ExcelWriter_WithoutEscaping      | .NET 6.0 | .NET 6.0 | 1.631 s | 0.0321 s | 0.0570 s | 1.607 s |  11.33 KB |
| ExcelWriter_WithEscaping         | .NET 8.0 | .NET 8.0 | 1.663 s | 0.0100 s | 0.0084 s | 1.663 s |  11.12 KB |
| ExcelWriter_WithEscaping         | .NET 6.0 | .NET 6.0 | 1.772 s | 0.0353 s | 0.0737 s | 1.738 s |  11.33 KB |

Benchmarks with issues:
  ExcelWriterVersionsBenchmarks.ExcelWriter_WithEscaping: .NET 7.0(Runtime=.NET 7.0)
  ExcelWriterVersionsBenchmarks.ExcelWriter_WithoutEscaping: .NET 7.0(Runtime=.NET 7.0)
  ExcelWriterVersionsBenchmarks.ExcelWriter_WithEscaping_Utf8: .NET 7.0(Runtime=.NET 7.0)
  ExcelWriterVersionsBenchmarks.ExcelWriter_WithoutEscaping_Utf8: .NET 7.0(Runtime=.NET 7.0)
