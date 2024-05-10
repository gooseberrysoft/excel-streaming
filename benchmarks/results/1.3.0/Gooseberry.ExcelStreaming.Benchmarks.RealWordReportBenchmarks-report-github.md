```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean           | Error        | StdDev       | Gen0   | Allocated |
|---------------------------- |---------- |---------------:|-------------:|-------------:|-------:|----------:|
| ExcelWriter_RealWorldReport | 100       |       410.2 μs |      3.40 μs |      3.18 μs | 0.9766 |   8.85 KB |
| ExcelWriter_RealWorldReport | 1000      |     3,280.4 μs |     40.67 μs |     33.96 μs |      - |   8.85 KB |
| ExcelWriter_RealWorldReport | 10000     |    30,993.8 μs |    505.37 μs |    447.99 μs |      - |   8.87 KB |
| ExcelWriter_RealWorldReport | 100000    |   303,899.9 μs |    791.38 μs |    617.86 μs |      - |   9.21 KB |
| ExcelWriter_RealWorldReport | 500000    | 1,508,141.7 μs | 10,515.48 μs |  9,836.18 μs |      - |   9.62 KB |
| ExcelWriter_RealWorldReport | 1000000   | 3,022,251.7 μs | 28,161.08 μs | 23,515.78 μs |      - |   9.62 KB |
