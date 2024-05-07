```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean           | Error        | StdDev       | Version | Gen0   | Allocated |
|---------------------------- |---------- |---------------:|-------------:|-------------:|-------- |-------:|----------:|
| ExcelWriter_RealWorldReport | 100       |       373.9 μs |      6.29 μs |      5.88 μs | 1.3.0   | 0.9766 |   8.85 KB |
| ExcelWriter_WithoutEscaping | 100       |     1,806.7 μs |     26.05 μs |     24.37 μs | 1.3.0   |      - |   8.99 KB |
| ExcelWriter_WithEscaping    | 100       |     1,897.2 μs |     36.57 μs |     34.20 μs | 1.3.0   |      - |   8.99 KB |
| ExcelWriter_RealWorldReport | 1000      |     3,372.8 μs |     40.70 μs |     33.99 μs | 1.3.0   |      - |   8.86 KB |
| ExcelWriter_WithoutEscaping | 1000      |    16,434.0 μs |    258.77 μs |    242.05 μs | 1.3.0   |      - |   9.01 KB |
| ExcelWriter_WithEscaping    | 1000      |    17,082.3 μs |    168.77 μs |    140.93 μs | 1.3.0   |      - |   9.01 KB |
| ExcelWriter_RealWorldReport | 10000     |    31,980.2 μs |    443.98 μs |    370.74 μs | 1.3.0   |      - |    8.9 KB |
| ExcelWriter_WithoutEscaping | 10000     |   162,550.0 μs |  1,680.26 μs |  1,489.51 μs | 1.3.0   |      - |   9.22 KB |
| ExcelWriter_WithEscaping    | 10000     |   169,705.7 μs |  1,923.78 μs |  1,799.50 μs | 1.3.0   |      - |   9.22 KB |
| ExcelWriter_RealWorldReport | 100000    |   311,851.3 μs |  2,257.40 μs |  2,001.13 μs | 1.3.0   |      - |   9.62 KB |
| ExcelWriter_RealWorldReport | 500000    | 1,551,753.2 μs |  5,599.00 μs |  4,371.33 μs | 1.3.0   |      - |   9.62 KB |
| ExcelWriter_WithoutEscaping | 100000    | 1,608,105.3 μs |  6,010.26 μs |  5,018.84 μs | 1.3.0   |      - |   9.75 KB |
| ExcelWriter_WithEscaping    | 100000    | 1,675,903.5 μs | 10,013.96 μs |  7,818.24 μs | 1.3.0   |      - |   9.75 KB |
| ExcelWriter_WithoutEscaping | 500000    | 7,858,517.2 μs | 40,653.12 μs | 36,037.93 μs | 1.3.0   |      - |   9.75 KB |
| ExcelWriter_WithEscaping    | 500000    | 8,475,834.5 μs | 51,801.36 μs | 48,455.02 μs | 1.3.0   |      - |   9.75 KB |
