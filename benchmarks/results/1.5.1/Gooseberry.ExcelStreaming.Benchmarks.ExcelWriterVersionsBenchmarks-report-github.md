```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                           | RowsCount | Mean           | Error        | StdDev       | Median         | Version | Gen0   | Allocated |
|--------------------------------- |---------- |---------------:|-------------:|-------------:|---------------:|-------- |-------:|----------:|
| ExcelWriter_RealWorldReport      | 100       |       364.0 μs |      4.18 μs |      3.70 μs |       363.4 μs | 1.5.1   | 1.9531 |  14.24 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 100       |     1,231.5 μs |     20.27 μs |     18.96 μs |     1,229.3 μs | 1.5.1   | 1.9531 |  14.83 KB |
| ExcelWriter_WithEscaping_Utf8    | 100       |     1,237.6 μs |     21.00 μs |     18.62 μs |     1,235.6 μs | 1.5.1   | 1.9531 |  14.85 KB |
| ExcelWriter_WithoutEscaping      | 100       |     1,245.1 μs |     23.59 μs |     28.98 μs |     1,251.0 μs | 1.5.1   |      - |  14.81 KB |
| ExcelWriter_WithEscaping         | 100       |     1,325.9 μs |     26.40 μs |     31.43 μs |     1,329.4 μs | 1.5.1   |      - |  14.91 KB |
| ExcelWriter_RealWorldReport      | 1000      |     2,486.6 μs |     25.72 μs |     24.06 μs |     2,481.5 μs | 1.5.1   |      - |  14.05 KB |
| ExcelWriter_WithEscaping         | 1000      |    11,279.0 μs |    221.63 μs |    207.31 μs |    11,312.5 μs | 1.5.1   |      - |  14.47 KB |
| ExcelWriter_WithoutEscaping      | 1000      |    11,638.8 μs |    228.18 μs |    234.33 μs |    11,639.0 μs | 1.5.1   |      - |  45.12 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 1000      |    12,031.6 μs |    237.94 μs |    429.05 μs |    11,809.0 μs | 1.5.1   |      - |  14.79 KB |
| ExcelWriter_WithEscaping_Utf8    | 1000      |    12,202.5 μs |    128.02 μs |    113.48 μs |    12,161.1 μs | 1.5.1   |      - |  14.51 KB |
| ExcelWriter_RealWorldReport      | 10000     |    24,381.8 μs |    230.65 μs |    215.75 μs |    24,333.7 μs | 1.5.1   |      - |  13.96 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 10000     |   107,012.7 μs |    894.90 μs |    837.09 μs |   107,178.1 μs | 1.5.1   |      - |  25.36 KB |
| ExcelWriter_WithEscaping_Utf8    | 10000     |   108,468.6 μs |  2,055.43 μs |  2,110.78 μs |   107,298.6 μs | 1.5.1   |      - | 281.42 KB |
| ExcelWriter_WithoutEscaping      | 10000     |   109,172.4 μs |  1,564.71 μs |  1,463.63 μs |   108,671.5 μs | 1.5.1   |      - |  25.13 KB |
| ExcelWriter_WithEscaping         | 10000     |   109,220.1 μs |  1,553.58 μs |  1,525.82 μs |   108,711.7 μs | 1.5.1   |      - |     25 KB |
| ExcelWriter_RealWorldReport      | 100000    |   231,758.1 μs |  3,387.18 μs |  3,002.64 μs |   231,507.9 μs | 1.5.1   |      - |  14.14 KB |
| ExcelWriter_WithEscaping_Utf8    | 100000    | 1,033,337.4 μs |  2,392.71 μs |  1,998.02 μs | 1,034,021.3 μs | 1.5.1   |      - | 140.73 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 100000    | 1,044,866.4 μs |  7,455.00 μs |  6,608.67 μs | 1,043,546.4 μs | 1.5.1   |      - | 136.69 KB |
| ExcelWriter_WithoutEscaping      | 100000    | 1,050,119.9 μs |  8,577.34 μs |  6,696.63 μs | 1,047,115.7 μs | 1.5.1   |      - | 153.45 KB |
| ExcelWriter_WithEscaping         | 100000    | 1,058,024.1 μs |  7,044.97 μs |  6,245.18 μs | 1,057,582.4 μs | 1.5.1   |      - | 136.88 KB |
| ExcelWriter_RealWorldReport      | 500000    | 1,105,868.4 μs |  6,314.20 μs |  5,906.30 μs | 1,103,461.8 μs | 1.5.1   |      - |  15.88 KB |
| ExcelWriter_WithEscaping_Utf8    | 500000    | 5,153,355.8 μs | 15,811.65 μs | 13,203.44 μs | 5,149,380.2 μs | 1.5.1   |      - | 653.32 KB |
| ExcelWriter_WithoutEscaping_Utf8 | 500000    | 5,178,344.4 μs | 10,893.08 μs |  9,096.22 μs | 5,179,011.3 μs | 1.5.1   |      - | 632.13 KB |
| ExcelWriter_WithoutEscaping      | 500000    | 5,213,166.0 μs | 10,240.66 μs |  9,078.08 μs | 5,211,837.5 μs | 1.5.1   |      - | 664.48 KB |
| ExcelWriter_WithEscaping         | 500000    | 5,281,583.5 μs | 68,896.26 μs | 64,445.61 μs | 5,309,193.7 μs | 1.5.1   |      - | 644.88 KB |
