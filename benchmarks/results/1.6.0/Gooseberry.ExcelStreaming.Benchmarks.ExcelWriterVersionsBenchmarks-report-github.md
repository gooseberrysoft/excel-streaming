```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5011/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.403
  [Host]     : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                        | RowsCount | Mean         | Error      | StdDev      | Median       | Version | Allocated |
|------------------------------ |---------- |-------------:|-----------:|------------:|-------------:|-------- |----------:|
| ExcelWriter_WithEscaping_Utf8 | 100       |     1.220 ms |  0.0241 ms |   0.0422 ms |     1.207 ms | 1.5.1   |  14.54 KB |
| ExcelWriter_WithEscaping      | 100       |     1.289 ms |  0.0245 ms |   0.0282 ms |     1.289 ms | 1.5.1   |  14.63 KB |
| ExcelWriter_WithEscaping_Utf8 | 1000      |     8.912 ms |  0.1716 ms |   0.2461 ms |     8.864 ms | 1.5.1   |  15.16 KB |
| ExcelWriter_WithEscaping      | 1000      |     9.434 ms |  0.1881 ms |   0.4433 ms |     9.218 ms | 1.5.1   |  15.05 KB |
| ExcelWriter_WithEscaping_Utf8 | 10000     |    86.804 ms |  1.2222 ms |   1.0835 ms |    86.341 ms | 1.5.1   |  30.97 KB |
| ExcelWriter_WithEscaping      | 10000     |    90.631 ms |  1.7232 ms |   3.5200 ms |    89.803 ms | 1.5.1   |  29.74 KB |
| ExcelWriter_WithEscaping      | 100000    |   871.094 ms |  9.0834 ms |   7.0917 ms |   872.708 ms | 1.5.1   | 182.91 KB |
| ExcelWriter_WithEscaping_Utf8 | 100000    |   911.618 ms | 18.0196 ms |  35.5689 ms |   911.163 ms | 1.5.1   |    189 KB |
| ExcelWriter_WithEscaping_Utf8 | 500000    | 4,510.884 ms | 79.7964 ms | 121.8575 ms | 4,492.755 ms | 1.5.1   | 879.27 KB |
| ExcelWriter_WithEscaping      | 500000    | 4,643.406 ms | 91.3141 ms | 139.4463 ms | 4,601.858 ms | 1.5.1   | 854.91 KB |
