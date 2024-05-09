```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                        | RowsCount | Mean          | Error      | StdDev     | Version | Allocated  |
|------------------------------ |---------- |--------------:|-----------:|-----------:|-------- |-----------:|
| ExcelWriter_WithEscaping_Utf8 | 100       |      1.545 ms |  0.0287 ms |  0.0268 ms | 1.5.1   |   15.18 KB |
| ExcelWriter_WithEscaping      | 100       |      1.582 ms |  0.0238 ms |  0.0185 ms | 1.5.1   |   15.34 KB |
| ExcelWriter_WithEscaping_Utf8 | 1000      |     11.570 ms |  0.2313 ms |  0.2375 ms | 1.5.1   |   15.76 KB |
| ExcelWriter_WithEscaping      | 1000      |     11.655 ms |  0.2206 ms |  0.2064 ms | 1.5.1   |   15.49 KB |
| ExcelWriter_WithEscaping_Utf8 | 10000     |    105.545 ms |  1.0024 ms |  0.8371 ms | 1.5.1   |   40.68 KB |
| ExcelWriter_WithEscaping      | 10000     |    106.743 ms |  1.1740 ms |  1.0981 ms | 1.5.1   |   39.51 KB |
| ExcelWriter_WithEscaping_Utf8 | 100000    |  1,033.685 ms |  2.8644 ms |  2.3919 ms | 1.5.1   |  286.36 KB |
| ExcelWriter_WithEscaping      | 100000    |  1,040.069 ms |  3.5074 ms |  3.1093 ms | 1.5.1   |   281.7 KB |
| ExcelWriter_WithEscaping_Utf8 | 500000    |  5,175.774 ms | 10.2418 ms |  9.0791 ms | 1.5.1   | 2371.66 KB |
| ExcelWriter_WithEscaping      | 500000    |  5,234.404 ms | 32.4442 ms | 30.3484 ms | 1.5.1   | 2365.66 KB |
| ExcelWriter_WithEscaping_Utf8 | 1000000   | 10,331.080 ms | 24.4270 ms | 19.0710 ms | 1.5.1   | 2750.41 KB |
| ExcelWriter_WithEscaping      | 1000000   | 10,441.862 ms | 50.0232 ms | 46.7917 ms | 1.5.1   | 2719.27 KB |
