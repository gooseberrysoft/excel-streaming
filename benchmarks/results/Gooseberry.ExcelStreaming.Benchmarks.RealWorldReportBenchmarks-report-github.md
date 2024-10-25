```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5011/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.403
  [Host]     : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.10 (8.0.1024.46610), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | RowsCount | Mean           | Error        | StdDev        | Median         | Gen0   | Allocated |
|---------------------------- |---------- |---------------:|-------------:|--------------:|---------------:|-------:|----------:|
| ExcelWriter_RealWorldReport | 100       |       987.2 μs |     16.59 μs |      14.70 μs |       985.1 μs | 1.9531 |  14.94 KB |
| ExcelWriter_RealWorldReport | 1000      |     7,346.7 μs |    146.30 μs |     129.69 μs |     7,322.4 μs |      - |  13.83 KB |
| ExcelWriter_RealWorldReport | 10000     |    64,722.2 μs |  1,291.23 μs |   2,048.02 μs |    64,900.7 μs |      - |  15.18 KB |
| ExcelWriter_RealWorldReport | 100000    |   645,843.6 μs | 12,911.60 μs |  35,126.97 μs |   660,441.0 μs |      - |  17.13 KB |
| ExcelWriter_RealWorldReport | 500000    | 3,278,696.3 μs | 65,158.05 μs | 122,382.62 μs | 3,263,196.8 μs |      - |  74.73 KB |
