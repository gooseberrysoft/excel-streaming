```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method           | Mean       | Error    | StdDev   | Allocated  |
|----------------- |-----------:|---------:|---------:|-----------:|
| NullZip          |   746.9 ms | 10.39 ms |  8.68 ms |   23.23 KB |
| DefaultZip       | 1,043.9 ms |  4.21 ms |  3.73 ms |  284.73 KB |
| SharpZipLib      | 2,003.3 ms |  7.52 ms |  6.67 ms | 1558.29 KB |
| SharpCompressLib | 3,073.3 ms | 18.38 ms | 16.30 ms | 2772.63 KB |
