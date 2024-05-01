```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method           | Mean       | Error    | StdDev   | Allocated  |
|----------------- |-----------:|---------:|---------:|-----------:|
| NullZip          |   771.0 ms |  7.40 ms |  6.93 ms |    2.95 KB |
| DefaultZip       | 1,676.0 ms | 17.05 ms | 13.31 ms |   10.55 KB |
| SharpZipLib      | 2,610.8 ms | 14.04 ms | 13.13 ms | 1301.08 KB |
| SharpCompressLib | 3,672.2 ms | 33.65 ms | 28.10 ms | 2978.44 KB |
