```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method           | Mean    | Error    | StdDev   | Allocated  |
|----------------- |--------:|---------:|---------:|-----------:|
| DefaultZip       | 1.892 s | 0.0338 s | 0.0462 s |   11.12 KB |
| SharpZipLib      | 2.736 s | 0.0236 s | 0.0197 s | 1320.04 KB |
| SharpCompressLib | 3.715 s | 0.0620 s | 0.0550 s | 2981.14 KB |
