```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                           | Mean    | Error    | StdDev   | Allocated |
|--------------------------------- |--------:|---------:|---------:|----------:|
| ExcelWriter_WithoutEscaping_Utf8 | 1.556 s | 0.0309 s | 0.0444 s |  11.09 KB |
| ExcelWriter_WithoutEscaping      | 1.597 s | 0.0317 s | 0.0339 s |  11.09 KB |
| ExcelWriter_WithEscaping_Utf8    | 1.605 s | 0.0321 s | 0.0343 s |  11.09 KB |
| ExcelWriter_WithEscaping         | 1.840 s | 0.0272 s | 0.0241 s |  11.09 KB |
