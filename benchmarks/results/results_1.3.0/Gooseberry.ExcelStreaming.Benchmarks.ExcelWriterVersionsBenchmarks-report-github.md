```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                      | Mean    | Error    | StdDev   | Allocated |
|---------------------------- |--------:|---------:|---------:|----------:|
| ExcelWriter_WithoutEscaping | 1.622 s | 0.0323 s | 0.0359 s |  10.27 KB |
| ExcelWriter_WithEscaping    | 1.856 s | 0.0272 s | 0.0227 s |  10.31 KB |
