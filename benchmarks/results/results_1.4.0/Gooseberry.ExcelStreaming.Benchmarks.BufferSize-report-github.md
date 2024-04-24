```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method      | BufferSizeKb | Mean    | Error    | StdDev   | Allocated |
|------------ |------------- |--------:|---------:|---------:|----------:|
| ExcelWriter | 100          | 1.659 s | 0.0326 s | 0.0467 s |  11.09 KB |
| ExcelWriter | 256          | 1.668 s | 0.0328 s | 0.0415 s |  11.09 KB |
| ExcelWriter | 64           | 1.680 s | 0.0332 s | 0.0326 s |  11.09 KB |
| ExcelWriter | 16           | 1.688 s | 0.0314 s | 0.0408 s |  11.23 KB |
| ExcelWriter | 500          | 1.701 s | 0.0304 s | 0.0474 s |  11.09 KB |
| ExcelWriter | 8            | 1.706 s | 0.0332 s | 0.0466 s |  11.23 KB |
| ExcelWriter | 4            | 1.765 s | 0.0352 s | 0.0493 s |  11.23 KB |
