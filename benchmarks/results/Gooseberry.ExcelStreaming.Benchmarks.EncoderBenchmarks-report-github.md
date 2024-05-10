```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method      | Mean     | Error   | StdDev   | Median   | Allocated |
|------------ |---------:|--------:|---------:|---------:|----------:|
| Utf8Bytes   | 248.5 μs | 4.88 μs | 10.07 μs | 247.8 μs |         - |
| UnsafeChars | 307.3 μs | 9.07 μs | 24.83 μs | 299.1 μs |         - |
