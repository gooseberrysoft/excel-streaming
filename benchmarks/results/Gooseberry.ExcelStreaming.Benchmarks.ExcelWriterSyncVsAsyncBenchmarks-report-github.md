```

BenchmarkDotNet v0.15.2, Windows 10 (10.0.19045.6093/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz (Max: 1.80GHz), 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.302
  [Host]     : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method   | RowsCount | Mean         | Error      | StdDev     | Median       | Gen0   | Allocated |
|--------- |---------- |-------------:|-----------:|-----------:|-------------:|-------:|----------:|
| AsyncZip | 100       |     1.391 ms |  0.0278 ms |  0.0599 ms |     1.408 ms | 1.9531 |  16.64 KB |
| SyncZip  | 100       |     1.612 ms |  0.0319 ms |  0.0615 ms |     1.616 ms | 1.9531 |     14 KB |
| AsyncZip | 1000      |    10.722 ms |  0.2127 ms |  0.4440 ms |    10.539 ms |      - |  16.66 KB |
| SyncZip  | 1000      |    14.373 ms |  0.2860 ms |  0.2809 ms |    14.317 ms |      - |     14 KB |
| AsyncZip | 10000     |    97.231 ms |  1.2172 ms |  1.0164 ms |    96.888 ms |      - |  16.73 KB |
| SyncZip  | 10000     |   136.363 ms |  2.6709 ms |  3.2801 ms |   135.843 ms |      - |     14 KB |
| AsyncZip | 100000    |   911.601 ms | 16.1153 ms | 15.0743 ms |   906.647 ms |      - |  16.87 KB |
| SyncZip  | 100000    | 1,350.763 ms | 26.7790 ms | 27.5001 ms | 1,350.734 ms |      - |  14.07 KB |
