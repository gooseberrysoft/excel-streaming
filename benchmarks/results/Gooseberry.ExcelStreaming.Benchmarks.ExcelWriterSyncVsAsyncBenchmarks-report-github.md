```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method   | RowsCount | Mean         | Error      | StdDev     | Allocated |
|--------- |---------- |-------------:|-----------:|-----------:|----------:|
| AsyncZip | 100       |     1.630 ms |  0.0315 ms |  0.0323 ms |  15.33 KB |
| SyncZip  | 100       |     1.758 ms |  0.0231 ms |  0.0216 ms |  10.49 KB |
| AsyncZip | 1000      |    11.563 ms |  0.0353 ms |  0.0276 ms |  15.35 KB |
| SyncZip  | 1000      |    15.862 ms |  0.1990 ms |  0.1662 ms |   10.7 KB |
| AsyncZip | 10000     |   107.735 ms |  1.5390 ms |  1.4396 ms |  39.55 KB |
| SyncZip  | 10000     |   156.068 ms |  1.2315 ms |  1.0284 ms |  12.21 KB |
| AsyncZip | 100000    | 1,045.828 ms |  7.4986 ms |  7.0142 ms | 283.48 KB |
| SyncZip  | 100000    | 1,554.097 ms | 14.5628 ms | 12.9096 ms |  26.86 KB |
