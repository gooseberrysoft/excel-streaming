```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method   | RowsCount | Mean         | Error     | StdDev    | Allocated |
|--------- |---------- |-------------:|----------:|----------:|----------:|
| AsyncZip | 100       |     1.706 ms | 0.0154 ms | 0.0128 ms |  13.25 KB |
| SyncZip  | 100       |     1.847 ms | 0.0368 ms | 0.0344 ms |  10.54 KB |
| AsyncZip | 1000      |    12.516 ms | 0.0775 ms | 0.0647 ms |  13.39 KB |
| SyncZip  | 1000      |    16.595 ms | 0.3158 ms | 0.3379 ms |  10.69 KB |
| AsyncZip | 10000     |   109.857 ms | 1.1263 ms | 0.9984 ms |  15.44 KB |
| SyncZip  | 10000     |   162.789 ms | 1.0781 ms | 0.9557 ms |     12 KB |
| AsyncZip | 100000    | 1,106.287 ms | 8.8515 ms | 6.9107 ms |  86.54 KB |
| SyncZip  | 100000    | 1,601.511 ms | 6.6775 ms | 5.9194 ms |  23.29 KB |
