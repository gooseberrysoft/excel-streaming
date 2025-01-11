```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]     : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method          | RowsCount | Mean         | Error      | StdDev     | Gen0   | Allocated |
|---------------- |---------- |-------------:|-----------:|-----------:|-------:|----------:|
| RealWorldReport | 100       |     1.046 ms |  0.0199 ms |  0.0195 ms | 1.9531 |  14.01 KB |
| RealWorldReport | 1000      |     7.164 ms |  0.0400 ms |  0.0374 ms |      - |  14.17 KB |
| RealWorldReport | 10000     |    62.502 ms |  0.8939 ms |  0.8362 ms |      - |  15.56 KB |
| RealWorldReport | 100000    |   638.745 ms | 12.1795 ms | 15.4031 ms |      - |  27.64 KB |
| RealWorldReport | 500000    | 3,247.661 ms | 64.1582 ms | 83.4238 ms |      - |  81.38 KB |
