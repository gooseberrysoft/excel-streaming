```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5965/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.302
  [Host]   : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 8.0 : .NET 8.0.18 (8.0.1825.31117), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 9.0 : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method          | Job      | Runtime  | RowsCount | Mean         | Error      | StdDev     | Median       | Allocated |
|---------------- |--------- |--------- |---------- |-------------:|-----------:|-----------:|-------------:|----------:|
| **RealWorldReport** | **.NET 8.0** | **.NET 8.0** | **100**       |     **1.171 ms** |  **0.0233 ms** |  **0.0249 ms** |     **1.166 ms** |  **13.37 KB** |
| RealWorldReport | .NET 9.0 | .NET 9.0 | 100       |     1.029 ms |  0.0202 ms |  0.0276 ms |     1.022 ms |  12.89 KB |
| **RealWorldReport** | **.NET 8.0** | **.NET 8.0** | **1000**      |     **7.890 ms** |  **0.1413 ms** |  **0.2115 ms** |     **7.841 ms** |   **13.9 KB** |
| RealWorldReport | .NET 9.0 | .NET 9.0 | 1000      |     7.514 ms |  0.1476 ms |  0.2843 ms |     7.405 ms |  13.43 KB |
| **RealWorldReport** | **.NET 8.0** | **.NET 8.0** | **10000**     |    **66.494 ms** |  **1.3297 ms** |  **3.2367 ms** |    **65.466 ms** |  **18.54 KB** |
| RealWorldReport | .NET 9.0 | .NET 9.0 | 10000     |    64.428 ms |  1.2388 ms |  3.5142 ms |    63.100 ms |  21.77 KB |
| **RealWorldReport** | **.NET 8.0** | **.NET 8.0** | **100000**    |   **648.958 ms** | **12.9503 ms** | **30.7779 ms** |   **636.136 ms** |  **65.39 KB** |
| RealWorldReport | .NET 9.0 | .NET 9.0 | 100000    |   635.639 ms | 12.6905 ms | 25.6354 ms |   636.715 ms | 106.74 KB |
| **RealWorldReport** | **.NET 8.0** | **.NET 8.0** | **500000**    | **3,228.361 ms** | **63.9838 ms** | **71.1178 ms** | **3,244.601 ms** | **274.79 KB** |
| RealWorldReport | .NET 9.0 | .NET 9.0 | 500000    | 3,187.095 ms | 61.0951 ms | 72.7293 ms | 3,189.199 ms | 490.32 KB |
