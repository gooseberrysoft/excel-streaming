```

BenchmarkDotNet v0.15.2, Windows 10 (10.0.19045.6093/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz (Max: 1.80GHz), 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.302
  [Host]   : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 9.0 : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI

Job=.NET 9.0  Runtime=.NET 9.0  

```
| Method        | RowsCount | Mean         | Error      | StdDev     | Gen0       | Gen1      | Gen2      | Allocated     |
|-------------- |---------- |-------------:|-----------:|-----------:|-----------:|----------:|----------:|--------------:|
| ExcelWriter   | 100       |     1.260 ms |  0.0369 ms |  0.1035 ms |          - |         - |         - |      14.23 KB |
| SpreadCheetah | 100       |     1.305 ms |  0.0258 ms |  0.0540 ms |          - |         - |         - |       8.61 KB |
| OpenXml       | 100       |     4.258 ms |  0.0825 ms |  0.1332 ms |   195.3125 |  195.3125 |  195.3125 |    1216.48 KB |
| ExcelWriter   | 1000      |     7.435 ms |  0.1484 ms |  0.2519 ms |          - |         - |         - |      14.32 KB |
| SpreadCheetah | 1000      |    11.655 ms |  0.2330 ms |  0.4377 ms |          - |         - |         - |       8.61 KB |
| OpenXml       | 1000      |    36.718 ms |  0.5714 ms |  0.5065 ms |  1214.2857 |  571.4286 |  500.0000 |   16630.37 KB |
| ExcelWriter   | 10000     |    73.579 ms |  1.3822 ms |  1.5363 ms |          - |         - |         - |      14.73 KB |
| SpreadCheetah | 10000     |   112.971 ms |  1.8776 ms |  1.5679 ms |          - |         - |         - |       9.11 KB |
| OpenXml       | 10000     |   349.437 ms |  6.1220 ms |  5.1121 ms | 10000.0000 | 3000.0000 | 3000.0000 |  142199.43 KB |
| ExcelWriter   | 100000    |   751.082 ms | 21.8239 ms | 64.3484 ms |          - |         - |         - |      29.38 KB |
| SpreadCheetah | 100000    | 1,210.076 ms | 25.2730 ms | 72.5129 ms |          - |         - |         - |       27.7 KB |
| OpenXml       | 100000    | 3,570.153 ms | 41.4747 ms | 34.6333 ms | 74000.0000 | 3000.0000 | 3000.0000 | 1225984.89 KB |
