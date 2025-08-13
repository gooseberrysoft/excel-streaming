```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.6093/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.302
  [Host]     : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 9.0.7 (9.0.725.31616), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method              | Mean       | Error     | StdDev     | Median     |
|-------------------- |-----------:|----------:|-----------:|-----------:|
| DateCustomFormatter |   3.574 ms | 0.1830 ms |  0.5396 ms |   3.413 ms |
| CustomFormatter     |  11.247 ms | 0.3841 ms |  1.1143 ms |  11.017 ms |
| DateDoubleFormatter |  83.862 ms | 1.6674 ms |  1.7123 ms |  83.485 ms |
| DoubleFormatter     | 144.717 ms | 5.4055 ms | 15.6822 ms | 142.020 ms |
