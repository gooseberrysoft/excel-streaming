```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]   : .NET 9.0.0 (9.0.24.52809), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 8.0 : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 9.0 : .NET 9.0.0 (9.0.24.52809), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                | Job      | Runtime  | RowsCount | Mean      | Error    | StdDev   |
|---------------------- |--------- |--------- |---------- |----------:|---------:|---------:|
| SimpleStringsWriting  | .NET 8.0 | .NET 8.0 | 100000    | 100.77 ms | 1.982 ms | 2.506 ms |
| SimpleStringsWriting  | .NET 9.0 | .NET 9.0 | 100000    |  56.64 ms | 0.661 ms | 0.618 ms |
| EscapedStringsWriting | .NET 8.0 | .NET 8.0 | 100000    | 144.89 ms | 0.927 ms | 0.723 ms |
| EscapedStringsWriting | .NET 9.0 | .NET 9.0 | 100000    |  80.38 ms | 1.097 ms | 0.973 ms |
| DateTimeWriting       | .NET 8.0 | .NET 8.0 | 100000    |  25.28 ms | 0.213 ms | 0.189 ms |
| DateTimeWriting       | .NET 9.0 | .NET 9.0 | 100000    |  27.08 ms | 0.540 ms | 0.887 ms |
| IntWriting            | .NET 8.0 | .NET 8.0 | 100000    | 128.58 ms | 1.342 ms | 1.255 ms |
| IntWriting            | .NET 9.0 | .NET 9.0 | 100000    |  87.84 ms | 1.260 ms | 1.117 ms |
| DecimalWriting        | .NET 8.0 | .NET 8.0 | 100000    |  33.93 ms | 0.288 ms | 0.270 ms |
| DecimalWriting        | .NET 9.0 | .NET 9.0 | 100000    |  31.38 ms | 0.438 ms | 0.410 ms |
