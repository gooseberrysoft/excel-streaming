```

BenchmarkDotNet v0.14.0, Windows 10 (10.0.19045.5247/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 9.0.101
  [Host]   : .NET 9.0.0 (9.0.24.52809), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 8.0 : .NET 8.0.11 (8.0.1124.51707), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  .NET 9.0 : .NET 9.0.0 (9.0.24.52809), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                | Job      | Runtime  | RowsCount | Mean      | Error    | StdDev    |
|---------------------- |--------- |--------- |---------- |----------:|---------:|----------:|
| SimpleStringsWriting  | .NET 8.0 | .NET 8.0 | 100000    | 121.03 ms | 2.555 ms |  7.453 ms |
| SimpleStringsWriting  | .NET 9.0 | .NET 9.0 | 100000    |  60.14 ms | 0.856 ms |  0.801 ms |
| EscapedStringsWriting | .NET 8.0 | .NET 8.0 | 100000    | 161.41 ms | 3.041 ms |  4.996 ms |
| EscapedStringsWriting | .NET 9.0 | .NET 9.0 | 100000    |  85.58 ms | 1.710 ms |  1.515 ms |
| DateTimeWriting       | .NET 8.0 | .NET 8.0 | 100000    | 210.90 ms | 4.169 ms |  6.491 ms |
| DateTimeWriting       | .NET 9.0 | .NET 9.0 | 100000    | 139.53 ms | 0.915 ms |  0.811 ms |
| IntWriting            | .NET 8.0 | .NET 8.0 | 100000    | 127.08 ms | 2.128 ms |  2.090 ms |
| IntWriting            | .NET 9.0 | .NET 9.0 | 100000    |  60.76 ms | 3.519 ms | 10.320 ms |
| DecimalWriting        | .NET 8.0 | .NET 8.0 | 100000    | 155.06 ms | 2.207 ms |  2.065 ms |
| DecimalWriting        | .NET 9.0 | .NET 9.0 | 100000    |  85.18 ms | 1.658 ms |  1.910 ms |
