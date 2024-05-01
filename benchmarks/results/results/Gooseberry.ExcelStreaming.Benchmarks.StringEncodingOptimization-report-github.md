```

BenchmarkDotNet v0.13.12, Windows 10 (10.0.19045.4291/22H2/2022Update)
11th Gen Intel Core i7-1185G7 3.00GHz, 1 CPU, 8 logical and 4 physical cores
.NET SDK 8.0.204
  [Host]     : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI
  DefaultJob : .NET 8.0.4 (8.0.424.16909), X64 RyuJIT AVX-512F+CD+BW+DQ+VL+VBMI


```
| Method                        | IterationCount | Mean           | Error         | StdDev        | Median         | Allocated |
|------------------------------ |--------------- |---------------:|--------------:|--------------:|---------------:|----------:|
| OptimizedImpl_Regular         | 1              |       287.5 ns |       4.39 ns |       3.67 ns |       288.8 ns |         - |
| DefaultImpl_Regular           | 1              |       466.2 ns |       8.95 ns |      12.84 ns |       467.2 ns |         - |
| OptimizedSimpleImpl_Escapable | 1              |       782.2 ns |      11.98 ns |      10.00 ns |       784.1 ns |         - |
| DefaultImpl_Escapable         | 1              |       791.3 ns |      33.62 ns |      97.55 ns |       741.0 ns |         - |
| OptimizedImpl_Escapable       | 1              |       962.7 ns |      19.10 ns |      31.39 ns |       959.9 ns |         - |
| OptimizedImpl_Regular         | 10             |     2,775.5 ns |      51.05 ns |      50.14 ns |     2,791.4 ns |         - |
| DefaultImpl_Regular           | 10             |     4,594.7 ns |      86.80 ns |      92.87 ns |     4,600.6 ns |         - |
| DefaultImpl_Escapable         | 10             |     7,223.7 ns |     131.48 ns |     109.80 ns |     7,215.8 ns |         - |
| OptimizedSimpleImpl_Escapable | 10             |     7,929.5 ns |     155.62 ns |     159.81 ns |     7,933.0 ns |         - |
| OptimizedImpl_Escapable       | 10             |    10,000.2 ns |     197.83 ns |     446.53 ns |     9,796.1 ns |         - |
| OptimizedImpl_Regular         | 100            |    27,760.9 ns |     534.14 ns |     675.52 ns |    27,797.1 ns |         - |
| DefaultImpl_Regular           | 100            |    46,311.3 ns |     867.08 ns |   1,400.17 ns |    46,025.7 ns |         - |
| DefaultImpl_Escapable         | 100            |    69,525.4 ns |   1,357.07 ns |   1,133.22 ns |    69,861.1 ns |         - |
| OptimizedSimpleImpl_Escapable | 100            |    75,654.1 ns |   1,344.39 ns |   1,438.49 ns |    76,078.6 ns |         - |
| OptimizedImpl_Escapable       | 100            |    97,709.2 ns |   1,340.82 ns |   1,119.64 ns |    97,940.7 ns |         - |
| OptimizedImpl_Regular         | 1000           |   276,892.2 ns |   1,673.29 ns |   1,483.33 ns |   276,419.1 ns |         - |
| DefaultImpl_Regular           | 1000           |   460,337.8 ns |   8,393.01 ns |   9,991.28 ns |   458,627.2 ns |         - |
| DefaultImpl_Escapable         | 1000           |   695,982.7 ns |  11,815.62 ns |   9,224.86 ns |   697,523.4 ns |         - |
| OptimizedSimpleImpl_Escapable | 1000           |   750,660.5 ns |  14,294.60 ns |  13,371.18 ns |   752,597.2 ns |         - |
| OptimizedImpl_Escapable       | 1000           |   958,150.1 ns |  10,922.82 ns |   9,121.05 ns |   954,130.9 ns |       1 B |
| OptimizedImpl_Regular         | 10000          | 2,798,804.1 ns |  55,431.06 ns |  61,611.48 ns | 2,790,460.9 ns |       2 B |
| DefaultImpl_Regular           | 10000          | 4,569,119.1 ns |  24,025.76 ns |  20,062.60 ns | 4,571,769.5 ns |       3 B |
| DefaultImpl_Escapable         | 10000          | 7,241,752.2 ns | 108,133.20 ns | 101,147.86 ns | 7,249,484.0 ns |       3 B |
| OptimizedSimpleImpl_Escapable | 10000          | 7,893,833.5 ns | 156,336.49 ns | 203,281.60 ns | 7,909,189.1 ns |       3 B |
| OptimizedImpl_Escapable       | 10000          | 9,987,148.9 ns | 187,738.60 ns | 175,610.80 ns | 9,899,782.8 ns |       6 B |
