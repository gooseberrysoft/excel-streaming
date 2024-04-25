using System.Text;
using System.Text.Encodings.Web;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using Gooseberry.ExcelStreaming.Writers;

namespace Gooseberry.ExcelStreaming.Benchmarks
{
    [MemoryDiagnoser]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    public class StringEncodingOptimization
    {
        private readonly BuffersChain _buffer = new BuffersChain(bufferSize: 4 * 1024, flushThreshold: 1.0);
        private readonly Encoder _encoder = Encoding.UTF8.GetEncoder();
        private static readonly Stream NullStream = new NullStream();

        private static string[] _regularStrings = new[]
        {
            "The wild winds weep",
            "And the night is a-cold;",
            "Come hither, Sleep,",
            "And my griefs infold:",
            "But lo! the morning peeps",
            "Over the eastern steeps,",
            "And the rustling birds of dawn",
            "The earth do scorn.",
            "They make mad the roaring winds,",
            "And with tempests play."
        };

        private static string[] _escapableStrings = new[]
        {
            "The wild > winds weep",
            "And the night < is a-cold;",
            "Come hither,& Sleep,",
            "And my griefs infold:&",
            "But lo! the \" morning peeps>",
            "Over <the eastern steeps,",
            "<And the rustling birds of dawn",
            "The > earth& do > \"scorn.",
            "They make <>\" mad \"the roaring winds,",
            "And with &tempests play."
        };


        [Params(1, 10, 100, 1000, 10_000)]
        public int IterationCount { get; set; }

        [Benchmark]
        public async Task DefaultImpl_Regular()
        {
            for (int i = 0; i < IterationCount; i++)
            {
                foreach (var str in _regularStrings)
                    str.AsSpan().WriteEscapedTo(_buffer, _encoder);

                await _buffer.FlushAll(NullStream, CancellationToken.None);
            }
        }

        [Benchmark]
        public async Task OptimizedImpl_Regular()
        {
            for (int i = 0; i < IterationCount; i++)
            {
                foreach (var str in _regularStrings)
                    OptimizedWrite(str);

                await _buffer.FlushAll(NullStream, CancellationToken.None);
            }
        }

        [Benchmark]
        public async Task DefaultImpl_Escapable()
        {
            for (int i = 0; i < IterationCount; i++)
            {
                foreach (var str in _escapableStrings)
                    str.AsSpan().WriteEscapedTo(_buffer, _encoder);

                await _buffer.FlushAll(NullStream, CancellationToken.None);
            }
        }

        [Benchmark]
        public async Task OptimizedSimpleImpl_Escapable()
        {
            for (int i = 0; i < IterationCount; i++)
            {
                foreach (var str in _escapableStrings)
                    OptimizedWrite2(str);

                await _buffer.FlushAll(NullStream, CancellationToken.None);
            }
        }

        [Benchmark]
        public async Task OptimizedImpl_Escapable()
        {
            for (int i = 0; i < IterationCount; i++)
            {
                foreach (var str in _escapableStrings)
                    OptimizedWrite(str);

                await _buffer.FlushAll(NullStream, CancellationToken.None);
            }
        }

        private unsafe void OptimizedWrite(string str)
        {
            int index = 0;
            fixed (char* pText = str)
            {
                index = HtmlEncoder.Default.FindFirstCharacterToEncode(pText, str.Length);
            }

            var chars = str.AsSpan();

            if (index == -1)
            {
                chars.WriteTo(_buffer, _encoder);
            }
            else if (index == 0)
            {
                chars.WriteEscapedTo(_buffer, _encoder);
            }
            else
            {
                var span = _buffer.GetSpan();
                var written = 0;

                chars[..index].WriteTo(_buffer, _encoder, ref span, ref written);
                chars[index..].WriteEscapedTo(_buffer, _encoder, ref span, ref written);

                _buffer.Advance(written);
            }
        }

        private unsafe void OptimizedWrite2(string str)
        {
            int index = 0;
            fixed (char* pText = str)
            {
                index = HtmlEncoder.Default.FindFirstCharacterToEncode(pText, str.Length);
            }

            var chars = str.AsSpan();

            if (index == -1)
            {
                chars.WriteTo(_buffer, _encoder);
            }
            else
            {
                chars.WriteEscapedTo(_buffer, _encoder);
            }
        }
    }
}