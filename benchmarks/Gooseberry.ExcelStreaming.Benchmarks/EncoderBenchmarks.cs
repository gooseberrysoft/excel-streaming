using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using BenchmarkDotNet.Order;
using DocumentFormat.OpenXml.Bibliography;
using System.Text.Encodings.Web;

namespace Gooseberry.ExcelStreaming.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
public class EncoderBenchmarks
{
    [Benchmark]
    public int UnsafeChars()
    {
        var indexToEncode = 0;
        for (var i = 0; i < 10_000; i++)
        {
            unsafe
            {
                var s1 = "Tags such as <img> and <input> directly introduce content into the page.";
                fixed (char* pText = s1)
                {
                    indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncode(pText, s1.Length);
                }

                var s2 = "The cat (Felis catus), commonly referred to as the domestic cat";
                fixed (char* pText = s2)
                {
                    indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncode(pText, s2.Length);
                }

                var s3 = "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf";
                fixed (char* pText = s3)
                {
                    indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncode(pText, s3.Length);
                }
            }
        }

        return indexToEncode;
    }

    [Benchmark]
    public int Utf8Bytes()
    {
        var indexToEncode = 0;

        for (var i = 0; i < 10_000; i++)
        {
            var s1 = "Tags such as <img> and <input> directly introduce content into the page."u8;
            indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncodeUtf8(s1);

            var s2 = "The cat (Felis catus), commonly referred to as the domestic cat"u8;
            indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncodeUtf8(s2);

            var s3 = "The dog (Canis familiaris or Canis lupus familiaris) is a domesticated descendant of the wolf"u8;
            indexToEncode = HtmlEncoder.Default.FindFirstCharacterToEncodeUtf8(s3);
        }

        return indexToEncode;
    }
}