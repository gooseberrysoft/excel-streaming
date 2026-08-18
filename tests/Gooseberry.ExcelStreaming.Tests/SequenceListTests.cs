using System.Runtime.InteropServices;
using FluentAssertions;
using Xunit;

namespace Gooseberry.ExcelStreaming.Tests
{
    public sealed class SequenceListTests
    {
        [Fact]
        public void Next_FirstCall_ReturnsZeroAndLengthIsInitialBucket()
        {
            using var list = new SequenceList<int>();

            var index = list.Next();

            index.Should().Be(0);
            list.Length.Should().Be(256);
        }

        [Fact]
        public void Next_Growth_HappensWhenBoundaryReached()
        {
            using var list = new SequenceList<int>();

            // first call grows to first bucket (256)
            list.Next();
            list.Length.Should().Be(256);

            // call until last index = 255 (total 256 items) -> no further growth
            for (int i = 1; i < 256; i++)
                list.Next();

            list.Length.Should().Be(256);

            // next call should force next extension (256 + 512 = 768)
            var idx = list.Next();
            idx.Should().Be(256);
            list.Length.Should().Be(768);
        }

        [Fact]
        public void SetAndGet_ValuesAreStoredAndRetrieved()
        {
            using var list = new SequenceList<int>();

            list.Next(); // index 0
            list.Next(); // index 1
            list.Next(); // index 2

            list[0] = 10;
            list[1] = 20;
            list[2] = 30;

            list[0].Should().Be(10);
            list[1].Should().Be(20);
            list[2].Should().Be(30);
        }

        [Fact]
        public void Set_Get_IndexGreaterThanLast_ThrowsArgumentException()
        {
            using var list = new SequenceList<int>();

            list.Next(); // lastIndex == 0

            Action setOutOfRange = () => list[1] = 123;
            Action getOutOfRange = () => { var _ = list[1]; };

            setOutOfRange.Should().Throw<ArgumentException>().WithMessage("Index must be less than*");
            getOutOfRange.Should().Throw<ArgumentException>().WithMessage("Index must be less than*");
        }

        [Fact]
        public void Dispose_ResetsLength_AndAllowsReuse()
        {
            var list = new SequenceList<int>();

            list.Next();
            list.Length.Should().BeGreaterThan(0);

            list.Dispose();
            list.Length.Should().Be(0);

            // after dispose we can reuse (first Next grows again)
            var newIndex = list.Next();
            newIndex.Should().Be(0);
            list.Length.Should().Be(256);

            list.Dispose();
        }

        [Fact]
        public void SetAndGet_AllIndexes()
        {
            const int count = 2_001_024;

            using var list = new SequenceList<int>();

            // allocate and set values
            for (var i = 0; i < count; i++)
            {
                var idx = list.Next();
                idx.Should().Be(i);
                list[idx] = i;
            }

            // verify values for every index
            for (var i = 0; i < count; i++)
            {
                list[i].Should().Be(i);
            }
        }

        [Fact]
        public void GetEnumerator_EmptyList_ReturnsNoElements()
        {
            using var list = new SequenceList<int>();

            var buckets = list.ToList();
            buckets.Should().BeEmpty();
        }

        [Fact]
        public void GetEnumerator_FullFirstBucket_ReturnsSingleFullBucket()
        {
            const int count = 256; // exactly one full bucket
            using var list = new SequenceList<int>();

            for (var i = 0; i < count; i++)
            {
                var idx = list.Next();
                list[idx] = i;
            }

            var buckets = list.ToArray();
            buckets.Should().HaveCount(1);
            buckets[0].Length.Should().Be(256);

            var values = buckets.SelectMany(m => m.ToArray()).ToArray();
            values.Should().HaveCount(count);
            values.Should().Equal(Enumerable.Range(0, count));
        }

        [Fact]
        public void GetEnumerator_PartialLastBucket_ReturnsOnlyPopulatedElements()
        {
            const int count = 300; // 256 + 44 -> two buckets, second partially filled
            using var list = new SequenceList<int>();

            for (var i = 0; i < count; i++)
            {
                var idx = list.Next();
                list[idx] = i;
            }

            var buckets = list.ToArray();
            buckets.Should().HaveCount(2);
            buckets[0].Length.Should().Be(256);
            buckets[1].Length.Should().Be(44);

            var values = buckets.SelectMany(MemoryMarshal.ToEnumerable).ToArray();
            values.Should().HaveCount(count);
            values.Should().Equal(Enumerable.Range(0, count));
        }

        [Fact]
        public void GetEnumerator_PartialLastBucket_ReturnsOnlyPopulatedElements_2()
        {
            const int count = 1_000_011; 
            using var list = new SequenceList<int>();

            for (var i = 0; i < count; i++)
            {
                var idx = list.Next();
                list[idx] = i;
            }
            
            var values = list.SelectMany(MemoryMarshal.ToEnumerable).ToArray();
            values.Should().HaveCount(count);
            values.Should().Equal(Enumerable.Range(0, count));
        }

        [Fact]
        public void GetRefValue_Throws_When_Index_Greater_Than_LastIndex()
        {
            var seq = new SequenceList<int>();

            Assert.Throws<ArgumentException>(() =>
            {
                // calling in a by-value context will still invoke the method and throw
                var _ = seq.GetRefValue(0);
            });
        }

        [Fact]
        public void GetRefValue_Returns_Ref_That_Can_Be_Changed_ByRef()
        {
            var seq = new SequenceList<int>();

            seq.Next(); // allocate first slot
            // obtain a by-ref to the element and modify it
            ref var r = ref seq.GetRefValue(0);
            r = 555;

            // value must be observable through the indexer
            seq[0].Should().Be(555);

            // modify again via the ref and ensure change is reflected
            r++;
            seq[0].Should().Be(556);
        }

    }
}