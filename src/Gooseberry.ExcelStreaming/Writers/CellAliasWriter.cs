using System;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Gooseberry.ExcelStreaming.Configuration;

namespace Gooseberry.ExcelStreaming.Writers;

internal static class CellAliasWriter
{
    private static readonly byte[] EncodedRowAlphabet = Encoding.UTF8.GetBytes(CellReference.RowAlphabet);
    private static readonly byte[] EncodedColumnAlphabet = Encoding.UTF8.GetBytes(CellReference.ColumnAlphabet);
    private static readonly byte Colon = Encoding.UTF8.GetBytes(":").Single();        
    
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this Merge merge,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        var aliasSize = CalculateSize(merge.TopLeft) + CalculateSize(merge.RightBottom) + 1;

        if (destination.Length < aliasSize)
        {
            buffer.Advance(written);
            destination = buffer.GetSpan(minSize: aliasSize);
            written = 0;
        }

        var span = destination.Slice(0, aliasSize);

        WriteAlias(merge.RightBottom, ref span);
            
        span[^1] = Colon;
        span = span[..^1];
        
        WriteAlias(merge.TopLeft, ref span);

        written += aliasSize;
        destination = destination.Slice(aliasSize);
    }    
    
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static void WriteTo(
        this CellReference cellReference,
        BuffersChain buffer,
        ref Span<byte> destination,
        ref int written)
    {
        var aliasSize = CalculateSize(cellReference); 
            
        if (destination.Length < aliasSize)
        {
            buffer.Advance(written);
            destination = buffer.GetSpan(minSize: aliasSize);
            written = 0;
        }

        var span = destination.Slice(0, aliasSize);
            
        WriteAlias(cellReference, ref span);
            
        written += aliasSize;
        destination = destination.Slice(aliasSize);
    }

    private static int CalculateSize(CellReference cellReference)
    {
        return CalculateSize(cellReference.Row, EncodedRowAlphabet.Length) +
            CalculateSize(cellReference.Column, EncodedColumnAlphabet.Length);
    }
    
    private static int CalculateSize(uint value, int alphabetLength)
    {
        var size = 0;
        var remain = (int)value;
        do
        {
            remain /= alphabetLength;
            size++;
        } while (remain > 0);

        return size;
    }
    
    private static void WriteAlias(CellReference cellReference, ref Span<byte> destination)
    {
        Write(cellReference.Row, EncodedRowAlphabet, ref destination);
        // Column alias A is 1, so pass column - 1
        Write(cellReference.Column - 1, EncodedColumnAlphabet, ref destination);
    }
 
    private static void Write(uint value, byte[] alphabet, ref Span<byte> destination)
    {
        var destinationIndex = destination.Length - 1;
            
        var remain = (int)value;
        do
        {
            var charIndex = remain % alphabet.Length;

            if (destinationIndex < 0)
                throw new InvalidOperationException("Destination has no enough size to write alias.");

            destination[destinationIndex] = alphabet[charIndex];

            destinationIndex--;
            remain /= alphabet.Length;
        } while (remain > 0);

        destination = destination.Slice(0, destinationIndex + 1);
    }
}