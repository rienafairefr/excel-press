#!/usr/bin/env python
import argparse
import sys
import math
import cStringIO
import olefile

parser = argparse.ArgumentParser(description='Compress VBA macro file')
parser.add_argument('-f', '--file', help='Decompressed VBA macro file', required=True)
args = parser.parse_args()

def main():
    eight_char = []
    attribute_header = '\x00\x41\x74\x74\x72\x69\x62\x75\x74'
    macro_name = 'VB_name = "Module1"\r\n'

    data = open(args.file, 'rb').read()
    full_data = "{macro_name} {data}".format(macro_name=macro_name, data=data)

    for char in full_data:
        hex_char = char.encode('hex')
        eight_char.append(hex_char)

    chunk = eight_char[:8]
    print chunk



def compress():
    compressed_data = ''

    # Set signature byte '\x01'
    compressed_data += '\x01'
    compressed_current = data
    while data < 0:
        CompressedChunkStart = CompressedCurrent
        DecompressedChunkStart = DecompressedCurrent
        compressChunk()

def compressChunk(decompressedChunk):
    CompressedEnd = CompressedChunkStart + 4098
    CompressedCurrent = CompressedChunkStart + 2
    DecompressedEnd TO the minimum of (DecompressedChunkStart + 4096) & DecompressedBufferEnd
    while (DecompressedCurrent < DecompressedEnd) and (CompressedCurrent < CompressedEnd):

        compressTokenSequence(CompressedEnd, DecompressedEnd)

    if DecompressedCurrent < DecompressedEnd:
        CompressRawChunk(DecompressedEnd - 1)
        CompressedFlag = 0
    else:
        CompressedFlag = 1

    Size = CompressedCurrent - CompressedChunkStart
    Header = 0x0000
    PackCompressedChunkSize(Size, Header)
    PackCompressedChunkFlag(CompressedFlag, Header)
    PackCompressedChunkSignature(Header)
    CompressedChunkHeader = Header

def PackCompressedChunkSize(Size, Header):
    temp1 = Header BITWISE & 0xF000
    temp2 = Size - 3
    Header = temp1 BITWISE OR temp2

def PackCompressedChunkFlag(CompressedFlag, Header):
    temp1 = Header BITWISE & 0x7FFF
    temp2 = CompressedFlag LEFT SHIFT BY 15
    Header = temp1 BITWISE OR temp2

def PackCompressedChunkSignature(Header):
    temp = Header BITWISE & 0x8FFF
    Header = temp BITWISE OR 0x3000

def compressTokenSequence(CompressedEnd, DecompressedEnd):
    FlagByteIndex = CompressedCurrent
    TokenFlags = 0b00000000
    INCREMENT CompressedCurrent
    for index FROM 0 TO 7 INCLUSIVE:
        if (DecompressedCurrent < DecompressedEnd) and (CompressedCurrent < CompressedEnd):
            TokenFlags = compressToken(CompressedEnd, DecompressedEnd, index, TokenFlags)

    FlagByteIndex = TokenFlags

def compressToken(CompressedEnd, DecompressedEnd, index, TokenFlags):
    Offset = 0
    Offset, Length = Matching(DecompressedEnd)
    if Offset != 0:
        if (CompressedCurrent + 1) < CompressedEnd:
            Token = PackCopyToken(Offset, Length)
            APPEND the bytes of the CopyToken (section 2.4.1.1.8) Token TO CompressedCurrent in little-endian order
            SetFlagBit(index, 1, Flags)
            INCREMENT CompressedCurrent BY 2
            INCREMENT DecompressedCurrent BY Length
        else:
            CompressedCurrent = CompressedEnd
    else:
        if CompressedCurrent < CompressedEnd:
            APPEND the byte of the LiteralToken at DecompressedCurrent TO CompressedCurrent
            INCREMENT CompressedCurrent
            INCREMENT DecompressedCurrent
        else:
            CompressedCurrent = CompressedEnd

def CompressRawChunk():
    CompressedCurrent = CompressedChunkStart + 2
    DecompressedCurrent = DecompressedChunkStart
    PadCount = 4096
    for each byte, B, FROM DecompressedChunkStart TO LastByte INCLUSIVE:
        COPY B TO CompressedCurrent
        INCREMENT CompressedCurrent
        INCREMENT DecompressedCurrent
        DECREMENT PadCount
    for counter FROM 1 TO PadCount INCLUSIVE:
        COPY 0x00 TO CompressedCurrent
        INCREMENT CompressedCurrent

def Matching(DecompressedEnd):
    Candidate = DecompressedCurrent - 1
    BestLength = 0
    while Candidate >= DecompressedChunkStart:
        C = Candidate
        D = DecompressedCurrent
        Len = 0
        while (D < DecompressedEnd) and (D = C):
            INCREMENT Len
            INCREMENT C
            INCREMENT D

        if Len > BestLength:
            BestLength = Len
            BestCandidate = Candidate

        DECREMENT Candidate
    if BestLength >= 3:
        MaximumLength = CopyTokenHelp()
        Length = the MINIMUM of BestLength and MaximumLength
        Offset = DecompressedCurrent - BestCandidate
    else:
        Length = 0
        Offset = 0

def PackCopyToken(Offset, Length):
    LengthMask, OffsetMask, BitCount = CopyTokenHelp()
    temp1 = Offset - 1
    temp2 = 16 - BitCount
    temp3 = Length - 3
    Token = (temp1 LEFT SHIFT BY temp2) BITWISE OR temp3

def SetFlagBit(index, 1, Flags):
    temp1 = Flag LEFT SHIFT BY Index
    temp2 = Byte BITWISE AND (BITWISE NOT temp1)
    Byte = temp2 BITWISE OR temp1

if __name__ == '__main__':
    main()
