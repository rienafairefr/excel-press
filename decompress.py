#!/usr/bin/env python
import argparse
import sys
import math
import cStringIO
import olefile

parser = argparse.ArgumentParser(description='Decompress VBA macro from unzipped .xls document')
parser.add_argument('-f', '--file', help='VBA macro file', required=True)
parser.add_argument('--readable', help='Output compressed VBA file', action='store_true', required=False)
parser.add_argument('--raw', help='Output compressed VBA file', action='store_true', required=False)
args = parser.parse_args()

def main():
    macro_file = open(args.file, 'rb').read()
    if args.raw:
        print search(macro_file)
    else:
        stdoutWriteChunked(decompress(macro_file))

def search(data):
    # Find the compressed data
    searchString = '\x00Attribut'
    position = data.find(searchString)
    if position != -1 and data[position + len(searchString)] == 'e':
        position = -1    

    compressedData = data[position - 3:]

    return compressedData

def decompress(data):
    # Find the compressed data
    searchString = '\x00Attribut'
    position = data.find(searchString)
    if position != -1 and data[position + len(searchString)] == 'e':
        position = -1    

    compressedData = data[position - 3:]

    # Actually decompress the data
    remainder = compressedData[1:]
    decompressed = ''
    while len(remainder) != 0:
        decompressedChunk, remainder = decompressChunk(remainder)
        decompressed += decompressedChunk

    return decompressed

def decompressChunk(compressedChunk):
    if len(compressedChunk) < 2:
        return None, None
    # 0x4f and 0xb4
    header = ord(compressedChunk[0]) + ord(compressedChunk[1]) * 0x100
    size = (header & 0x0FFF) + 3
    flagCompressed = header & 0x8000
    data = compressedChunk[2:2 + size - 2]

    decompressedChunk = ''
    while len(data) != 0:
        tokens, data = parseTokenSequence(data)
        for token in tokens:
            if len(token) == 1:
                decompressedChunk += token
            else:
                numberOfOffsetBits = offsetBits(decompressedChunk)
                copyToken = ord(token[0]) + ord(token[1]) * 0x100
                offset = 1 + (copyToken >> (16 - numberOfOffsetBits))
                length = 3 + (((copyToken << numberOfOffsetBits) & 0xFFFF) >> numberOfOffsetBits)
                copy = decompressedChunk[-offset:]
                copy = copy[0:length]
                lengthCopy = len(copy)
                while length > lengthCopy: #a#
                    if length - lengthCopy >= lengthCopy:
                        copy += copy[0:lengthCopy]
                        length -= lengthCopy
                    else:
                        copy += copy[0:length - lengthCopy]
                        length -= length - lengthCopy
                decompressedChunk += copy
    return decompressedChunk, compressedChunk[size:]

def offsetBits(data):
    numberOfBits = int(math.ceil(math.log(len(data), 2)))
    if numberOfBits < 4:
        numberOfBits = 4
    elif numberOfBits > 12:
        numberOfBits = 12
    return numberOfBits

def stdoutWriteChunked(data):
    while data != '':
        sys.stdout.write(data[0:10000])
        try:
            sys.stdout.flush()
        except IOError:
            return
        data = data[10000:]

def parseTokenSequence(data):
    flags = ord(data[0])
    data = data[1:]
    result = []
    for mask in [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80]:
        if len(data) > 0:
            if flags & mask:
                result.append(data[0:2])
                data = data[2:]
            else:
                result.append(data[0])
                data = data[1:]
    return result, data

if __name__ == '__main__':
    main()
