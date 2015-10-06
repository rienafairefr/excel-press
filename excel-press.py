#!/usr/bin/env python
import argparse
import cStringIO
import math
#import olefile
import struct
import sys

parser = argparse.ArgumentParser(description='Compress or decompress a VBA macro file')
parser.add_argument('-c', '--compress', help='Compress the specified VBA macro file', required=False)
parser.add_argument('-d', '--decompress', help='Decompressed the specified VBA macro file', required=False)
parser.add_argument('--raw', help='Output the VBA data as compressed (HEX)', action='store_true', required=False)
args = parser.parse_args()

def main():
	if args.compress:
		offset = 0
		data = open(args.compress, 'rb').read()

		decompressed = UnCompressedVBAStream(data, offset)
		compressed = decompressed.Compress()
		print compressed
	else:
		macro_file = open(args.decompress, 'rb').read()
		print decompress(macro_file)

# Decompression
def decompress(data):
	# Find the compressed data
	searchString = '\x00Attribut'
	position = data.find(searchString)
	if position != -1 and data[position + len(searchString)] == 'e':
		position = -1    

	compressedData = data[position - 3:]

	if args.raw:
		return compressedData

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

	header = ord(compressedChunk[0]) + ord(compressedChunk[1]) * 0x100
	size = (header & 0x0FFF) + 3
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
				while length > lengthCopy:
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

# Compression
class VBAStreamBase:
	CHUNKSIZE = 4096
	def __init__(self, data, offset):
		self.mnOffset = offset
		self.data = data

	def CopyTokenHelp(self):
		difference = self.DecompressedCurrent - self.DecompressedChunkStart
		bitCount = 0
		while ((1 << bitCount) < difference):
			bitCount +=1

		if bitCount < 4:
			bitCount = 4;

		lengthMask = 0xFFFF >> bitCount
		offSetMask = ~lengthMask
		maximumLength = (0xFFFF >> bitCount) + 3

		return lengthMask, offSetMask, bitCount, maximumLength

class UnCompressedVBAStream(VBAStreamBase):
	def PackCopyToken(self, offset, length):
		lengthMask, offSetMask, bitCount, maximumLength = self.CopyTokenHelp()
		temp1 = offset - 1
		temp2 = 16 - bitCount
		temp3 = length - 3
		copyToken = (temp1 << temp2) | temp3
		return copyToken

	def CompressRawChunk(self):
		self.CompressedCurrent = self.CompressedChunkStart + 2
		self.DecompressedCurrent  = self.DecompressedChunkStart
		padCount = self.CHUNKSIZE
		lastByte = self.DecompressedChunkStart + padCount
		if self.DecompressedBufferEnd < lastByte:
		   lastByte =  self.DecompressedBufferEnd

		for index in xrange(self.DecompressedChunkStart,  lastByte):
			self.CompressedContainer[self.CompressedCurrent] = self.data[index]
			self.CompressedCurrent = self.CompressedCurrent + 1
			self.DecompressedCurrent = self.DecompressedCurrent + 1
			padCount = padCount - 1

		for index in xrange(0, padCount):
			self.CompressedContainer[self.CompressedCurrent] = 0x0;
			self.CompressedCurrent = self.CompressedCurrent + 1

	def Matching(self, decompressedEnd):
		candidate = self.DecompressedCurrent - 1
		bestLength = 0
		while candidate >= self.DecompressedChunkStart:
			C = candidate
			D = self.DecompressedCurrent
			Len = 0
			while D < decompressedEnd and (self.data[D] == self.data[C]):
				Len = Len + 1
				C = C + 1
				D = D + 1
			if Len > bestLength:
				bestLength = Len
				bestCandidate = candidate
			candidate = candidate - 1
		if bestLength >=  3:
			lengthMask, offSetMask, bitCount, maximumLength = self.CopyTokenHelp()
			length = bestLength
			if (maximumLength < bestLength):
				length = maximumLength
			offset = self.DecompressedCurrent - bestCandidate
		else:
			length = 0
			offset = 0
		return offset, length

	def CompressToken(self, compressedEnd, decompressedEnd, index, flags):
		offset = 0
		offset, length = self.Matching(decompressedEnd)
		if offset:
			if (self.CompressedCurrent + 1) < compressedEnd:
				copyToken = self.PackCopyToken(offset, length)
				struct.pack_into("<H", self.CompressedContainer, self.CompressedCurrent, copyToken )

				# Set flag bit
				temp1 = 1 << index
				temp2 = flags & ~temp1
				flags = temp2 | temp1

				self.CompressedCurrent = self.CompressedCurrent + 2
				self.DecompressedCurrent = self.DecompressedCurrent + length
			else:
				self.CompressedCurrent = compressedEnd
		else:
			if self.CompressedCurrent < compressedEnd:
				self.CompressedContainer[self.CompressedCurrent] = self.data[self.DecompressedCurrent]
				self.CompressedCurrent = self.CompressedCurrent + 1
				self.DecompressedCurrent = self.DecompressedCurrent + 1
			else:
				self.CompressedCurrent = compressedEnd

		return flags

	def CompressTokenSequence(self, compressedEnd, decompressedEnd):
		flagByteIndex = self.CompressedCurrent
		tokenFlags = 0
		self.CompressedCurrent = self.CompressedCurrent + 1
		for index in xrange(0,8):
			if ((self.DecompressedCurrent < decompressedEnd) and (self.CompressedCurrent < compressedEnd)):
				tokenFlags = self.CompressToken(compressedEnd, decompressedEnd, index, tokenFlags)
		self.CompressedContainer[flagByteIndex] = tokenFlags

	def CompressDecompressedChunk(self):
		self.CompressedContainer.extend(bytearray(self.CHUNKSIZE + 2))
		compressedEnd = self.CompressedChunkStart + 4098
		self.CompressedCurrent = self.CompressedChunkStart + 2
		decompressedEnd = self.DecompressedBufferEnd
		if (self.DecompressedChunkStart + self.CHUNKSIZE) < self.DecompressedBufferEnd:
			decompressedEnd = (self.DecompressedChunkStart + self.CHUNKSIZE)

		while (self.DecompressedCurrent < decompressedEnd) and (self.CompressedCurrent < compressedEnd):
				self.CompressTokenSequence(compressedEnd, decompressedEnd)

		if self.DecompressedCurrent < decompressedEnd:
			self.CompressRawChunk(decompressedEnd - 1)
			compressedFlag = 0
		else:
			compressedFlag = 1
		size = self.CompressedCurrent - self.CompressedChunkStart
		header = 0x0000

		# Pack compressed chunk size
		temp1 = header & 0xF000
		temp2 = size - 3
		header = temp1 | temp2

		# Pack compressed chunk flag
		temp1 = header & 0x7FFF
		temp2 = compressedFlag << 15
		header = temp1 | temp2

		# Pack compressed chunk signature
		temp1 = header & 0x8FFF
		Header = temp1 | 0x3000

		struct.pack_into("<H", self.CompressedContainer, self.CompressedChunkStart, Header)

		if (self.CompressedCurrent):
			self.CompressedContainer = self.CompressedContainer[0:self.CompressedCurrent]

	def Compress(self):
		self.CompressedContainer = bytearray()
		self.CompressedCurrent = 0
		self.CompressedChunkStart = 0
		self.DecompressedCurrent = 0
		self.DecompressedBufferEnd = len(self.data)
		self.DecompressedChunkStart = 0

		SignatureByte = 0x01
		self.CompressedContainer.append(SignatureByte)
		self.CompressedCurrent = self.CompressedCurrent + 1

		while self.DecompressedCurrent < self.DecompressedBufferEnd:
			self.CompressedChunkStart = self.CompressedCurrent
			self.DecompressedChunkStart = self.DecompressedCurrent
			self.CompressDecompressedChunk()

		return self.CompressedContainer

if __name__ == '__main__':
	main()
