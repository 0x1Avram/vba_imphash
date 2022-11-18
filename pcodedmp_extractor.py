import itertools
import sys
from struct import unpack_from


"""
Wrapper over the getTheIdentifiers function.
'vba_project_stream' parameter is a bytes object.
"""
def get_all_identifiers(vba_project_stream):
    identifiers = getTheIdentifiers(vba_project_stream)
    print(f'\t[PCODEDMP] All Identifiers = {identifiers}.')
    return identifiers


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
vbaProjectData parameter is a bytes object.
Added some commented comments at the end.
"""
def getTheIdentifiers(vbaProjectData):
    identifiers = []
    try:
        magic = getWord(vbaProjectData, 0, '<')
        if magic != 0x61CC:
            return identifiers
        version = getWord(vbaProjectData, 2, '<')
        unicodeRef  = (version >= 0x5B) and (not version in [0x60, 0x62, 0x63]) or (version == 0x4E)
        unicodeName = (version >= 0x59) and (not version in [0x60, 0x62, 0x63]) or (version == 0x4E)
        nonUnicodeName = ((version <= 0x59) and (version != 0x4E)) or (0x5F > version > 0x6B)
        word = getWord(vbaProjectData, 5, '<')
        if word == 0x000E:
            endian = '>'
        else:
            endian = '<'
        offset = 0x1E
        offset, numRefs = getVar(vbaProjectData, offset, endian, False)
        offset += 2
        for _ in itertools.repeat(None, numRefs):
            offset, refLength = getVar(vbaProjectData, offset, endian, False)
            if refLength == 0:
                offset += 6
            else:
                if ((unicodeRef and (refLength < 5)) or ((not unicodeRef) and (refLength < 3))):
                    offset += refLength
                else:
                    if unicodeRef:
                        c = vbaProjectData[offset + 4]
                    else:
                        c = vbaProjectData[offset + 2]
                    offset += refLength
                    if chr(ord(c)) in ['C', 'D']:
                        offset = skipStructure(vbaProjectData, offset, endian, False, 1, False)
            offset += 10
            offset, word = getVar(vbaProjectData, offset, endian, False)
            if word:
                offset = skipStructure(vbaProjectData, offset, endian, False, 1, False)
                offset, wLength = getVar(vbaProjectData, offset, endian, False)
                if wLength:
                    offset += 2
                offset += wLength + 30
        # Number of entries in the class/user forms table
        offset = skipStructure(vbaProjectData, offset, endian, False, 2, False)
        # Number of compile-time identifier-value pairs
        offset = skipStructure(vbaProjectData, offset, endian, False, 4, False)
        offset += 2
        # Typeinfo typeID
        offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
        # Project description
        offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
        # Project help file name
        offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
        offset += 0x64
        # Skip the module descriptors
        offset, numProjects = getVar(vbaProjectData, offset, endian, False)
        for _ in itertools.repeat(None, numProjects):
            offset, wLength = getVar(vbaProjectData, offset, endian, False)
            # Code module name
            if unicodeName:
                offset += wLength
            if nonUnicodeName:
                if wLength:
                    offset, wLength = getVar(vbaProjectData, offset, endian, False)
                offset += wLength
            # Stream time
            offset = skipStructure(vbaProjectData, offset, endian, False, 1, False)
            offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
            offset, _ = getVar(vbaProjectData, offset, endian, False)
            if version >= 0x6B:
                offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
            offset = skipStructure(vbaProjectData, offset, endian, False, 1, True)
            offset += 2
            if version != 0x51:
                offset += 4
            offset = skipStructure(vbaProjectData, offset, endian, False, 8, False)
            offset += 11
        offset += 6
        offset = skipStructure(vbaProjectData, offset, endian, True, 1, False)
        offset += 6
        offset, w0 = getVar(vbaProjectData, offset, endian, False)
        offset, numIDs = getVar(vbaProjectData, offset, endian, False)
        offset, w1 = getVar(vbaProjectData, offset, endian, False)
        offset += 4
        numJunkIDs = numIDs + w1 - w0
        numIDs = w0 - w1
        # Skip the junk IDs
        for _ in itertools.repeat(None, numJunkIDs):
            offset += 4
            idType, idLength = getTypeAndLength(vbaProjectData, offset, endian)
            offset += 2
            if idType > 0x7F:
                offset += 6
            offset += idLength
        # Now offset points to the start of the variable names area
        i = 0
        for _ in itertools.repeat(None, numIDs):
            i += 1
            start_offset = offset
            isKwd = False
            ident = ''
            idType, idLength = getTypeAndLength(vbaProjectData, offset, endian)
            offset += 2
            if (idLength == 0) and (idType == 0):
                offset += 2
                idType, idLength = getTypeAndLength(vbaProjectData, offset, endian)
                offset += 2
                isKwd = True
            if idType & 0x80:
                offset += 6
            if idLength:
                ident = decode(vbaProjectData[offset:offset + idLength])
                identifiers.append(ident)
                offset += idLength
            if not isKwd:
                offset += 4
            # end_offset = offset
            # print(f'[PCODEDMP][IDENTIFIERS] i = {i}: ident = {ident}; isKwd = {isKwd}; '
            #      f'idType = {hex(idType)}; idLength = {hex(idLength)}; '
            #      f'start_offset = {hex(start_offset)}; end_offset = {hex(end_offset)}.\n'
            #      f'{hexdump(vbaProjectData[start_offset:end_offset])}')
    except Exception as e:
        print('[PCODEDMP] Error: {}.'.format(e), file=sys.stderr)
    return identifiers


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
def getWord(buffer, offset, endian):
    return unpack_from(endian + 'H', buffer, offset)[0]


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
def getVar(buffer, offset, endian, isDWord):
    if isDWord:
        value = getDWord(buffer, offset, endian)
        offset += 4
    else:
        value = getWord(buffer, offset, endian)
        offset += 2
    return offset, value


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
def getDWord(buffer, offset, endian):
    return unpack_from(endian + 'L', buffer, offset)[0]


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
def skipStructure(buffer, offset, endian, isLengthDW, elementSize, checkForMinusOne):
    if isLengthDW:
        length = getDWord(buffer, offset, endian)
        offset += 4
        skip = checkForMinusOne and (length == 0xFFFFFFFF)
    else:
        length = getWord(buffer, offset, endian)
        offset += 2
        skip = checkForMinusOne and (length == 0xFFFF)
    if not skip:
        offset += length * elementSize
    return offset


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
def getTypeAndLength(buffer, offset, endian):
    if endian == '>':
        return ord(buffer[offset]), ord(buffer[offset + 1])
    else:
        return ord(buffer[offset + 1]), ord(buffer[offset])


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
codec = 'latin1'    # Assume 'latin1' unless redefined by the 'dir' stream
def decode(x):
    return x.decode(codec, errors='replace')


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
Changed xrange -> range
"""
def hexdump(buffer, length=16):
    theHex = lambda data: ' '.join('{:02X}'.format(ord(i)) for i in data)
    theStr = lambda data: ''.join(chr(ord(i)) if (31 < ord(i) < 127) else '.' for i in data)
    result = ''
    for offset in range(0, len(buffer), length):
        data = buffer[offset:offset + length]
        result += '{:08X}   {:{}}    {}\n'.format(offset, theHex(data), length * 3 - 1, theStr(data))
    return result


"""
Code from https://github.com/bontchev/pcodedmp
LICENSE: GNU General Public License v3.0
"""
def ord(x):
    return x
