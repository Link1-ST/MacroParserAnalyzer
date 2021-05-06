#____________________________________________________________________________________
'''
@author: Saarthik Tannan

This program works with Python 3

This program currently supports: MS Office 2003 file types -- .doc and .xls

Note: Portions of the code below are taken from olevba and officeparser

olevba: https://github.com/decalage2/oletools/blob/master/oletools/olevba.py

officeparser: https://github.com/unixfreak0037/officeparser

'''
#____________________________________________________________________________________

# === LICENSE ==================================================================

# olevba is copyright (c) 2014-2021 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.


# olevba contains modified source code from the officeparser project, published
# under the following MIT License (MIT):
#
# officeparser is copyright (c) 2014 John William Davison
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#____________________________________________________________________________________

# Imports
try:
    import sys
    import os
    import logging
    import struct
    from io import BytesIO
    import math
    import errno
    import olefile
    from oletools import olevba

    
# Print message if error   
except ImportError as e:
    print("Error: ", str(e))

# Global variables
# Command line option
global cmdOption
# The file path of the macro file
global filename
# Stream name containing the VBA module
global streamNameStr
# VBA source code in str format
global codeStr
# VBA module name
global moduleNameStr
# VBA module file name which contains an extension based on the module type such as bas, cls, frm
global moduleFilenameStr
# VBA module path
global codePath
# Store decompressed data from dir stream
global dirStream
# The file application (i.e. Word, Excel, or PowerPoint)
global filetype
# The VBA root directory
global vbaRootDir
# The module extensions
global moduleExt
# A list which contain the following modules: (stream_path, vba_filename, vba_code)
global modules
# OLE file
global ole
# A list which contains the modules needed for to be included with the analysis results
global modules2 
# A list which contains the modules (codePath, and moduleFilenameStr)
global modules2a

'''
Windows-specific error code indicating an invalid pathname.

See Also
----------
https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes--0-499-
    Official listing of all such codes.
'''
ERROR_INVALID_NAME = 123

def helpMenu():
    # Purpose: Print the help menu
    print("\n")
    print("MacroParserAnalyzer Help Menu");
    print("--------------------------------------------------------------\n");
    print("Usage: python MacroParserAnalyzer.py [options] <filename>\n");
    print("Options: \n");
    print("-h, --help     \t shows the help menu\n");
    print("-c, --code     \t display only the VBA source code\n");
    print("-a, --analysis \t display only the analysis results\n");
    print("-d, --detailed \t display full results\n");
    print("--------------------------------------------------------------\n");

def is_pathname_valid(pathname: str) -> bool:
    '''
    `True` if the passed pathname is a valid pathname for the current OS;
    `False` otherwise.
    '''
    # If this pathname is either not a string or is but is empty, this pathname
    # is invalid.
    try:
        if not isinstance(pathname, str) or not pathname:
            return False

        # Strip this pathname's Windows-specific drive specifier (e.g., `C:\`)
        # if any. Since Windows prohibits path components from containing `:`
        # characters, failing to strip this `:`-suffixed prefix would
        # erroneously invalidate all valid absolute Windows pathnames.
        _, pathname = os.path.splitdrive(pathname)

        # Directory guaranteed to exist. If the current OS is Windows, this is
        # the drive to which Windows was installed (e.g., the "%HOMEDRIVE%"
        # environment variable); else, the typical root directory.
        root_dirname = os.environ.get('HOMEDRIVE', 'C:') \
            if sys.platform == 'win32' else os.path.sep
        assert os.path.isdir(root_dirname)   # ...Murphy and her ironclad Law

        # Append a path separator to this directory if needed.
        root_dirname = root_dirname.rstrip(os.path.sep) + os.path.sep

        # Test whether each path component split from this pathname is valid or
        # not, ignoring non-existent and non-readable path components.
        for pathname_part in pathname.split(os.path.sep):
            try:
                os.lstat(root_dirname + pathname_part)
            # If an OS-specific exception is raised, its error code
            # indicates whether this pathname is valid or not. Unless this
            # is the case, this exception implies an ignorable kernel or
            # filesystem complaint (e.g., path not found or inaccessible).
            #
            # Only the following exceptions indicate invalid pathnames:
            #
            # * Instances of the Windows-specific "WindowsError" class
            #   defining the "winerror" attribute whose value is
            #   "ERROR_INVALID_NAME". Under Windows, "winerror" is more
            #   fine-grained and hence useful than the generic "errno"
            #   attribute. When a too-long pathname is passed, for example,
            #   "errno" is "ENOENT" (i.e., no such file or directory) rather
            #   than "ENAMETOOLONG" (i.e., file name too long).
            # * Instances of the cross-platform "OSError" class defining the
            #   generic "errno" attribute whose value is either:
            #   * Under most POSIX-compatible OSes, "ENAMETOOLONG".
            #   * Under some edge-case OSes (e.g., SunOS, *BSD), "ERANGE".
            except OSError as exc:
                if hasattr(exc, 'winerror'):
                    if exc.winerror == ERROR_INVALID_NAME:
                        return False
                elif exc.errno in {errno.ENAMETOOLONG, errno.ERANGE}:
                    return False
    # If a "TypeError" exception was raised, it almost certainly has the
    # error message "embedded NUL character" indicating an invalid pathname.
    except TypeError as exc:
        return False
    # If no exception was raised, all path components and hence this
    # pathname itself are valid. (Praise be to the curmudgeonly python.)
    else:
        return True
    # If any other exception was raised, this is an unrelated fatal issue
    # (e.g., a bug). Permit this exception to unwind the call stack.
    #
    # Did we mention this should be shipped with Python already?

def is_path_creatable(pathname: str) -> bool:
    '''
    `True` if the current user has sufficient permissions to create the passed
    pathname; `False` otherwise.
    '''
    # Parent directory of the passed path. If empty, we substitute the current
    # working directory (CWD) instead.
    dirname = os.path.dirname(pathname) or os.getcwd()
    return os.access(dirname, os.W_OK)

def is_path_exists_or_creatable(pathname: str) -> bool:
    '''
    `True` if the passed pathname is a valid pathname for the current OS _and_
    either currently exists or is hypothetically creatable; `False` otherwise.

    This function is guaranteed to _never_ raise exceptions.
    '''
    try:
        # To prevent "os" module calls from raising undesirable exceptions on
        # invalid pathnames, is_pathname_valid() is explicitly called first.
        return is_pathname_valid(pathname) and (
            os.path.exists(pathname) or is_path_creatable(pathname))
    # Report failure on non-fatal filesystem complaints (e.g., connection
    # timeouts, permissions issues) implying this path to be inaccessible. All
    # other exceptions are unrelated fatal issues and should not be caught here.
    except OSError:
        return False

def check_value(name, expected, value):
    # Purpose: Check if value is correct        
    if expected != value:
        logging.error("invalid value for {0} expected {1:04X} got {2:04X}".format(name, expected, value))
                
                    
def copytoken_help(decompressed_current, decompressed_chunk_start):
    """
    Purpose: Compute bit masks to decode a CopyToken according to MS-OVBA 2.4.1.3.19.1 CopyToken Help

    decompressed_current: number of decompressed bytes so far, i.e. len(decompressed_container)
    decompressed_chunk_start: offset of the current chunk in the decompressed container
    return length_mask, offset_mask, bit_count, maximum_length
    """
    difference = decompressed_current - decompressed_chunk_start
    bit_count = int(math.ceil(math.log(difference, 2)))
    bit_count = max([bit_count, 4])
    length_mask = 0xFFFF >> bit_count
    offset_mask = ~length_mask
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length

def decompress_stream(compressed_container):
    """
    Purpose: Decompress a stream according to MS-OVBA section 2.4.1

    :param compressed_container bytearray: bytearray or bytes compressed according to the MS-OVBA 2.4.1.3.6 Compression algorithm
    :return: the decompressed container as a bytes string
    :rtype: bytes
    """
    # 2.4.1.2 State Variables

    # The following state is maintained for the CompressedContainer (section 2.4.1.1.1):
    # CompressedRecordEnd: The location of the byte after the last byte in the CompressedContainer (section 2.4.1.1.1).
    # CompressedCurrent: The location of the next byte in the CompressedContainer (section 2.4.1.1.1) to be read by
    #                    decompression or to be written by compression.

    # The following state is maintained for the current CompressedChunk (section 2.4.1.1.4):
    # CompressedChunkStart: The location of the first byte of the CompressedChunk (section 2.4.1.1.4) within the
    #                       CompressedContainer (section 2.4.1.1.1).

    # The following state is maintained for a DecompressedBuffer (section 2.4.1.1.2):
    # DecompressedCurrent: The location of the next byte in the DecompressedBuffer (section 2.4.1.1.2) to be written by
    #                      decompression or to be read by compression.
    # DecompressedBufferEnd: The location of the byte after the last byte in the DecompressedBuffer (section 2.4.1.1.2).

    # The following state is maintained for the current DecompressedChunk (section 2.4.1.1.3):
    # DecompressedChunkStart: The location of the first byte of the DecompressedChunk (section 2.4.1.1.3) within the
    #                         DecompressedBuffer (section 2.4.1.1.2).

    # Check the input is a bytearray, otherwise convert it (assuming it's bytes):
    if not isinstance(compressed_container, bytearray):
        compressed_container = bytearray(compressed_container)
    
    decompressed_container = bytearray()  # result
    compressed_current = 0

    sig_byte = compressed_container[compressed_current]
    if sig_byte != 0x01:
        raise ValueError('invalid signature byte {0:02X}'.format(sig_byte))

    compressed_current += 1

    #NOTE: the definition of CompressedRecordEnd is ambiguous. Here we assume that
    # CompressedRecordEnd = len(compressed_container)
    while compressed_current < len(compressed_container):
        # 2.4.1.1.5
        compressed_chunk_start = compressed_current
        # chunk header = first 16 bits
        compressed_chunk_header = \
            struct.unpack("<H", compressed_container[compressed_chunk_start:compressed_chunk_start + 2])[0]
        # chunk size = 12 first bits of header + 3
        chunk_size = (compressed_chunk_header & 0x0FFF) + 3
        # chunk signature = 3 next bits - should always be 0b011
        chunk_signature = (compressed_chunk_header >> 12) & 0x07
        if chunk_signature != 0b011:
            raise ValueError('Invalid CompressedChunkSignature in VBA compressed stream')
        # chunk flag = next bit - 1 == compressed, 0 == uncompressed
        chunk_flag = (compressed_chunk_header >> 15) & 0x01

        #MS-OVBA 2.4.1.3.12: the maximum size of a chunk including its header is 4098 bytes (header 2 + data 4096)
        # The minimum size is 3 bytes
        # NOTE: there seems to be a typo in MS-OVBA, the check should be with 4098, not 4095 (which is the max value
        # in chunk header before adding 3.
        # Also the first test is not useful since a 12 bits value cannot be larger than 4095.
        if chunk_flag == 1 and chunk_size > 4098:
            raise ValueError('CompressedChunkSize=%d > 4098 but CompressedChunkFlag == 1' % chunk_size)
        if chunk_flag == 0 and chunk_size != 4098:
            raise ValueError('CompressedChunkSize=%d != 4098 but CompressedChunkFlag == 0' % chunk_size)

        # get the end of the compressed data:
        compressed_end = min([len(compressed_container), compressed_chunk_start + chunk_size])
        # read after chunk header:
        compressed_current = compressed_chunk_start + 2

        if chunk_flag == 0:
            # MS-OVBA 2.4.1.3.3 Decompressing a RawChunk
            # uncompressed chunk: read the next 4096 bytes as-is
            decompressed_container.extend(compressed_container[compressed_current:compressed_current + 4096])
            compressed_current += 4096
        else:
            # MS-OVBA 2.4.1.3.2 Decompressing a CompressedChunk
            # compressed chunk
            decompressed_chunk_start = len(decompressed_container)
            while compressed_current < compressed_end:
                # MS-OVBA 2.4.1.3.4 Decompressing a TokenSequence
                # log.debug('compressed_current = %d / compressed_end = %d' % (compressed_current, compressed_end))
                # FlagByte: 8 bits indicating if the following 8 tokens are either literal (1 byte of plain text) or
                # copy tokens (reference to a previous literal token)
                flag_byte = compressed_container[compressed_current]
                compressed_current += 1
                for bit_index in range(0, 8):
                    if compressed_current >= compressed_end:
                        break
                    # MS-OVBA 2.4.1.3.5 Decompressing a Token
                    # MS-OVBA 2.4.1.3.17 Extract FlagBit
                    flag_bit = (flag_byte >> bit_index) & 1
                    if flag_bit == 0:  # LiteralToken
                        # copy one byte directly to output
                        decompressed_container.extend([compressed_container[compressed_current]])
                        compressed_current += 1
                    else:  # CopyToken
                        # MS-OVBA 2.4.1.3.19.2  CopyToken
                        copy_token = \
                            struct.unpack("<H", compressed_container[compressed_current:compressed_current + 2])[0]

                        length_mask, offset_mask, bit_count, _ = copytoken_help(
                            len(decompressed_container), decompressed_chunk_start)
                        length = (copy_token & length_mask) + 3
                        temp1 = copy_token & offset_mask
                        temp2 = 16 - bit_count
                        offset = (temp1 >> temp2) + 1
                        copy_source = len(decompressed_container) - offset
                        for index in range(copy_source, copy_source + length):
                            decompressed_container.extend([decompressed_container[index]])
                        compressed_current += 2
    return bytes(decompressed_container)


def decodeBytes(bytesString, errors='replace'):
        """
        Purpose: Decode a bytes string to a unicode string
        :param bytes_string: bytes, bytes string to be decoded
        :param errors: str, mode to handle unicode conversion errors
        :return: str/unicode, decoded string
        """        
        return bytesString.decode("utf-8", errors=errors)
    
    
def parseVBAProjectStream (ole, projectStreamPath, dirStream):
    # Purpose: Parse the PROJECT stream from the VBA project
    # :params: OleFileIO object  and the PROJECT stream path
    # :return: A dictionary of module extensions
    
    # Reference: [MS-OVBA] 2.3.1 PROJECT Stream
    # Sample content of the PROJECT stream:
        ##    ID="{5312AC8A-349D-4950-BDD0-49BE3C4DD0F0}"
        ##    Document=ThisDocument/&H00000000
        ##    Module=NewMacros
        ##    Name="Project"
        ##    HelpContextID="0"
        ##    VersionCompatible32="393222000"
        ##    CMG="F1F301E705E705E705E705"
        ##    DPB="8F8D7FE3831F2020202020"
        ##    GC="2D2FDD81E51EE61EE6E1"
        ##
        ##    [Host Extender Info]
        ##    &H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000
        ##    &H00000002={000209F2-0000-0000-C000-000000000046};Word8.0;&H00000000
        ##
        ##    [Workspace]
        ##    ThisDocument=22, 29, 339, 477, Z
        ##    NewMacros=-4, 42, 832, 510, C
    
    # Global variable
    global moduleExt
    
    # File extensions for files in the PROJECT stream
    MODULE_EXTENSION = "bas"
    CLASS_EXTENSION = "cls"
    FORM_EXTENSION = "frm"    
    
    # Loop counter
    loopCounter = 0
    
    # Open the PROJECT stream
    projectStream = ole.openstream(projectStreamPath)
    
    moduleExt = {}
    
    for line in projectStream:
        # Decode a bytes string to a unicode string
        line = decodeBytes(line)
        line = line.strip()
        if '=' in line:
            # split line at the 1st equal sign:
            name, value = line.split('=', 1)
            # looking for code modules
            # add the code module as a key in the dictionary
            # the value will be the extension needed later
            # The value is converted to lowercase, to allow case-insensitive matching (issue #3)
            value = value.lower()
            if name == 'Document':
                # split value at the 1st slash, keep 1st part:
                value = value.split('/', 1)[0]
                moduleExt[value] = CLASS_EXTENSION
            elif name == 'Module':
                moduleExt[value] = MODULE_EXTENSION
            elif name == 'Class':
                moduleExt[value] = CLASS_EXTENSION
            elif name == 'BaseClass':
                moduleExt[value] = FORM_EXTENSION
    
        loopCounter +=1  
        if (loopCounter > 1):
            break
           
        # Get the Project records
        # PROJECTSYSKIND Record
        PROJECTSYSKIND_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTSYSKIND_Id', 0x0001, PROJECTSYSKIND_Id)
        PROJECTSYSKIND_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTSYSKIND_Size', 0x0004, PROJECTSYSKIND_Size)
        PROJECTSYSKIND_SysKind = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTSYSKIND_SysKind == 0x00:
            logging.debug("16-bit Windows")
        elif PROJECTSYSKIND_SysKind == 0x01:
            logging.debug("32-bit Windows")
        elif PROJECTSYSKIND_SysKind == 0x02:
            logging.debug("Macintosh")
        elif PROJECTSYSKIND_SysKind == 0x03:
            logging.debug("64-bit Windows")
        else:
            logging.error("invalid PROJECTSYSKIND_SysKind {0:04X}".format(PROJECTSYSKIND_SysKind))

        # PROJECTLCID Record
        PROJECTLCID_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTLCID_Id', 0x0002, PROJECTLCID_Id)
        PROJECTLCID_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTLCID_Size', 0x0004, PROJECTLCID_Size)
        PROJECTLCID_Lcid = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTLCID_Lcid', 0x409, PROJECTLCID_Lcid)

        # PROJECTLCIDINVOKE Record
        PROJECTLCIDINVOKE_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTLCIDINVOKE_Id', 0x0014, PROJECTLCIDINVOKE_Id)
        PROJECTLCIDINVOKE_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTLCIDINVOKE_Size', 0x0004, PROJECTLCIDINVOKE_Size)
        PROJECTLCIDINVOKE_LcidInvoke = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTLCIDINVOKE_LcidInvoke', 0x409, PROJECTLCIDINVOKE_LcidInvoke)

        # PROJECTCODEPAGE Record
        PROJECTCODEPAGE_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTCODEPAGE_Id', 0x0003, PROJECTCODEPAGE_Id)
        PROJECTCODEPAGE_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTCODEPAGE_Size', 0x0002, PROJECTCODEPAGE_Size)
        PROJECTCODEPAGE_CodePage = struct.unpack("<H", dirStream.read(2))[0]

        # PROJECTNAME Record
        PROJECTNAME_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTNAME_Id', 0x0004, PROJECTNAME_Id)
        PROJECTNAME_SizeOfProjectName = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTNAME_SizeOfProjectName < 1 or PROJECTNAME_SizeOfProjectName > 128:
            logging.error("PROJECTNAME_SizeOfProjectName value not in range: {0}".format(PROJECTNAME_SizeOfProjectName))
        PROJECTNAME_ProjectName = dirStream.read(PROJECTNAME_SizeOfProjectName)

        # PROJECTDOCSTRING Record
        PROJECTDOCSTRING_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTDOCSTRING_Id', 0x0005, PROJECTDOCSTRING_Id)
        PROJECTDOCSTRING_SizeOfDocString = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTNAME_SizeOfProjectName > 2000:
            logging.error("PROJECTDOCSTRING_SizeOfDocString value not in range: {0}".format(PROJECTDOCSTRING_SizeOfDocString))
        PROJECTDOCSTRING_DocString = dirStream.read(PROJECTDOCSTRING_SizeOfDocString)
        PROJECTDOCSTRING_Reserved = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTDOCSTRING_Reserved', 0x0040, PROJECTDOCSTRING_Reserved)
        PROJECTDOCSTRING_SizeOfDocStringUnicode = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTDOCSTRING_SizeOfDocStringUnicode % 2 != 0:
            logging.error("PROJECTDOCSTRING_SizeOfDocStringUnicode is not even")
        PROJECTDOCSTRING_DocStringUnicode = dirStream.read(PROJECTDOCSTRING_SizeOfDocStringUnicode)

        # PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
        PROJECTHELPFILEPATH_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTHELPFILEPATH_Id', 0x0006, PROJECTHELPFILEPATH_Id)
        PROJECTHELPFILEPATH_SizeOfHelpFile1 = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTHELPFILEPATH_SizeOfHelpFile1 > 260:
            logging.error("PROJECTHELPFILEPATH_SizeOfHelpFile1 value not in range: {0}".format(PROJECTHELPFILEPATH_SizeOfHelpFile1))
        PROJECTHELPFILEPATH_HelpFile1 = dirStream.read(PROJECTHELPFILEPATH_SizeOfHelpFile1)
        PROJECTHELPFILEPATH_Reserved = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTHELPFILEPATH_Reserved', 0x003D, PROJECTHELPFILEPATH_Reserved)
        PROJECTHELPFILEPATH_SizeOfHelpFile2 = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTHELPFILEPATH_SizeOfHelpFile2 != PROJECTHELPFILEPATH_SizeOfHelpFile1:
            logging.error("PROJECTHELPFILEPATH_SizeOfHelpFile1 does not equal PROJECTHELPFILEPATH_SizeOfHelpFile2")
        PROJECTHELPFILEPATH_HelpFile2 = dirStream.read(PROJECTHELPFILEPATH_SizeOfHelpFile2)
        if PROJECTHELPFILEPATH_HelpFile2 != PROJECTHELPFILEPATH_HelpFile1:
            logging.error("PROJECTHELPFILEPATH_HelpFile1 does not equal PROJECTHELPFILEPATH_HelpFile2")

        # PROJECTHELPCONTEXT Record
        PROJECTHELPCONTEXT_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTHELPCONTEXT_Id', 0x0007, PROJECTHELPCONTEXT_Id)
        PROJECTHELPCONTEXT_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTHELPCONTEXT_Size', 0x0004, PROJECTHELPCONTEXT_Size)
        PROJECTHELPCONTEXT_HelpContext = struct.unpack("<L", dirStream.read(4))[0]

        # PROJECTLIBFLAGS Record
        PROJECTLIBFLAGS_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTLIBFLAGS_Id', 0x0008, PROJECTLIBFLAGS_Id)
        PROJECTLIBFLAGS_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTLIBFLAGS_Size', 0x0004, PROJECTLIBFLAGS_Size)
        PROJECTLIBFLAGS_ProjectLibFlags = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTLIBFLAGS_ProjectLibFlags', 0x0000, PROJECTLIBFLAGS_ProjectLibFlags)

        # PROJECTVERSION Record
        PROJECTVERSION_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTVERSION_Id', 0x0009, PROJECTVERSION_Id)
        PROJECTVERSION_Reserved = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTVERSION_Reserved', 0x0004, PROJECTVERSION_Reserved)
        PROJECTVERSION_VersionMajor = struct.unpack("<L", dirStream.read(4))[0]
        PROJECTVERSION_VersionMinor = struct.unpack("<H", dirStream.read(2))[0]

        # PROJECTCONSTANTS Record
        PROJECTCONSTANTS_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTCONSTANTS_Id', 0x000C, PROJECTCONSTANTS_Id)
        PROJECTCONSTANTS_SizeOfConstants = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTCONSTANTS_SizeOfConstants > 1015:
            logging.error("PROJECTCONSTANTS_SizeOfConstants value not in range: {0}".format(PROJECTCONSTANTS_SizeOfConstants))
        PROJECTCONSTANTS_Constants = dirStream.read(PROJECTCONSTANTS_SizeOfConstants)
        PROJECTCONSTANTS_Reserved = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTCONSTANTS_Reserved', 0x003C, PROJECTCONSTANTS_Reserved)
        PROJECTCONSTANTS_SizeOfConstantsUnicode = struct.unpack("<L", dirStream.read(4))[0]
        if PROJECTCONSTANTS_SizeOfConstantsUnicode % 2 != 0:
            logging.error("PROJECTCONSTANTS_SizeOfConstantsUnicode is not even")
        PROJECTCONSTANTS_ConstantsUnicode = dirStream.read(PROJECTCONSTANTS_SizeOfConstantsUnicode)

        # array of REFERENCE records
        check = None
        while True:
            check = struct.unpack("<H", dirStream.read(2))[0]
            logging.debug("reference type = {0:04X}".format(check))
            if check == 0x000F:
                break

            if check == 0x0016:
                # REFERENCENAME
                REFERENCE_Id = check
                REFERENCE_SizeOfName = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCE_Name = dirStream.read(REFERENCE_SizeOfName)
                REFERENCE_Reserved = struct.unpack("<H", dirStream.read(2))[0]
                check_value('REFERENCE_Reserved', 0x003E, REFERENCE_Reserved)
                REFERENCE_SizeOfNameUnicode = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCE_NameUnicode = dirStream.read(REFERENCE_SizeOfNameUnicode)
                continue

            if check == 0x0033:
                # REFERENCEORIGINAL (followed by REFERENCECONTROL)
                REFERENCEORIGINAL_Id = check
                REFERENCEORIGINAL_SizeOfLibidOriginal = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEORIGINAL_LibidOriginal = dirStream.read(REFERENCEORIGINAL_SizeOfLibidOriginal)
                continue

            if check == 0x002F:
                # REFERENCECONTROL
                REFERENCECONTROL_Id = check
                REFERENCECONTROL_SizeTwiddled = struct.unpack("<L", dirStream.read(4))[0] # ignore
                REFERENCECONTROL_SizeOfLibidTwiddled = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCECONTROL_LibidTwiddled = dirStream.read(REFERENCECONTROL_SizeOfLibidTwiddled)
                REFERENCECONTROL_Reserved1 = struct.unpack("<L", dirStream.read(4))[0] # ignore
                check_value('REFERENCECONTROL_Reserved1', 0x0000, REFERENCECONTROL_Reserved1)
                REFERENCECONTROL_Reserved2 = struct.unpack("<H", dirStream.read(2))[0] # ignore
                check_value('REFERENCECONTROL_Reserved2', 0x0000, REFERENCECONTROL_Reserved2)
                # optional field
                check2 = struct.unpack("<H", dirStream.read(2))[0]
                if check2 == 0x0016:
                    REFERENCECONTROL_NameRecordExtended_Id = check
                    REFERENCECONTROL_NameRecordExtended_SizeofName = struct.unpack("<L", dirStream.read(4))[0]
                    REFERENCECONTROL_NameRecordExtended_Name = dirStream.read(REFERENCECONTROL_NameRecordExtended_SizeofName)
                    REFERENCECONTROL_NameRecordExtended_Reserved = struct.unpack("<H", dirStream.read(2))[0]
                    check_value('REFERENCECONTROL_NameRecordExtended_Reserved', 0x003E, REFERENCECONTROL_NameRecordExtended_Reserved)
                    REFERENCECONTROL_NameRecordExtended_SizeOfNameUnicode = struct.unpack("<L", dirStream.read(4))[0]
                    REFERENCECONTROL_NameRecordExtended_NameUnicode = dirStream.read(REFERENCECONTROL_NameRecordExtended_SizeOfNameUnicode)
                    REFERENCECONTROL_Reserved3 = struct.unpack("<H", dirStream.read(2))[0]
                else:
                    REFERENCECONTROL_Reserved3 = check2

                check_value('REFERENCECONTROL_Reserved3', 0x0030, REFERENCECONTROL_Reserved3)
                REFERENCECONTROL_SizeExtended = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCECONTROL_SizeOfLibidExtended = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCECONTROL_LibidExtended = dirStream.read(REFERENCECONTROL_SizeOfLibidExtended)
                REFERENCECONTROL_Reserved4 = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCECONTROL_Reserved5 = struct.unpack("<H", dirStream.read(2))[0]
                REFERENCECONTROL_OriginalTypeLib = dirStream.read(16)
                REFERENCECONTROL_Cookie = struct.unpack("<L", dirStream.read(4))[0]
                continue

            if check == 0x000D:
                # REFERENCEREGISTERED
                REFERENCEREGISTERED_Id = check
                REFERENCEREGISTERED_Size = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEREGISTERED_SizeOfLibid = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEREGISTERED_Libid = dirStream.read(REFERENCEREGISTERED_SizeOfLibid)
                REFERENCEREGISTERED_Reserved1 = struct.unpack("<L", dirStream.read(4))[0]
                check_value('REFERENCEREGISTERED_Reserved1', 0x0000, REFERENCEREGISTERED_Reserved1)
                REFERENCEREGISTERED_Reserved2 = struct.unpack("<H", dirStream.read(2))[0]
                check_value('REFERENCEREGISTERED_Reserved2', 0x0000, REFERENCEREGISTERED_Reserved2)
                continue

            if check == 0x000E:
                # REFERENCEPROJECT
                REFERENCEPROJECT_Id = check
                REFERENCEPROJECT_Size = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEPROJECT_SizeOfLibidAbsolute = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEPROJECT_LibidAbsolute = dirStream.read(REFERENCEPROJECT_SizeOfLibidAbsolute)
                REFERENCEPROJECT_SizeOfLibidRelative = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEPROJECT_LibidRelative = dirStream.read(REFERENCEPROJECT_SizeOfLibidRelative)
                REFERENCEPROJECT_MajorVersion = struct.unpack("<L", dirStream.read(4))[0]
                REFERENCEPROJECT_MinorVersion = struct.unpack("<H", dirStream.read(2))[0]
                continue

            logging.error('invalid or unknown check Id {0:04X}'.format(check))
            sys.exit(0)
            
        PROJECTMODULES_Id = check #struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTMODULES_Id', 0x000F, PROJECTMODULES_Id)
        PROJECTMODULES_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTMODULES_Size', 0x0002, PROJECTMODULES_Size)
        PROJECTMODULES_Count = struct.unpack("<H", dirStream.read(2))[0]
        PROJECTMODULES_ProjectCookieRecord_Id = struct.unpack("<H", dirStream.read(2))[0]
        check_value('PROJECTMODULES_ProjectCookieRecord_Id', 0x0013, PROJECTMODULES_ProjectCookieRecord_Id)
        PROJECTMODULES_ProjectCookieRecord_Size = struct.unpack("<L", dirStream.read(4))[0]
        check_value('PROJECTMODULES_ProjectCookieRecord_Size', 0x0002, PROJECTMODULES_ProjectCookieRecord_Size)
        PROJECTMODULES_ProjectCookieRecord_Cookie = struct.unpack("<H", dirStream.read(2))[0]
        
        logging.debug("parsing {0} modules".format(PROJECTMODULES_Count))
        
        parseModules(ole, dirStream, PROJECTMODULES_Count)
        
def parseModules(ole, dirStream, PROJECTMODULES_Count):
    # Purpose: Parse the module records
    
    # Global variable
    global modules
    global modules2
    global modules2a
    global vbaRootDir
    global codePath
    global codeStr
    global moduleExt
    global moduleFilenameStr
    global output
    global analysisResults
    
    try:
        # A list which is used to store the module information (codePath, moduleFilenameStr, and codeStr)
        modules = []
        # A list which is used to store the module information (codePath, moduleFilenameStr)
        modules2 = []
        # A list that stores the analysis results
        modules2a = []
         
        # Loop through to get the MODULE records
        for x in range(0, PROJECTMODULES_Count):
                # 2.3.4.2.3.2.1 MODULENAME Record
                # Specifies a VBA identifier as the name of the containing MODULE Record
                MODULENAME_Id = struct.unpack("<H", dirStream.read(2))[0]
                check_value('MODULENAME_Id', 0x0019, MODULENAME_Id)
                MODULENAME_SizeOfModuleName = struct.unpack("<L", dirStream.read(4))[0]
                MODULENAME_ModuleName = dirStream.read(MODULENAME_SizeOfModuleName)
                # account for optional sections
                section_id = struct.unpack("<H", dirStream.read(2))[0]
                
                if section_id == 0x0047:
                    # 2.3.4.2.3.2.2 MODULENAMEUNICODE Record
                    # Specifies a VBA identifier as the name of the containing MODULE Record (section 2.3.4.2.3.2).
                    MODULENAMEUNICODE_Id = section_id
                    MODULENAMEUNICODE_SizeOfModuleNameUnicode = struct.unpack("<L", dirStream.read(4))[0]
                    MODULENAMEUNICODE_ModuleNameUnicode = dirStream.read(MODULENAMEUNICODE_SizeOfModuleNameUnicode).decode('UTF-16LE', 'replace')
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
                    
                if section_id == 0x001A:
                    # 2.3.4.2.3.2.3 MODULESTREAMNAME Record
                    # Specifies the stream name of the ModuleStream (section 2.3.4.3) in the VBA Storage (section 2.3.4)
                    # corresponding to the containing MODULE Record
                    MODULESTREAMNAME_id = section_id
                    MODULESTREAMNAME_SizeOfStreamName = struct.unpack("<L", dirStream.read(4))[0]
                    MODULESTREAMNAME_StreamName = dirStream.read(MODULESTREAMNAME_SizeOfStreamName)
                    MODULESTREAMNAME_Reserved = struct.unpack("<H", dirStream.read(2))[0]
                    check_value('MODULESTREAMNAME_Reserved', 0x0032, MODULESTREAMNAME_Reserved)
                    MODULESTREAMNAME_SizeOfStreamNameUnicode = struct.unpack("<L", dirStream.read(4))[0]
                    MODULESTREAMNAME_StreamNameUnicode = dirStream.read(MODULESTREAMNAME_SizeOfStreamNameUnicode).decode('UTF-16LE', 'replace')
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
                    
                if section_id == 0x001C:
                    # 2.3.4.2.3.2.4 MODULEDOCSTRING Record
                    # Specifies the description for the containing MODULE Record
                    MODULEDOCSTRING_Id = section_id
                    check_value('MODULEDOCSTRING_Id', 0x001C, MODULEDOCSTRING_Id)
                    MODULEDOCSTRING_SizeOfDocString = struct.unpack("<L", dirStream.read(4))[0]
                    MODULEDOCSTRING_DocString = dirStream.read(MODULEDOCSTRING_SizeOfDocString)
                    MODULEDOCSTRING_Reserved = struct.unpack("<H", dirStream.read(2))[0]
                    check_value('MODULEDOCSTRING_Reserved', 0x0048, MODULEDOCSTRING_Reserved)
                    MODULEDOCSTRING_SizeOfDocStringUnicode = struct.unpack("<L", dirStream.read(4))[0]
                    MODULEDOCSTRING_DocStringUnicode = dirStream.read(MODULEDOCSTRING_SizeOfDocStringUnicode)
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
                  
                if section_id == 0x0031:
                    # 2.3.4.2.3.2.5 MODULEOFFSET Record
                    # Specifies the location of the source code within the ModuleStream (section 2.3.4.3)
                    # that corresponds to the containing MODULE Record
                    MODULEOFFSET_Id = section_id
                    check_value('MODULEOFFSET_Id', 0x0031, MODULEOFFSET_Id)
                    MODULEOFFSET_Size = struct.unpack("<L", dirStream.read(4))[0]
                    check_value('MODULEOFFSET_Size', 0x0004, MODULEOFFSET_Size)
                    MODULEOFFSET_TextOffset = struct.unpack("<L", dirStream.read(4))[0]
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
    
                if section_id == 0x001E:
                    # 2.3.4.2.3.2.6 MODULEHELPCONTEXT Record
                    # Specifies the Help topic identifier for the containing MODULE Record
                    MODULEHELPCONTEXT_Id = section_id
                    check_value('MODULEHELPCONTEXT_Id', 0x001E, MODULEHELPCONTEXT_Id)
                    MODULEHELPCONTEXT_Size = struct.unpack("<L", dirStream.read(4))[0]
                    check_value('MODULEHELPCONTEXT_Size', 0x0004, MODULEHELPCONTEXT_Size)
                    # HelpContext (4 bytes): An unsigned integer that specifies the Help topic identifier
                    # in the Help file specified by PROJECTHELPFILEPATH Record
                    MODULEHELPCONTEXT_HelpContext = struct.unpack("<L", dirStream.read(4))[0]
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
                 
                if section_id == 0x002C:
                    # 2.3.4.2.3.2.7 MODULECOOKIE Record
                    # Specifies ignored data.
                    MODULECOOKIE_Id = section_id
                    check_value('MODULECOOKIE_Id', 0x002C, MODULECOOKIE_Id)
                    MODULECOOKIE_Size = struct.unpack("<L", dirStream.read(4))[0]
                    check_value('MODULECOOKIE_Size', 0x0002, MODULECOOKIE_Size)
                    MODULECOOKIE_Cookie = struct.unpack("<H", dirStream.read(2))[0]
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
    
                if section_id == 0x0021 or section_id == 0x0022:
                    # 2.3.4.2.3.2.8 MODULETYPE Record
                    # Specifies whether the containing MODULE Record (section 2.3.4.2.3.2) is a procedural module,
                    # document module, class module, or designer module.
                    # Id (2 bytes): An unsigned integer that specifies the identifier for this record.
                    # MUST be 0x0021 when the containing MODULE Record (section 2.3.4.2.3.2) is a procedural module.
                    # MUST be 0x0022 when the containing MODULE Record (section 2.3.4.2.3.2) is a document module,
                    # class module, or designer module.
                    MODULETYPE_Id = section_id
                    MODULETYPE_Reserved = struct.unpack("<L", dirStream.read(4))[0]
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
    
                if section_id == 0x0025:
                    # 2.3.4.2.3.2.9 MODULEREADONLY Record
                    # Specifies that the containing MODULE Record (section 2.3.4.2.3.2) is read-only.
                    MODULEREADONLY_Id = section_id
                    check_value('MODULEREADONLY_Id', 0x0025, MODULEREADONLY_Id)
                    MODULEREADONLY_Reserved = struct.unpack("<L", dirStream.read(4))[0]
                    check_value('MODULEREADONLY_Reserved', 0x0000, MODULEREADONLY_Reserved)
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
     
                if section_id == 0x0028:
                    # 2.3.4.2.3.2.10 MODULEPRIVATE Record
                    # Specifies that the containing MODULE Record (section 2.3.4.2.3.2) is only usable from within
                    # the current VBA project.
                    MODULEPRIVATE_Id = section_id
                    check_value('MODULEPRIVATE_Id', 0x0028, MODULEPRIVATE_Id)
                    MODULEPRIVATE_Reserved = struct.unpack("<L", dirStream.read(4))[0]
                    check_value('MODULEPRIVATE_Reserved', 0x0000, MODULEPRIVATE_Reserved)
                    section_id = struct.unpack("<H", dirStream.read(2))[0]
                    
                if section_id == 0x002B: # TERMINATOR
                    # Terminator (2 bytes): An unsigned integer that specifies the end of this record. MUST be 0x002B.
                    # Reserved (4 bytes): MUST be 0x00000000. MUST be ignored.
                    MODULE_Reserved = struct.unpack("<L", dirStream.read(4))[0]
                    check_value('MODULE_Reserved', 0x0000, MODULE_Reserved)
                    section_id = None
    
                if section_id != None:
                    logging.warning('unknown or invalid module section id {0:04X}'.format(section_id))
    
                logging.debug("ModuleName = {0}".format(MODULENAME_ModuleName))
                logging.debug("StreamName = {0}".format(MODULESTREAMNAME_StreamName))
                logging.debug("TextOffset = {0}".format(MODULEOFFSET_TextOffset))
    
                # Used to store VBA code
                codeData = None
                
                # Check for other stream and module names in case some are missing
                tryNames  = (str(MODULESTREAMNAME_StreamName, "utf-8"), MODULESTREAMNAME_StreamNameUnicode, MODULENAMEUNICODE_ModuleNameUnicode)
                # tryNames  = (MODULESTREAMNAME_StreamName, MODULESTREAMNAME_StreamNameUnicode, MODULENAMEUNICODE_ModuleNameUnicode)
                for streamName in tryNames:
                    if streamName is not None:
                        try:
                            codePath = vbaRootDir + "/" + 'VBA/' + streamName
                            logging.debug('opening VBA code stream %s' %codePath)
                            codeData = ole.openstream(codePath).read()
                            break
                            
                        except IOError as ioe:
                            logging.debug('failed to open stream VBA/%r (%r), try other name'
                                      % (streamName, ioe))
                
                # Log message if no code data
                if codeData is None:
                    logging.info("Could not open stream %d of %d ('VBA/' + one of %r)!"
                             % (x, PROJECTMODULES_Count,
                                '/'.join("'" + streamName + "'"
                                         for streamName in tryNames)))
                
                logging.debug("length of code_data = {0}".format(len(codeData)))
                logging.debug("offset of code_data = {0}".format(MODULEOFFSET_TextOffset))
                
                # Get VBA code
                codeData = codeData[MODULEOFFSET_TextOffset:]
                if len(codeData) > 0:
                    codeData = decompress_stream(bytearray(codeData))
                    # store the raw code encoded as bytes with the project's code page:
                    codeRaw = codeData
                    # Decode it to unicode str:
                    codeStr = decodeBytes(codeData)
                    
                    # Issue with getting the extensions above so need to set it based on the moduleName:
                    if (MODULENAMEUNICODE_ModuleNameUnicode == "ThisDocument" or MODULENAMEUNICODE_ModuleNameUnicode == "ThisWorkbook" 
                        or MODULENAMEUNICODE_ModuleNameUnicode == "Sheet1"):
                            fileExt = "cls"
                    elif (MODULENAMEUNICODE_ModuleNameUnicode == "NewMacros" or "Module1"):
                        fileExt = "bas"
                    
                    moduleFilenameStr = u'{0}.{1}'.format(MODULENAMEUNICODE_ModuleNameUnicode, fileExt)
                    
                    logging.debug('extracted file {0}'.format(moduleFilenameStr))
                    
                    # Store VBA Code information in list
                    modules.append(codePath)
                    modules.append(moduleFilenameStr)
                    modules.append(codeStr)
                    
                    # Do analyis if -a or --analysis entered
                    if (cmdOption == "-a" or cmdOption == "--analysis" or "-d" or "--detailed"):
                        modules2.append(codePath)
                        modules2.append(moduleFilenameStr)
                        modules2a.append(macroAnalysis())
                    
                else:
                    logging.warning("module stream {0} has code data length 0".format(MODULESTREAMNAME_StreamNameUnicode))

    except Exception as exc:
        logging.info('Error parsing module {0} of {1}:'
                     .format(x, PROJECTMODULES_Count),
                     exc_info=True)    
        raise
            
def printVBACode():
    # Purpose: Print VBA code for each OLE file
    # Global variable
    global modules
    global filename
    
    print("Message: For macros, the vba code 'Attribute VB_Name =' is hidden by Microsoft\n")
    
    
    print("VBA CODE")
    print("-------------------------------------------------------------------------------")
    print("File: %s" % filename)
    print("-------------------------------------------------------------------------------")
    
    for i in range(0, len(modules)):
        print("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
        if (modules[i] == ""):
            print("(empty macro)")
        else:
            print(modules[i])
        print("-------------------------------------------------------------------------------")
    
    print("-------------------------------------------------------------------------------")
    print("\n\n")
    
def writeVBACodeToFile(writeFile):
    # Purpose: Write VBA code to a file
    # Global variables
    global modules
    global filename
    
    # Write to a file
    with open(writeFile, "a") as fo:
        fo.write("Message: For macros, the vba code 'Attribute VB_Name =' is hidden by Microsoft\n")
        fo.write("VBA CODE\n")
        fo.write("-------------------------------------------------------------------------------\n")
        fo.write("File: %s\n" % filename)
        fo.write("-------------------------------------------------------------------------------\n")

        for i in range(0, len(modules)):
            fo.write("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\n")
            if (modules[i] == ""):
                fo.write("(empty macro)")
            else:
                fo.write(modules[i] + "\n")
            fo.write("-------------------------------------------------------------------------------\n")
    
        fo.write("-------------------------------------------------------------------------------\n")
        fo.write("\n\n")

def vbaCodeFromWord2003File(mFile):
    # Purpose: Parse VBA code from Word .doc file
    # Open an OLE file from disk
    with olefile.OleFileIO(mFile) as ole:
        # The dir stream path
        dirStreamPath = "Macros/VBA/dir"
        # Read the data from the dir stream (compressed)
        dirCompressed = ole.openstream(dirStreamPath).read()
        # Decompress the data
        dirStream =  BytesIO(decompress_stream(bytearray(dirCompressed)))
        
        # Parse the PROJECT stream from the VBA project
        projectStreamPath = "Macros/PROJECT"
        parseVBAProjectStream(ole, projectStreamPath, dirStream)
       
def vbaCodeFromExcel2003File(mFile):
    # Purpose Parse VBA code from Excel .xls file
    # Open an OLE file from disk
    with olefile.OleFileIO(mFile) as ole:
        # The dir stream path
        dirStreamPath = "_VBA_PROJECT_CUR/VBA/dir"
        # Read the data from the dir stream (compressed)
        dirCompressed = ole.openstream(dirStreamPath).read()
        # Decompress the data
        dirStream =  BytesIO(decompress_stream(bytearray(dirCompressed)))
        
        # Parse the PROJECT stream from the VBA project
        projectStreamPath = "_VBA_PROJECT_CUR/PROJECT"

        parseVBAProjectStream(ole, projectStreamPath, dirStream)

'''
def vbaCodeFromPowerPoint2003File(data):
    # Purpose Parse VBA code from PowerPoint .ppt file
    print("NOT FINISHED YET\n")
'''
        
'''
def vbaCodeFromOffic2007(data):
    # Purpose: Parse VBA from Offic 2007 files
    # Word, Excel, and PowerPoint files are Flat OPC 
    print("NOT Finished yet")
 '''   
    
def vbaCodeParser(mFile):
    # Purpose: Display or save the VBA source code to a file
    
    try:    
        # Global variable
        # Store file application (Word, Excel, or PowerPoint)
        global filetype
        # Store vba root directory
        global vbaRootDir

        # Local variables
        ''' MS Office 2003 file types (.doc, .xls, and .ppt) magic number: D0 CF 11 E0 A1 B1 1A E1        
            MS Office 2010 file types (.docm, .xlsm, and .pptm) magic number: 50 4B 03 04'''
        # Used for first two bytes of MS Office file type magic number
        partialMagicNumberMSO2003 = b"\xD0\xCF"
        partialMagicNumberMSO2007 = b"\x50\x4B"
        # Used to store file application type (i.e. Word, Excel, or PowerPoint) 
        #filetype = ""
        # fileContainsMacro will be set to True if the file contains macros
        fileContainsMacro = False
        '''fileContainsProjectStream and fileContainsDirStream will be set to True if the the VBA root contains 
        the VBA/_VBA_PROJECT, or PROJECT and VBA/Dir stream.'''
        fileContainsProjectStream = False
        fileContainsVBAProjectStream = False
        fileContainsDirStream = False
        
        # Parse the VBA source code    
        with open(mFile, "rb") as fi:
            # Get the first two bytes of the file (this is the partial magic number) 
            partialfileMagicNumber = fi.read(2)
            fi.read(6)
    
            # If the first two bytes of the file is DO CF then Microsoft 2003 file type
            if (partialMagicNumberMSO2003 == partialfileMagicNumber):
                # Open an OLE file from disk
                with olefile.OleFileIO(mFile) as ole:
                    ''' olefile.OleFileIO.exists() checks if a given stream or storage exists in the OLE file.
                    The provided path is case-insensitive.'''
                    '''
                    Find the file type by checking for key words (WordDocument, Workbook, and PowerPoint Document)
                    Find the VBA project root (different in MS Word, Excel, etc):
                      - Word 97-2003: Macros
                      - Excel 97-2003: _VBA_PROJECT_CUR
                      - PowerPoint 97-2003: VBA macros are stored within the binary structure of the presentation, not in an OLE storage
                    
                     According to MS-OVBA section 2.2.1:
                      - Word: the VBA project root storage MUST contain a VBA storage and a PROJECT stream 
                      - Excel: The root/VBA storage MUST contain a _VBA_PROJECT stream and a dir stream
                    '''
                    
                    # The file is a Word document. 
                    if (ole.exists("worddocument")): 
                        filetype = "Word"
                        # Check if contains a macro, storage and streams (project and dir)
                        if (ole.exists("macros")):
                            vbaRootDir = "Macros"
                        if (ole.exists("macros/vba")): 
                            fileContainsMacro = True
                        if (ole.exists("macros/project")):
                            fileContainsProjectStream = True
                        if (ole.exists ("macros/vba/_vba_project")):
                            fileContainsVBAProjectStream = True
                        if (ole.exists("macros/vba/dir")):
                            fileContainsDirStream = True

                    '''The file is an Excel document. Check if it contains a macro, storage, and streams (project and dir). 
                    Store the vba root directory '''
                    if (ole.exists("workbook")):
                        filetype = "Excel"
                        if(ole.exists("_vba_project_cur")):
                            vbaRootDir = "_VBA_PROJECT_CUR"
                        if (ole.exists("_vba_project_cur/vba")): 
                            fileContainsMacro = True
                        if (ole.exists("_vba_project_cur/project")):
                            fileContainsProjectStream = True
                        if (ole.exists("_vba_project_cur/vba/_vba_project")):
                            fileContainsVBAProjectStream = True
                        if (ole.exists("_vba_project_cur/vba/dir")):
                            fileContainsDirStream = True
                        
                    # The file is a PowerPoint document
                    if ole.exists("powerpoint document"):
                        filetype = "PowerPoint"
                        
                # Go back to the beginning of the file
                fi.seek(0)
                
                # Read the entire file
                data = fi.read()

                # The file is a Word doc file which contains macro
                if (filetype == "Word" and fileContainsMacro == True and fileContainsProjectStream == True 
                    and fileContainsVBAProjectStream == True and fileContainsDirStream == True):
                    vbaCodeFromWord2003File(mFile)
                
                # This is an Excel xls file which contains a macro
                if (filetype == "Excel" and fileContainsMacro == True and fileContainsProjectStream == True 
                    and fileContainsVBAProjectStream == True and fileContainsDirStream == True):
                    vbaCodeFromExcel2003File(mFile)
                
                # This is a PowerPoint ppt file which contains a macro. Not supported yet 
                if (filetype == "PowerPoint"):
                    print("PowerPoint ppt file is not supported yet\n")
                    sys.exit(0)
                    #vbaCodeFromPowerPoint2003File(data)    
                    
            # If The first two bytes of the file is 50 4B then Microsoft 2007 file type
            elif (partialMagicNumberMSO2007 == partialfileMagicNumber):
                print("Office 2007 files (docm, xlsm, and pptm) not supported yet\n")
                # Go back to the beginning of the file
                #fi.seek(0)
                #data = fi.read()
                sys.exit(0)

            # The first two bytes are something else
            else:
                print("Incorrect file type\n")
        
    # Print message if error                
    except Exception as e:
        print("Error: ", str(e))
        
def printAnalysisResults():
    # Purpose: Print analysis results to console
    # Global variable
    global module2
    global module2a
    global filename

    print("VBA Analysis Results")
    print("-------------------------------------------------------------------------------")
    print("File: %s" % filename)
    print("-------------------------------------------------------------------------------")
    
    for i in range(0, int(len(modules2)/2)):
        print("-------------------------------------------------------------------------------")
        print(modules2[2*i])
        print("-------------------------------------------------------------------------------")
        print(modules2[2*i+1])
        
        if (len(modules2a[i]) > 0):
            for j in range(0, len(modules2a[i])):
                print("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -")
                print("Type: " + modules2a[i][j][0])
                print("Keyword: " + modules2a[i][j][1])
                print("Description: " + modules2a[i][j][2])
    
    print("-------------------------------------------------------------------------------")
    print("\n\n")
    
def writeVBAAnalysisToFile(writeFile):
    # Global variable
    global module2
    global module2a
    global filename
    
    # Write to a file
    with open(writeFile, "a") as fo:
        fo.write("VBA Analysis Results\n")
        fo.write("-------------------------------------------------------------------------------\n")
        fo.write("File: %s\n" % filename)
        fo.write("-------------------------------------------------------------------------------\n")
        
        for i in range(0, int(len(modules2)/2)):
            fo.write("-------------------------------------------------------------------------------\n")
            fo.write(modules2[2*i] + "\n")
            fo.write("-------------------------------------------------------------------------------\n")
            fo.write(modules2[2*i+1] + "\n")
            
            if (len(modules2a[i]) > 0):
                for j in range(0, len(modules2a[i])):
                    fo.write("- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -\n")
                    fo.write("Type: " + modules2a[i][j][0] + "\n")
                    fo.write("Keyword: " + modules2a[i][j][1] + "\n")
                    fo.write("Description: " + modules2a[i][j][2] + "\n")
    
        fo.write("-------------------------------------------------------------------------------\n")
        fo.write("\n\n")

def macroAnalysis():
    """
    Purpose: Analyze the provided VBA code to detect suspicious keywords,
    auto-executable macros, IOC patterns, obfuscation patterns
    such as hex-encoded strings.
    """
    # Global variables
    global codeStr
    
    return olevba.scan_vba(codeStr, True)


# main function
def main():
    # Global variable
    global filename
    global cmdOption
    
    # Local variable
    argsList = sys.argv
    
    # Store the command line option
    if (len(argsList) > 1):
        cmdOption = argsList[1] 
    
    # If no arguments, -h, or --help option entered: Show help menu
    if (len(argsList) == 1 or argsList[1] == "-h" or argsList[1] == "--help"):
        helpMenu()
        
    # If -c or --code option entered: Display the vba code to console or save it to a file
    elif ((argsList[1] == "-c" or argsList[1] == "--code") and os.path.isfile(argsList[2])):
        filename = argsList[2]
        output = input("Enter output (console or file): ")
        print("\n")
                
        # Print VBA code to screen
        if (output == "console"):
            vbaCodeParser(filename)
            printVBACode()
            
        # Write VBA code to a file
        elif (output == "file"):
            # Ask the user for file to write 
            writeFile = input("Enter filename or path of file to write: ")
            
            ''' If user enters file or path to file which is writeable then get the VBA code and write it to the file'''
            if (writeFile != "" and is_path_exists_or_creatable(writeFile) == True):
                vbaCodeParser(filename)
                writeVBACodeToFile(writeFile)
            
            # Print message is user does not enter a file or path to file which is writeable. Do not get the VBA code
            else:
                print("Cannot write to file. Not getting VBA code\n")
            
        else:
            print("Incorrect output option\n")
            
    # If -a or --analysis option entered: Display the analysis results to console or save it to a file  
    elif ((argsList[1] == "-a" or argsList[1] == "--analysis") and os.path.isfile(argsList[2])):
        filename = argsList[2]
        output = input("Enter output (console or file): ")
        print("\n")
        if (output == "console"):
            # Get the VBA code
            # Note: The parseModules function calls the macroAnalysis function if -a or --analysis entered
            vbaCodeParser(filename)
            printAnalysisResults()
            
        elif(output == "file"):
            # Ask the user for file to write 
            writeFile = input("Enter filename or path of file to write: ")
            
            ''' If user enters file or path to file which is writeable then get the VBA code and write it to the file'''
            if (writeFile != "" and is_path_exists_or_creatable(writeFile) == True):
                vbaCodeParser(filename)
                writeVBAAnalysisToFile(writeFile)
            
            # Print message is user does not enter a file or path to file which is writeable. Do not get the VBA code
            else:
                print("Cannot write to file. Not getting VBA code\n")
            
        
        else:
            print("Incorrect output option\n")
             
    # If -d or --detailed option entered: Display the full results to console or save it to a file 
    elif ((argsList[1] == "-d" or argsList[1] == "--detailed") and os.path.isfile(argsList[2])):
        output = input("Enter output option (console or file): ")
        filename = argsList[2]
        print("\n")
        if (output == "console"):
            vbaCodeParser(filename)
            printVBACode()
            printAnalysisResults()
            
        elif(output == "file"):
            # Ask the user for file to write 
            writeFile = input("Enter filename or path of file to write: ")
            
            ''' If user enters file or path to file which is writeable then get the VBA code and write it to the file'''
            if (writeFile != "" and is_path_exists_or_creatable(writeFile) == True):
                vbaCodeParser(filename)
                writeVBACodeToFile(writeFile)
                writeVBAAnalysisToFile(writeFile)
        else:
            print("Incorrect output option\n")
    
    # Print message if any arguments are incorrect      
    else:
        print("Incorrect arguments\n")

if __name__ == '__main__':
    main()
    