'FreeBasic include file for static libzip.a <= v1.3.0
'Include file created in handwork by R.W. march 2018
'
'Copyright (C) for libzip 1999-2015 by Dieter Baron And Thomas Klausner
'The authors can be contacted at <libzip@nih.at>zip_source_begin_write
'This file Is part of libzip, a library To manipulate ZIP archives.
'Redistribution And use in source And Binary forms, With Or without
'modification, are permitted provided that the following conditions
'are met:
'1. Redistributions of source code must retain the above copyright
'   notice, this list of conditions And the following disclaimer.
'2. Redistributions in Binary form must reproduce the above copyright
'   notice, this list of conditions And the following disclaimer in
'   the documentation And/Or other materials provided With the
'   distribution.
'3. The names of the authors may Not be used To endorse Or promote
'   products derived from this software without specific prior
'   written permission.
'
'THIS SOFTWARE Is PROVIDED BY THE AUTHORS ``As Is''AND ANY EXPRESS
'Or IMPLIED WARRANTIES, INCLUDING, BUT Not LIMITED To, THE IMPLIED
'WARRANTIES OF MERCHANTABILITY And FITNESS For A PARTICULAR PURPOSE
'ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHORS BE LIABLE For Any
'DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, Or CONSEQUENTIAL
'DAMAGES (INCLUDING, BUT Not LIMITED To, PROCUREMENT OF SUBSTITUTE
'GOODS Or SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
'INTERRUPTION) HOWEVER CAUSED And On Any THEORY OF LIABILITY, WHETHER
'IN CONTRACT, STRICT LIABILITY, Or TORT (INCLUDING NEGLIGENCE Or
'OTHERWISE) ARISING IN Any WAY Out OF THE USE OF THIS SOFTWARE, EVEN
'If ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

#ifndef _HAD_ZIP_BI
#define _HAD_ZIP_BI

#inclib "zip"
#inclib "z"                'At least when using a static libzip, zlib needs to be linked in too.

#include once "crt/stdio.bi"
#include once "crt/time.bi"
#include once "crt/stdint.bi"

Extern "c"

'<zipconf.bi>
#define LIBZIP_VERSION "1.3.0"
#define LIBZIP_VERSION_MAJOR 1
#define LIBZIP_VERSION_MINOR 3
#define LIBZIP_VERSION_MICRO 0

Type zip_int8_t As int8_t
#define ZIP_INT8_MIN INT8_MIN
#define ZIP_INT8_MAX INT8_MAX

Type zip_uint8_t As uint8_t
#define ZIP_UINT8_MAX UINT8_MAX

Type zip_int16_t As int16_t
#define ZIP_INT16_MIN INT16_MIN
#define ZIP_INT16_MAX INT16_MAX

Type zip_uint16_t As uint16_t
#define ZIP_UINT16_MAX UINT16_MAX

Type zip_int32_t As int32_t
#define ZIP_INT32_MIN INT32_MIN
#define ZIP_INT32_MAX INT32_MAX

Type zip_uint32_t As uint32_t
#define ZIP_UINT32_MAX UINT32_MAX

Type zip_int64_t As int64_t
#define ZIP_INT64_MIN INT64_MIN
#define ZIP_INT64_MAX INT64_MAX

Type zip_uint64_t As uint64_t
#define ZIP_UINT64_MAX UINT64_MAX
''</zipconf.bi>

'Flags for zip_open:
#define ZIP_CREATE           1
#define ZIP_EXCL             2
#define ZIP_CHECKCONS        4
#define ZIP_TRUNCATE         8
#define ZIP_RDONLY          16 

'Flags for zip_name_locate, zip_fopen, zip_stat, ...:
#define ZIP_FL_NOCASE         1u 'ignore case on name lookup
#define ZIP_FL_NODIR          2u 'ignore directory component
#define ZIP_FL_COMPRESSED     4u 'read compressed data
#define ZIP_FL_UNCHANGED      8u 'use original data, ignoring changes
#define ZIP_FL_RECOMPRESS    16u 'force recompression of data
#define ZIP_FL_ENCRYPTED     32u 'read encrypted data (implies ZIP_FL_COMPRESSED)
#define ZIP_FL_ENC_GUESS      0u 'guess string encoding (is default)
#define ZIP_FL_ENC_RAW       64u 'get unmodified string
#define ZIP_FL_ENC_STRICT   128u 'follow specification strictly
#define ZIP_FL_LOCAL	      256u 'in local header
#define ZIP_FL_CENTRAL	    512u 'in central directory
'                          1024u  reserved for internal use
#define ZIP_FL_ENC_UTF_8   2048u 'string is UTF-8 encoded
#define ZIP_FL_ENC_CP437   4096u 'string is CP437 encoded
#define ZIP_FL_OVERWRITE   8192u 'zip_file_add: if file with name exists, overwrite (replace) it 

'Archive global flags flags:
'#define ZIP_AFL_TORRENT    1u   'torrent zipped
#define ZIP_AFL_RDONLY      2u   'read only -- cannot be cleared

'Create a new extra field:
#define ZIP_EXTRA_FIELD_ALL	ZIP_UINT16_MAX
#define ZIP_EXTRA_FIELD_NEW	ZIP_UINT16_MAX 

'Flags for compression and encryption sources:
#define ZIP_CODEC_ENCODE    1 'compress/encrypt

'libzip error codes:
#define ZIP_ER_OK             0  'N No error
#define ZIP_ER_MULTIDISK      1  'N Multi-disk zip archives not supported
#define ZIP_ER_RENAME         2  'S Renaming temporary file failed
#define ZIP_ER_CLOSE          3  'S Closing zip archive failed
#define ZIP_ER_SEEK           4  'S Seek error
#define ZIP_ER_READ           5  'S Read error
#define ZIP_ER_WRITE          6  'S Write error
#define ZIP_ER_CRC            7  'N CRC error
#define ZIP_ER_ZIPCLOSED      8  'N Containing zip archive was closed
#define ZIP_ER_NOENT          9  'N No such file
#define ZIP_ER_EXISTS        10  'N File already exists
#define ZIP_ER_OPEN          11  'S Can't open file
#define ZIP_ER_TMPOPEN       12  'S Failure to create temporary file
#define ZIP_ER_ZLIB          13  'Z Zlib error
#define ZIP_ER_MEMORY        14  'N Malloc failure
#define ZIP_ER_CHANGED       15  'N Entry has been changed
#define ZIP_ER_COMPNOTSUPP   16  'N Compression method not supported
#define ZIP_ER_EOF           17  'N Premature EOF
#define ZIP_ER_INVAL         18  'N Invalid argument
#define ZIP_ER_NOZIP         19  'N Not a zip archive
#define ZIP_ER_INTERNAL      20  'N Internal error
#define ZIP_ER_INCONS        21  'N Zip archive inconsistent
#define ZIP_ER_REMOVE        22  'S Can't remove file
#define ZIP_ER_DELETED       23  'N Entry has been deleted
#define ZIP_ER_ENCRNOTSUPP   24  'N Encryption method not supported
#define ZIP_ER_RDONLY        25  'N Read-only archive
#define ZIP_ER_NOPASSWD      26  'N No password provided
#define ZIP_ER_WRONGPASSWD   27  'N Wrong password provided
#define ZIP_ER_OPNOTSUPP     28  'N Operation not supported
#define ZIP_ER_INUSE         29  'N Resource still in use
#define ZIP_ER_TELL          30  'S Tell error

'type of system error value:
#define ZIP_ET_NONE       0      'sys_err unused
#define ZIP_ET_SYS        1      'sys_err is errno
#define ZIP_ET_ZLIB       2      'sys_err is zlib error code

'compression methods:
#define ZIP_CM_DEFAULT        -1  'better of deflate or store
#define ZIP_CM_STORE           0  'stored (uncompressed)
#define ZIP_CM_SHRINK          1  'shrunk
#define ZIP_CM_REDUCE_1        2  'reduced with factor 1
#define ZIP_CM_REDUCE_2        3  'reduced with factor 2
#define ZIP_CM_REDUCE_3        4  'reduced with factor 3
#define ZIP_CM_REDUCE_4        5  'reduced with factor 4
#define ZIP_CM_IMPLODE         6  'imploded
'                              7   Reserved for Tokenizing compression algorithm
#define ZIP_CM_DEFLATE         8  'deflated
#define ZIP_CM_DEFLATE64       9  'deflate64
#define ZIP_CM_PKWARE_IMPLODE 10  'PKWARE imploding
'                             11   Reserved by PKWARE
#define ZIP_CM_BZIP2          12  'compressed using BZIP2 algorithm
'                             13   Reserved by PKWARE
#define ZIP_CM_LZMA           14  'LZMA (EFS)
'                            15-17 Reserved by PKWARE
#define ZIP_CM_TERSE          18  'compressed using IBM TERSE (new)
#define ZIP_CM_LZ77           19  'IBM LZ77 z Architecture (PFS)
#define ZIP_CM_WAVPACK        97  'WavPack compressed data
#define ZIP_CM_PPMD           98  'PPMd version I, Rev 1

'encryption methods:
#define ZIP_EM_NONE            0  'not encrypted
#define ZIP_EM_TRAD_PKWARE     1  'traditional PKWARE encryption
#if 0                             'Strong Encryption Header not parsed yet
#define ZIP_EM_DES        &h6601  'strong encryption: DES
#define ZIP_EM_RC2_OLD    &h6602  'strong encryption: RC2, version < 5.2
#define ZIP_EM_3DES_168   &h6603
#define ZIP_EM_3DES_112   &h6609
#define ZIP_EM_AES_128    &h660e
#define ZIP_EM_AES_192    &h660f
#define ZIP_EM_AES_256    &h6610
#define ZIP_EM_RC2        &h6702  'strong encryption: RC2, version >= 5.2
#define ZIP_EM_RC4        &h6801
#endif
#define ZIP_EM_UNKNOWN    &hffff  'unknown algorithm

#define ZIP_OPSYS_DOS	  	      &H00u
#define ZIP_OPSYS_AMIGA	 	      &H01u
#define ZIP_OPSYS_OPENVMS	      &H02u
#define ZIP_OPSYS_UNIX	  	    &H03u
#define ZIP_OPSYS_VM_CMS	      &H04u
#define ZIP_OPSYS_ATARI_ST	    &H05u
#define ZIP_OPSYS_OS_2		      &H06u
#define ZIP_OPSYS_MACINTOSH	    &H07u
#define ZIP_OPSYS_Z_SYSTEM	    &H08u
#define ZIP_OPSYS_CPM	  	      &H09u
#define ZIP_OPSYS_WINDOWS_NTFS  &H0au
#define ZIP_OPSYS_MVS	  	      &H0bu
#define ZIP_OPSYS_VSE	  	      &H0cu
#define ZIP_OPSYS_ACORN_RISC	  &H0du
#define ZIP_OPSYS_VFAT	  	    &H0eu
#define ZIP_OPSYS_ALTERNATE_MVS	&H0fu
#define ZIP_OPSYS_BEOS	  	    &H10u
#define ZIP_OPSYS_TANDEM	      &H11u
#define ZIP_OPSYS_OS_400	      &H12u
#define ZIP_OPSYS_OS_X	  	    &H13u
#define ZIP_OPSYS_DEFAULT	      ZIP_OPSYS_UNIX 

Type zip_source_cmd As Long
Enum
  ZIP_SOURCE_OPEN_            'prepare for reading
  ZIP_SOURCE_READ_            'read data
  ZIP_SOURCE_CLOSE_           'reading is done
  ZIP_SOURCE_STAT_            'get meta information
  ZIP_SOURCE_ERROR_           'get error information
  ZIP_SOURCE_FREE_            'cleanup and free resources - TODO: name conflict with zip_source_free()
  ZIP_SOURCE_SEEK_            'set position for reading
  ZIP_SOURCE_TELL_            'get read position
  ZIP_SOURCE_BEGIN_WRITE_     'prepare for writing
  ZIP_SOURCE_COMMIT_WRITE_    'writing is done
  ZIP_SOURCE_ROLLBACK_WRITE_  'discard written changes
  ZIP_SOURCE_WRITE_           'write data
  ZIP_SOURCE_SEEK_WRITE_      'set position for writing
  ZIP_SOURCE_TELL_WRITE_      'get write position
  ZIP_SOURCE_SUPPORTS         'check whether source supports command
  ZIP_SOURCE_REMOVE           'remove file 
End Enum

'#define ZIP_SOURCE_ERR_LOWER    -2

Type zip_source_cmd_t As zip_source_cmd

#define ZIP_SOURCE_MAKE_COMMAND_BITMASK(cmd) (1 Shl (cmd)) 'Alternative: #define ZIP_SOURCE_MAKE_COMMAND_BITMASK(cmd)    (((zip_int64_t)1)<<(cmd))
#define ZIP_SOURCE_SUPPORTS_READABLE (((((ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_OPEN_) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_READ_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_CLOSE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_STAT_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_ERROR_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_FREE_))
#define ZIP_SOURCE_SUPPORTS_SEEKABLE (((ZIP_SOURCE_SUPPORTS_READABLE Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_SEEK_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_TELL_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_SUPPORTS))
#define ZIP_SOURCE_SUPPORTS_WRITABLE (((((((ZIP_SOURCE_SUPPORTS_SEEKABLE Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_BEGIN_WRITE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_COMMIT_WRITE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_ROLLBACK_WRITE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_WRITE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_SEEK_WRITE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_TELL_WRITE_)) Or ZIP_SOURCE_MAKE_COMMAND_BITMASK(ZIP_SOURCE_REMOVE))

Type zip_source_args_seek  'for use by sources
	offset As zip_int64_t
	whence As Long
End Type

'ToDo:
'typedef struct zip_source_args_seek zip_source_args_seek_t;
'#define ZIP_SOURCE_GET_ARGS(type, data, len, error) ((len) < sizeof(type) ? zip_error_set((error), ZIP_ER_INVAL, 0), (type *)NULL : (type *)(data)) 

Type zip_error        'error information. use zip_error_*() to access
  zip_err As Long	    'libzip error code (ZIP_ER_*)
  sys_err As Long     'copy of errno (E*) or zlib error code
  str As zstring Ptr  'string representation or NULL
End Type

#define ZIP_STAT_NAME               &h0001u
#define ZIP_STAT_INDEX_             &h0002u  'TODO: name conflict with zip_stat_index()
#define ZIP_STAT_SIZE               &h0004u
#define ZIP_STAT_COMP_SIZE          &h0008u
#define ZIP_STAT_MTIME              &h0010u
#define ZIP_STAT_CRC                &h0020u
#define ZIP_STAT_COMP_METHOD        &h0040u
#define ZIP_STAT_ENCRYPTION_METHOD  &h0080u
#define ZIP_STAT_FLAGS              &h0100u

Type zip_stat_
	valid As zip_uint64_t              'which fields have valid values
	Name As Const zstring Ptr          'name of the file
	index As zip_uint64_t              'index within archive
	size As zip_uint64_t               'size of file (uncompressed)
	comp_size As zip_uint64_t          'size of file (compressed)
	mtime As time_t                    'modification time
	crc As zip_uint32_t                'crc of file data
	comp_method As zip_uint16_t        'compression method used
	encryption_method As zip_uint16_t  'encryption method used
	flags As zip_uint32_t              'reserved for future use
End Type

Type zip_t As zip
'Type zip_file As Any
'Type zip_source As Any
Type zip_t As zip
Type zip_error_t As zip_error
Type zip_file_t As zip_file
Type zip_source_t As zip_source
Type zip_stat_t As zip_stat
Type zip_flags_t As zip_uint32_t

Type zip_source_callback As Function (ByVal As Any Ptr,ByVal As Any Ptr,ByVal As zip_uint64_t,ByVal As zip_source_cmd) As zip_int64_t

'Declare Function zip_add(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As zip_source Ptr) As zip_int64_t                                  'deprecated in libzip 0.11, use zip_file_add()
'Declare Function zip_add_dir(ByVal As zip_t Ptr,ByVal As Const zstring Ptr) As zip_int64_t                                                      'deprecated in libzip 0.11, use zip_dir_add() 
Declare Function zip_archive_set_tempdir(ByVal As zip_t Ptr,ByVal As Const zstring Ptr) As Integer
Declare Function zip_close(ByVal As zip_t Ptr) As Integer
Declare Function zip_delete(ByVal As zip_t Ptr,ByVal As zip_uint64_t) As Integer
Declare Function zip_dir_add(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As Integer) As zip_int64_t
Declare Sub zip_discard(ByVal As zip_t Ptr)
Declare Sub zip_error_clear(ByVal As zip_t Ptr)
'Declare Sub zip_error_get(ByVal As zip_t Ptr,ByVal As Integer Ptr,ByVal As Integer Ptr)                                                         'deprecated in libzip 1.0, use zip_file_get_error(), zip_error_code_zip(), / zip_error_code_system()
Declare Function zip_error_code_system(ByVal As Const zip_error_t Ptr) As Integer 'declare function zip_error_code_system(byval as const zip_error_t ptr) as long
Declare Function zip_error_code_zip(ByVal As Const zip_error_t Ptr) As Integer
Declare Sub zip_error_fini(ByVal As zip_error_t Ptr)
'Declare Function zip_error_get_sys_type(ByVal As Integer) As Integer                                                                          'deprecated in libzip 1.0, use zip_error_system_type()
Declare Sub zip_error_init(ByVal As zip_error_t Ptr)
Declare Sub zip_error_init_with_code(ByVal As zip_error_t Ptr,ByVal As Integer)
Declare Sub zip_error_set(ByVal As zip_error_t Ptr,ByVal As Integer,ByVal As Integer)
Declare Function zip_error_strerror(ByVal As zip_error_t Ptr) As Const zstring Ptr
'Declare Function zip_error_to_str(ByVal As UByte Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As Integer) As Integer                      'deprecated in libzip 1.0, use zip_error_init_with_code() and zip_error_strerror()
Declare Function zip_error_system_type(ByVal As Const zip_error_t Ptr) As Integer
Declare Function zip_file_get_error(ByVal As zip_file_t Ptr) As zip_error_t Ptr
Declare Function zip_get_error(ByVal As zip_t Ptr) As zip_error_t Ptr
Declare Function zip_source_error(ByVal As zip_source_t Ptr) As zip_error_t Ptr
Declare Function zip_error_to_data(ByVal As Const zip_error_t Ptr,ByVal As Any Ptr,ByVal As zip_uint64_t) As zip_int64_t
Declare Function zip_fclose(ByVal As zip_file_t Ptr) As Integer
Declare Function zip_fdopen(ByVal As Integer,ByVal As Integer,ByVal As Integer Ptr) As zip_t Ptr
Declare Function zip_file_add(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As zip_source_t Ptr,ByVal As Integer) As zip_int64_t
Declare Sub zip_file_error_clear(ByVal As zip_file_t Ptr)
'Declare Sub zip_file_error_get(ByVal As zip_file_t Ptr,ByVal As Integer Ptr,ByVal As Integer Ptr)                                               'deprecated. Use zip_error_code_system(3), zip_error_code_zip(3), zip_file_get_error(3), and zip_get_error(3)
Declare Function zip_file_extra_field_delete(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint16_t,ByVal As Integer) As Integer
Declare Function zip_file_extra_field_delete_by_id(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint16_t,ByVal As zip_uint16_t,ByVal As Integer) As Integer
Declare Function zip_file_extra_field_get(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint16_t,BayVal As zip_uint16_t Ptr,ByVal As zip_uint16_t Ptr,ByVal As Integer) As zip_uint8_t Ptr
Declare Function zip_file_extra_field_get_by_id(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint16_t,ByVal As zip_uint16_t,ByVal As zip_uint16_t Ptr,ByVal As Integer) As zip_uint8_t Ptr
Declare Function zip_file_extra_field_set(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint16_t,ByVal As zip_uint16_t,ByVal As Const  zip_uint8_t Ptr,ByVal As zip_uint16_t,ByVal As Integer) As Integer
Declare Function zip_file_extra_fields_count(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer) As zip_int16_t
Declare Function zip_file_extra_fields_count_by_id(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint16_t,ByVal As Integer) As zip_int16_t
Declare Function zip_file_get_comment(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_uint32_t Ptr,ByVal As Integer) As Const zstring Ptr
Declare Function zip_file_get_external_attributes(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As zip_uint8_t Ptr,ByVal As zip_uint32_t Ptr) As Integer
Declare Function zip_file_rename(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Const zstring Ptr,ByVal As Integer) As Integer
Declare Function zip_file_replace(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_source_t Ptr,ByVal As Integer) As Integer
Declare Function zip_file_set_comment(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Const zstring Ptr,ByVal As zip_uint16_t,ByVal As Integer) As Integer
Declare Function zip_file_set_external_attributes(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As zip_uint8_t,ByVal As zip_uint32_t) As Integer
Declare Function zip_file_set_mtime(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As time_t,ByVal As Integer) As Integer
Declare Function zip_file_strerror(ByVal As zip_file_t Ptr) As Const zstring Ptr
Declare Function zip_fopen(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As Integer) As zip_file_t Ptr 
Declare Function zip_fopen_encrypted(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As Integer,ByVal As Const zstring Ptr) As zip_file_t Ptr 
Declare Function zip_fopen_index(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer) As zip_file_t Ptr
Declare Function zip_fopen_index_encrypted(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As Const zstring Ptr) As zip_file_t Ptr 
Declare Function zip_fread(ByVal As zip_file_t Ptr,ByVal As Any Ptr,ByVal As zip_uint64_t) As zip_int64_t
Declare Function zip_get_archive_comment(ByVal As zip_t Ptr,ByVal As Integer Ptr,ByVal As Integer) As Const zstring Ptr
'Declare Function zip_get_file_comment(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer Ptr,ByVal As Integer) As Const zstring Ptr      'deprecated in libzip 0.11, use zip_file_get_comment()
'Declare Function zip_get_file_extra(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer Ptr,ByVal As Integer) As Const zstring Ptr        'Old version of libzip.bi, doesn't exist (?)
Declare Function zip_get_archive_flag(ByVal As zip_t Ptr,ByVal As Integer,ByVal As Integer) As Integer
Declare Function zip_get_name(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer) As Const zstring Ptr
Declare Function zip_get_num_entries(ByVal As zip_t Ptr,ByVal As Integer) As zip_int64_t
'Declare Function zip_get_num_files(ByVal As zip_t Ptr) As Integer                                                                               'deprecated in libzip 0.11, use zip_get_num_entries(instead)
Declare Function zip_name_locate(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As Integer) As zip_int64_t                                  'New version ToDo: Check!
'Declare Function zip_name_locate(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As Integer) As Integer                                     'Old version
Declare Function zip_open(ByVal As Const zstring Ptr,ByVal As Integer,ByVal As Integer Ptr) As zip_t Ptr
Declare Function zip_open_from_source(ByVal As zip_source_t Ptr,ByVal As Integer,ByVal As zip_error_t Ptr) As zip_t Ptr
'Declare Function zip_rename(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Const zstring Ptr) As Integer                                     'deprecated in libzip 0.11, use zip_file_rename()
'Declare Function zip_replace(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_source Ptr) As Integer                                       'deprecated, use zip_file_replace(3) with an empty flags argument
Declare Function zip_set_archive_comment(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As zip_uint16_t) As Integer                         'New version ToDo: Check!
'Declare Function zip_set_archive_comment(ByVal As zip_t Ptr,ByVal As Const UByte Ptr,ByVal As Integer) As Integer                               'Old version
Declare Function zip_set_archive_flag(ByVal As zip_t Ptr,ByVal As Integer,ByVal As Integer) As Integer
Declare Function zip_set_default_password(ByVal As zip_t Ptr,ByVal As Const zstring Ptr) As Integer
Declare Function zip_set_file_compression(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_int32_t,ByVal As zip_uint32_t) As Integer
'Declare Function zip_set_file_comment(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Const UByte Ptr,ByVal As Integer) As Integer            'deprecated in libzip 0.11, use zip_file_set_comment()
'Declare Function zip_set_file_extra(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Const UByte Ptr,ByVal As Integer) As Integer              'Old version of libzip.bi, doesn't exist (?)
Declare Function zip_source_begin_write(ByVal As zip_source_t Ptr) As Long
Declare Function zip_source_buffer(ByVal As zip_t Ptr,ByVal As Const Any Ptr,ByVal As zip_uint64_t,ByVal As Integer) As zip_source_t Ptr
Declare Sub zip_source_buffer_create(ByVal As Const Any Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As zip_error_t Ptr)
Declare Function zip_source_close(ByVal As zip_source_t Ptr) As Integer
Declare Function zip_source_commit_write(ByVal As zip_source_t Ptr) As Integer
Declare Sub zip_source_file(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As zip_uint64_t,ByVal As zip_int64_t)                      'New version ToDo: Check!
'Declare Function zip_source_file(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As zip_uint64_t,ByVal As zip_int64_t) As zip_source Ptr    'Old version
Declare Sub zip_source_file_create(ByVal As Const zstring Ptr,ByVal As zip_uint64_t,ByVal As zip_int64_t,ByVal As zip_error_t Ptr)
Declare Sub zip_source_filep(ByVal As zip_t Ptr,ByVal As Any Ptr,ByVal As zip_uint64_t,ByVal As zip_int64_t)                               'New version ToDo: Check!
'Declare Function zip_source_filep(ByVal As zip_t Ptr,ByVal As FILE Ptr,ByVal As zip_uint64_t,ByVal As zip_int64_t) As zip_source Ptr            'Old version
Declare Sub zip_source_filep_create(ByVal As Any Ptr,ByVal As zip_uint64_t,ByVal As zip_int64_t,ByVal As zip_error_t Ptr)
Declare Sub zip_source_free(ByVal As zip_source_t Ptr)
Declare Sub zip_source_function(ByVal As zip_t Ptr,ByVal As zip_source_callback,ByVal As Any Ptr)                                          'New version ToDo: Check!
'Declare Function zip_source_function(ByVal As zip_t Ptr,ByVal As zip_source_callback,ByVal As Any Ptr) As zip_source Ptr                        'Old version
Declare Function zip_source_function_create(ByVal As zip_source_callback,ByVal As Any Ptr,ByVal As zip_error_t Ptr) As zip_source_t Ptr
Declare Function zip_source_is_deleted(ByVal As zip_source_t Ptr) As Integer
Declare Sub zip_source_keep(ByVal As zip_source_t Ptr)
'Declare Function zip_source_make_command_bitmap(zip_source_cmd_t,...) As zip_int64_t  'ToDo - c: ZIP_EXTERN zip_int64_t zip_source_make_command_bitmap(zip_source_cmd_t,...); 
Declare Function zip_source_open(ByVal As zip_source_t Ptr) As Integer
Declare Function zip_source_read(ByVal As zip_source_t Ptr,ByVal As Any Ptr,ByVal As zip_uint64_t) As zip_int64_t
Declare Sub zip_source_rollback_write(ByVal As zip_source_t Ptr)
Declare Function zip_source_seek(ByVal As zip_source_t Ptr,ByVal As zip_int64_t,ByVal As Integer) As Integer
Declare Function zip_source_seek_compute_offset(ByVal As zip_uint64_t,ByVal As zip_uint64_t,ByVal As Any Ptr,ByVal As zip_uint64_t,ByVal As zip_error_t Ptr) As zip_int64_t
Declare Function zip_source_seek_write(ByVal As zip_source_t Ptr,ByVal As zip_int64_t,ByVal As Integer) As Integer
Declare Function zip_source_stat(ByVal As zip_source_t Ptr,ByVal As zip_stat_ Ptr) As Integer
Declare Function zip_source_tell(ByVal As zip_source_t Ptr) As zip_int64_t
Declare Function zip_source_tell_write(ByVal As zip_source_t Ptr) As zip_int64_t
Declare Function zip_source_write(ByVal As zip_source_t Ptr,ByVal As Const Any Ptr,ByVal As zip_uint64_t) As zip_int64_t
Declare Function zip_source_zip(ByVal As zip_t Ptr,ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As zip_flags_t,ByVal As zip_uint64_t,ByVal As zip_int64_t) As zip_source_t Ptr
'Declare Function zip_source_zip(ByVal As zip_t Ptr,ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As zip_uint64_t,ByVal As zip_int64_t) As zip_source Ptr  'Old version
Declare Function zip_stat(ByVal As zip_t Ptr,ByVal As Const zstring Ptr,ByVal As Integer,ByVal As zip_stat_ Ptr) As Integer
Declare Function zip_stat_index(ByVal As zip_t Ptr,ByVal As zip_uint64_t,ByVal As Integer,ByVal As zip_stat_ Ptr) As Integer
Declare Sub zip_stat_init(ByVal As zip_stat_ Ptr)
Declare Function zip_strerror(ByVal As zip_t Ptr) As Const zstring Ptr
Declare Function zip_unchange(ByVal As zip_t Ptr,ByVal As zip_uint64_t) As Integer
Declare Function zip_unchange_all(ByVal As zip_t Ptr) As Integer
Declare Function zip_unchange_archive(ByVal As zip_t Ptr) As Integer

#ifdef __FB_WIN32__
Declare Function zip_source_win32a(ByVal As zip_t Ptr, ByVal As Const zstring Ptr, ByVal As zip_uint64_t, ByVal As zip_int64_t) As zip_source_t Ptr
Declare Function zip_source_win32a_create(ByVal As Const zstring Ptr, ByVal As zip_uint64_t, ByVal As zip_int64_t, ByVal As zip_error_t Ptr) As zip_source_t Ptr
Declare Function zip_source_win32handle(ByVal As zip_t Ptr, ByVal As Any Ptr, ByVal As zip_uint64_t, ByVal As zip_int64_t) As zip_source_t Ptr
Declare Function zip_source_win32handle_create(ByVal As Any Ptr, ByVal As zip_uint64_t, ByVal As zip_int64_t, ByVal As zip_error_t Ptr) As zip_source_t Ptr
Declare Function zip_source_win32w(ByVal As zip_t Ptr, ByVal As Const wstring Ptr, ByVal As zip_uint64_t, ByVal As zip_int64_t) As zip_source_t Ptr
Declare Function zip_source_win32w_create(ByVal As Const wstring Ptr, ByVal As zip_uint64_t, ByVal As zip_int64_t, ByVal As zip_error_t Ptr) As zip_source_t Ptr
#endif 

End Extern

#endif

'BESETTINGS (don't change!):
'BECURSOR=6B
'BETOGGLE=
'BETARGET=0