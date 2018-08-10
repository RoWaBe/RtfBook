'Unfertig. Gerade erst angefangen
'FreeBasic UnZip example using LibZip static library (linked into the own executable) with Unicode characters
'Decryption example. File encryptedfile.txt in archive testfile.zip has been encrypted, all other files are not encrypted.
'see http://www.nih.at/libzip/
'see http://libzip.org/documentation/libzip.html
'see http://www.freebasic.net/wiki/wikka.php?wakka=ExtLibZip
'************************************************************************************************************************
#include Once "windows.bi"
#include Once "zip.bi"     'ZIP archives library.
#define UNICODE
'************************************************************************************************************************
Dim i As Integer,j As Integer,k As Integer
Dim swArchiveToLoad As Wstring*MAX_PATH,swFileToLoad As wString*MAX_PATH,szPassword As String
Dim pszFileNameUTF8 As Const zString Ptr,szFileNameUTF8 As zString*MAX_PATH,swFileName As wString*MAX_PATH
Dim sBuffer As String,szBuffer As zString*32768,iLenFile As Long
Dim pZipArchive As Any Ptr,pZipFile As Any Ptr,ZipStat As zip_stat_
'
swArchiveToLoad="testfile.zip"                          'ZIP archive to open
'swFileToLoad="test.txt"                                'File to extract from zip archive
swFileToLoad="File with long name and umlauts ÄÖÜß.txt" 'File with extended characters to extract from zip archive
swFileToLoad="encryptedfile.txt"                        'encrypted File to extract from zip archive
szPassword="mypass"

pZipArchive=zip_open(swArchiveToLoad,ZIP_RDONLY,@i)
If (i<>NULL) Then
  MessageBox(NULL,"Could not open "+swArchiveToLoad,"Error "+Str(i),MB_ICONSTOP) 
Else
  i=zip_get_num_entries(pZipArchive,NULL) 'ZIP_FL_UNCHANGED
  'j=zip_libzip_version()  'new in libzip v1.3.1
  j=0
  sBuffer=""
  Do 
    pszFileNameUTF8=zip_get_name(pZipArchive,j,ZIP_FL_ENC_GUESS)       'Get filename in UTF-8 format
    If pszFileNameUTF8=NULL Then Exit Do
    k=MultiByteToWideChar(CP_UTF8,0,pszFileNameUTF8,-1,swFileName,SizeOf(swFileName)) 'Convert UTF-8 filename into 16-bit Unicode
    sBuffer=sBuffer+Str(k)+" bytes: "+Left(swFileName,k)+Chr(13)         '
    j=j+1
  Loop
  '
  WideCharToMultiByte(CP_UTF8,0,ByVal Cast(LPCWSTR,SAdd(swFileToLoad)),-1,szFileNameUTF8,SizeOf(szFileNameUTF8),ByVal NULL,ByVal NULL)    'Convert filename to UTF-8
  zip_stat(pZipArchive,szFileNameUTF8,ZIP_STAT_NAME Or ZIP_STAT_SIZE Or ZIP_STAT_COMP_SIZE Or ZIP_STAT_COMP_METHOD Or ZIP_STAT_ENCRYPTION_METHOD,@ZipStat)
  sBuffer="Filename (UTF-8 format): "+*ZipStat.name+Chr(13)+_
          "File size: "+Str(ZipStat.size)+Chr(13)+_
          "Compressed size:"+Str(ZipStat.comp_size)+Chr(13)+_
          "Compression method: "+Str(ZipStat.comp_method)+Chr(13)+_
          "Encryption method: "+Str(ZipStat.encryption_method)
          'Further type members not used here: ZipStat.valid, ZipStat.index, ZipStat.mtime, ZipStat.crc, ZipStat.flags
  '
  pZipFile=zip_fopen(pZipArchive,szFileNameUTF8,0)
  If (pZipFile=NULL) Then
    MessageBox(NULL,"Could not open file"+Chr(13)+szFileNameUTF8+Chr(13)+"in archive "+swArchiveToLoad,"Error "+Str(@pZipArchive),MB_ICONSTOP) 
  Else
    iLenFile=zip_fread(pZipFile,SAdd(szBuffer),SizeOf(szBuffer))
    MessageBox(NULL,"File content:"+Chr(13,13,13)+szBuffer+Chr(13,13,13)+sBuffer,"File size of "+szFileNameUTF8+": "+Str(iLenFile)+" Bytes",NULL) 
    '
    zip_fclose(pZipFile)
  End If
  zip_close(pZipArchive)
End If  

'BESETTINGS (don't change!):
'BECURSOR=21
'BETOGGLE=
'BETARGET=1