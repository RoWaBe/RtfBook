RTF book viewer

©Roland Walter 2018, Freeware
Updates here: www.rowalt.de/rtfbook/ 
or here: github.com/RoWaBe/RtfBook/

The program runs under Windows XP upwards and Linux Wine, Linux Crossover and ReactOS.

General information

The RTF book viewer is an eBook viewer for eBooks in *.rbk format. These are essentially ZIP-packed RTF files, nothing more. This also made a slim, fast eBook viewer possible. And I needed something like that for my own work. 
The eBook format is most similar to the ComicBook format, where JPEG files are packed into a ZIP file. For more details, see the detailed description of the format. RTF books can also be encrypted and can contain JPEG and PNG images as well as extractable attachments in addition to RTF files.

The program is largely selfexplaining. The eBooks can theoretically be up to 2GB in size, because only one chapter (i.e. an RTF file or an image file) is loaded at one time.
The program interface is automatically set to German or English. Alternatively, the language can be set in the menu "File"?"Settings". More languages would be possible theoretically , but the work has to be done by someone -)
The program is suitable also for large eBooks. I have successfully tested it with 60MB books (=ZIP files) containing more than 40000 chapters (single RTF files). The speed was surprisingly good, but please don't blame me for not wanting to kill my hard drive to determine the real limits ;-)
Encrypted RTF books: When creating an RTF book, it can be encrypted using the normal encryption function for ZIP files. It is also possible to encrypt only individual chapters. It is also possible (but can be very annoying) to use a different key for each chapter.
Displaying of RTF files is based on the library msftedit.dll, which has been part of the system since Windows XP. It is recommended to search for the latest version of this DLL. 

Terms of use and copyright

The program is freeware, therefore it may be used and distributed for free (also commercially etc.). The rule is: "Look the gifted horse into his mouth, but don't blame me if it bites and bucks. This means: You use the software under the condition that you alone have the responsibility.
The ZIP library of Dieter Baron and Thomas Klausner was used for the ZIP functionality. Great work, by the way, thank you!
