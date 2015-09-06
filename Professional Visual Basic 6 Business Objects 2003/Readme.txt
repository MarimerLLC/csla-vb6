***** Visual Basic 6.0 Business Objects *****

This installation will install the source code for Visual Basic 6.0 Business Objects onto your machine. 

However, due to the nature of the book the code will not necessarily run straight off but rather will require some additional tinkering so that it will run on your machine(s).

1. The UI interface constructed in Chapter 7 will fail to load properly due to its use of Windows Common Controls 6.0 to provide the ListView control. As such if you load this project you will get several error messages and all instances of the ListView control will be replaced with a PictureBox control. You will have to add the Windows Common Controls 6.0 Component to the project and reinsert the LsitView controls and configure them correctly as described in Chapter 7 of the book. 

2. You will note that these files contains only the code and no compiled versions are included. Thus from Chapter 9 onwards you will need to compile this code yourself. This is mainly because you may need to adjust some of the code to run on your machine(s) - see below.

3. For Chapter 4 and from Chapter 8 onwards you may need to adjust the path where the relevant databases can be found. A copy of Video.mdb and Person.mdb are included in this installation and the code uses their location but only if you install the files to C:\Wrox\VB6 Pro Objects   For Chapter 12 you will also need to upsize the database into SQL Server yourself for the code to work.

4. From Chapter 10 onwards you will need to change the constant PERSIST_SERVER such that it specifies the name or IP address of the machine upon which you will running the program. 

5. For Chapter 12 you will need to setup the dll to be run in MTS yourself as described in the book.

6. The last few chapters using Web applications as a front end will require you to set up the relevant directory structure as specified in the book and copy the source code to that location.



If you have any problems with the code please contact Feedback@wrox.com


********      **********        **********