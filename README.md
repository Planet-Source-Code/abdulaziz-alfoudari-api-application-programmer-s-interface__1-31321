<div align="center">

## API \- Application Programmer's Interface


</div>

### Description

Visit us at: http://www.vbparadise.com. As a programmer you might never have to use a Windows API.... Nahhh! The fact is that serious programmers use API all the time. When they need to do something that VB cannot handle, they turn to the Windows API! The API are procedures that exist in files on your PC which you can call from within your VB program - and there are thousands of them!. Written by Microsoft, debugged by tens of thousands of users, and available for free with Windows - the API are one of the very best tools you have available to add power to your VB application.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Abdulaziz Alfoudari](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/abdulaziz-alfoudari.md)
**Level**          |Advanced
**User Rating**    |4.4 (114 globes from 26 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/abdulaziz-alfoudari-api-application-programmer-s-interface__1-31321/archive/master.zip)





### Source Code

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="712" id="AutoNumber1">
 <tr>
 <td width="712">
 <h4 align="center"><font face="Verdana" size="5">Visit Us At:
 <span style="font-weight: 400"><a href="http://www.vbparadise.com">
 http://www.vbparadise.com</a></span></font></h4>
 <h4></h4>
 <h4><font face="Verdana" size="2">API - Application Programmer's Interface</font></h4>
 <p><font face="Verdana" size="2">When Microsoft wrote Windows they put a huge
 amount of code into procedure libraries which programmers can access. No matter
 which language you're using (VB, C++, ...) you can use the Windows API to
 greatly expand the power of your application. </font><p>
 <font face="Verdana" size="2">There are a lot of Windows programs whose code is
 spread across multiple files. The .EXE file does not always contain all of the
 code that a programmer might use. For example, by creating his own library of
 procedures (usually in the form of a file with a .DLL extension) a programmer
 can allow more than one of his applications to access the same code. </font><p>
 <font face="Verdana" size="2">Microsoft does a similar thing with Windows.
 There are many files which have code that you can access, but the three most
 often named files are: </font>
 <ul>
        <li><font face="Verdana" size="2"><b>user32.dll</b> - controls
        the visible objects that you see on the screen </font></li>
        <li><font face="Verdana" size="2"><b>gdi32</b> - home of most
        graphics oriented API </font></li>
        <li><font face="Verdana" size="2"><b>kernel32.dll</b> - provides
        access to low level operating system features </font></li>
 </ul>
 <p><font face="Verdana" size="2">Later, I'll bring up some of the other files
 whose procedures you might want access. However, there are some key issues
 which you should note before making a decision to use an API call. </font>
 <ul>
        <li><font face="Verdana" size="2"><b>Version Compatibility</b><br>
        Microsoft has long been known to update it's files without much
        fanfare - read that as without telling anyone about it until it's
        already happened! And often, the updated code may not perform
        exactly as did the older version. Users often find this out by
        seeing unexected errors or by having their system crash and/or
        lock up on them! In VB5 there were a huge number of programmer's
        who got bit by this problem. </font><p>
        <font face="Verdana" size="2">If you stick with the basic 3 OS
        files listed above, you won't see to much of this. But the
        further away you go from the main 3 files, the more likely you
        are to get into code which hasn't seen the testing and
        improvement cycle that the main Windows OS files have gone
        through. </font></li>
        <li><font face="Verdana" size="2"><b>File Size</b><br>
        One of the <b>very major</b> downsides to the concept of API is
        that all of this great code lives in some very big files! Worse
        yet, sometimes the API you want are spread over multiple files
        and you may be using only one or two procedures from enormous
        files which have hundreds of procedures in them. Where this
        becomes a problem is in a)load time - where it can takes several
        seconds to load the procedure libraries, and b) - where you want
        to distribute your application and in order to make sure that all
        of the procedure libraries are on your user's machine, you have
        to put all of them into the distribute files. This can add many
        megabytes of files to your distribution applications. It is a
        major problem for distribution of software over the net, where 5
        minutes per megabyte can deter a usage from trying out an
        application just because he doesn't want to wait for the
        download! </font></li>
        <li><font face="Verdana" size="2"><b>Documentation</b><br>
        Finding the documentation of what procedures are in a library and
        how to use them can be very difficult. On my PC I have 3,380
        files with a .DLL extension with a total size of 539MB. That's a
        lot of code! Unfortunately I can count on one hand the pages of
        documentation that I have to tell me what that code is or does!
        You'll learn <b>how</b> to use DLLs in this tutorial, but without
        the documentation from the creator of the DLLs you cannot use
        them successfully. </font></li>
 </ul>
 <p><font face="Verdana" size="2">Despite these problems, the powerful magic of
 the API is that they are code which you don't have to write. If you've read my
 Beginner's section you know that I am a big fan of using 3rd party software to
 lighten my own programming load. As with 3rd party controls, the API provide
 procedures which someone else wrote, debugged, and made availble for you to
 benefit from. In the Windows DLLs files, there are literally thousands of
 procedures. The key to API programming is learning which of these procedures
 are useful and which ones you are unlikely to ever need! This tutorial tries to
 address just that problem. </font><p><font face="Verdana" size="2"><b>Getting
 Started</b> It's actually simpler than you might imagine. By now, you've
 already written procedures for your own VB programs. Using procedures from
 other files is almost exactly the same as using procedures from within your own
 program. </font><p><font face="Verdana" size="2">The one big difference is that
 you must tell your application which file the procedure is contained in. To do
 so, you must put 1 line of code into your VB program. And, you have to do this
 for every external procedure that you plan to use. I suppose it would be nice
 for VB to have the ability to find the procedures for you - but you can see
 that searching through the 3,380 procedures on my PC might slow my applications
 down a lot! </font><p><font face="Verdana" size="2">Ok, let's get to an
 example. Telling VB about the procedure you want to use is known as &quot;declaring&quot;
 the procedure, and (no surprise) it uses a statement which starts with the word
 declare. Here's what a declaration looks like: </font><p>
 <pre><font face="Verdana">Declare Function ExitWindowsEx Lib &quot;user32&quot; (ByVal uFlags as Long, ByVal dwReserved as Long) as Long
</font></pre>
 <p><font face="Verdana" size="2">Let's take about the parts of the declaration
 statement: </font>
 <ul>
        <li><font face="Verdana" size="2"><b>&quot;Declare&quot;</b><br>
        This is reserved word that VB uses to begin a declaration. There
        is no alternative - you have to use it. </font></li>
        <li><font face="Verdana" size="2"><b>&quot;Function&quot;</b><br>
        Also a reserved word, but in this case it distinguishes between a
        SUB procedured and a FUNCTION procedure. The API use Function
        procedures so that they can return a value to indicate the
        results of the action. Although you can discard the returned
        value, it's also possible to check the return value to determine
        that the action was successfully completed. completed.
        alternative - you have to use it. </font></li>
        <li><font face="Verdana" size="2"><b>&quot;ExitWindowsEx&quot;</b><br>
        Inside each DLL is a list of the procedures that it contains.
        Normally, in a VB declaration statement you simply type in the
        name of the procedure just as it is named in the DLL. Sometimes,
        the DLL true name of the procedure may be a name that is illegal
        in VB. For those cases, VB allows you to put in the text string
        &quot;Alias NewProcedurename&quot; right behind the filename. In this
        example, VB would make a call to the procedure by using the name
        &quot;NewProcedureName&quot;. </font></li>
        <li><font face="Verdana" size="2"><b>&quot;Lib 'user32'&quot;</b><br>
        Here's where you tell VB which file the procedure is in. Normally
        you would put &quot;user32.dll&quot;, showing the extension of the
        procedure library. For the special case of the three Windows
        system DLLs listed above, the API will recognize the files when
        simply named &quot;user32&quot;, &quot;kernel32&quot;, and &quot;gdi32&quot; - without the DLL
        extensions shown. In most other cases you must give the complete
        file name. Unless the file is in the system PATH, you must also
        give the complete path to the file. </font></li>
        <li><font face="Verdana" size="2"><b>&quot;(ByVal uFlags as Long ...)&quot;</b><br>
        Exactly like your own procedures, Windows API functions can have
        a list of arguments. However, while your VB procedures often use
        arguments passed by reference (i.e., their values can be
        changed), most Windows API require that the arguments be passed
        by value (i.e, a copy of the argument is passed to the DLL and
        the originial variable cannot be changed). </font><p>
        <font face="Verdana" size="2">Also, you'll note that a constant
        or variable is normally used as the argument for an API call.
        It's technically acceptable to simply use a number for an
        argument but it is common practice among experienced programmers
        to create constants (or variables) whose name is easy to remember
        and then to use those in the argument list. When you're reading
        or debugging your code later, the use of these easy to read
        constant/variable names makes it much easier to figure out what
        went wrong! </font></li>
        <li><font face="Verdana" size="2"><b>&quot;as Long&quot;</b><br>
        This is exactly like the code you use to create your own
        functions. Windows API are functions which return values and you
        must define what type of variable is returned. </font></li>
 </ul>
 <p><font face="Verdana" size="2">While I make it sound simple (and it is),
 there are still issues which ought to concern you when using the Windows API.
 Because the API code executes outside the VB program itself, your own program
 is susceptable to error in the external procedure. If the external procedure
 crashes, then your own program will crash as well. It is very common for an API
 problem to freeze your system and force a reboot. </font><p>
 <font face="Verdana" size="2">The biggest issue that VB programmers would see
 in this case is that any unsaved code <b>will be lost!</b>. So remember the
 rule when using API - save often! </font><p><font face="Verdana" size="2">
 Because many of the DLLs you will use have been debugged extensively you
 probably won't see many cases where the DLL crashes because of programming bug.
 Far more frequently VB programmers will see a crash because they passed
 arguments to the procedure which the procedure could not handle! For example,
 passing a string when an integer was needed will likely crash the system. The
 DLLs don't include extensive protection in order to keep their own code size
 small and fast. </font><p><font face="Verdana" size="2">It is simple to say
 that if you pass the correct type of argument, that you won't see API crashes.
 However, the documentation is not always clear exactly what argument type is
 needed, plus when writing code it is all too common to simply make a mistake!
 </font><p><font face="Verdana" size="2">Finally, it is the case that most of
 the DLLs you'll want to use were written in C++. The significance of this is
 that the data types in C++ do not map cleanly into the data types that are used
 in Visual Basic. Here are some of the issues which you need to be aware of:
 </font><p>
 <ul>
        <li><font face="Verdana" size="2"><b>Issue1</b> </font></li>
        <li><font face="Verdana" size="2"><b>Issue2</b> </font></li>
 </ul>
 <p><font face="Verdana" size="2">Okay, stay with me just a bit longer and we'll
 get into the actual use of some API. But first, here is a list of other DLLs
 which have procedures that could be of use to you. These DLLs will show up
 later in this tutorial when we get to the API which I recommend that you
 consider for use in your own applications. </font><p>
 <ul>
        <li><font face="Verdana" size="2">Advapi32.dll - Advanced API
        services including many security and Registry calls </font></li>
        <li><font face="Verdana" size="2">Comdlg32.dll - Common dialog
        API library </font></li>
        <li><font face="Verdana" size="2">Lz32.dll - 32-bit compression
        routines </font></li>
        <li><font face="Verdana" size="2">Mpr.dll - Multiple Provider
        Router library </font></li>
        <li><font face="Verdana" size="2">Netapi32.dll - 32-bit Network
        API library </font></li>
        <li><font face="Verdana" size="2">Shell32.dll - 32-bit Shell API
        library </font></li>
        <li><font face="Verdana" size="2">Version.dll - Version library
        </font></li>
        <li><font face="Verdana" size="2">Winmm.dll - Windows multimedia
        library </font></li>
        <li><font face="Verdana" size="2">Winspool.drv - Print spoolder
        interface </font></li>
 </ul>
 <p><font face="Verdana" size="2">Often, the documentation that you might find
 for an API will be written for a C++ programmer. Here's a short table which
 helps you translate the C++ variable type declaration to its equivalent in
 Visual Basic: </font><p><table cellSpacing="0" cellPadding="0">
 <tr>
 <td><font face="Verdana" size="2">ATOM </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Integer </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">BOOL </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">BYTE </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Byte </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">CHAR </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Byte </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">COLORREF </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">DWORD </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">HWND </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">HDC </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">HMENU </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">INT </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">UINT </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LONG </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPARAM </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPDWORD </font></td>
 <td><font face="Verdana" size="2">variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPINT </font></td>
 <td><font face="Verdana" size="2">variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPUINT </font></td>
 <td><font face="Verdana" size="2">variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPRECT </font></td>
 <td><font face="Verdana" size="2">variable as Type any variable of that User
 Type </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPSTR </font></td>
 <td><font face="Verdana" size="2">ByVal variable as String </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPCSTR </font></td>
 <td><font face="Verdana" size="2">ByVal variable as String </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPVOID </font></td>
 <td><font face="Verdana" size="2">variable As Any use ByVal when passing a
 string </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPWORD </font></td>
 <td><font face="Verdana" size="2">variable as Integer </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">LPRESULT </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">NULL </font></td>
 <td><font face="Verdana" size="2">ByVal Nothing or ByVal 0&amp; or vbNullString
 </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">SHORT </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Integer </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">VOID </font></td>
 <td><font face="Verdana" size="2">Sub Procecure not applicable </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">WORD </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Integer </font></td>
 </tr>
 <tr>
 <td><font face="Verdana" size="2">WPARAM </font></td>
 <td><font face="Verdana" size="2">ByVal variable as Long </font></td>
 </tr>
 </table>
 <p><font face="Verdana" size="2">We're not quite ready to get into using the
 API. Here is a scattering of issues/comments about using API which you will
 want to be aware of: </font><p>
 <ul>
        <li><font face="Verdana" size="2"><b>Declare</b> </font>
        <ul>
               <li><font face="Verdana" size="2">DECLARE in
               standard module are PUBLIC by default and be used
               anywhere in your app </font></li>
               <li><font face="Verdana" size="2">DECLARE in any
               other module are PRIVATE to that module and MUST BE
               marked PRIVATE </font></li>
               <li><font face="Verdana" size="2">Procedure names
               are CASE-SENSITIVE </font></li>
               <li><font face="Verdana" size="2">You cannot
               Declare a 16-bit API function in VB6 </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>ALIAS</b> </font>
        <ul>
               <li><font face="Verdana" size="2">Is the &quot;real&quot;
               name of the procedure as found in the DLL </font>
               </li>
               <li><font face="Verdana" size="2">If the API uses
               string, you MUST use ALIAS with &quot;A&quot; to specify the
               correct character set (A=ANSI W=UNICODE) </font>
               </li>
               <li><font face="Verdana" size="2">WinNT supports W,
               but Win95/Win98 do not </font></li>
               <li><font face="Verdana" size="2">Some DLLs have
               illegal VB name, so you must use ALIAS to rename
               the procedure </font></li>
               <li><font face="Verdana" size="2">Can also be the
               ordinal number of the procedure </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Variable Type</b> </font>
        <ul>
               <li><font face="Verdana" size="2">Very few DLLs
               recognize VARIANT </font></li>
               <li><font face="Verdana" size="2">ByRef is VB
               default </font></li>
               <li><font face="Verdana" size="2">Most DLLs expect
               ByVal </font></li>
               <li><font face="Verdana" size="2">In C
               documentation, C passes all arguments except arrays
               by value </font></li>
               <li><font face="Verdana" size="2">AS ANY can be
               used but it turns off all type checking </font>
               </li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Strings</b> </font>
        <ul>
               <li><font face="Verdana" size="2">API generally
               require fixed length strings </font></li>
               <li><font face="Verdana" size="2">Pass string ByVal
               means passing pointer to first data byte in the
               string </font></li>
               <li><font face="Verdana" size="2">Pass string ByRef
               means passing memory address to another memory
               addresss which refers to first data byte in the
               string </font></li>
               <li><font face="Verdana" size="2">Most DLLs expect
               LPSTR (ASCIIZ) strings (end in a null character),
               which point to the first data byte </font></li>
               <li><font face="Verdana" size="2">VB Strings should
               be passed ByVal (in general) </font></li>
               <li><font face="Verdana" size="2">VB uses BSTR
               strings (header + data bytes) - BSTR is passed as a
               pointer to the header </font></li>
               <li><font face="Verdana" size="2">DLL can modify
               data in a string variable that it receives as an
               argument - WARNING: if returned value is longer
               than passed value, system error occurs! </font>
               </li>
               <li><font face="Verdana" size="2">Generally, API do
               not expect string buffers longer than 255
               characters </font></li>
               <li><font face="Verdana" size="2">C &amp; VB both treat
               a string array as an array of pointers to string
               data </font></li>
               <li><font face="Verdana" size="2">Most API require
               you to pass the length of the string and to fill
               the string wih spaces </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Arrays</b> </font>
        <ul>
               <li><font face="Verdana" size="2">Pass entire array
               by passing the first element of the array ByRef
               </font></li>
               <li><font face="Verdana" size="2">Pass individual
               elements of array just like any other variable
               </font></li>
               <li><font face="Verdana" size="2">If pass pass
               binary data to DLL, use array of Byte characters
               </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Callback Function</b> </font>
        <ul>
               <li><font face="Verdana" size="2">Use AddressOf to
               pass a user-defined function that the DLL procedure
               can use </font></li>
               <li><font face="Verdana" size="2">Must have
               specific set of arguments, AS DEFINED by the API
               procedure </font></li>
               <li><font face="Verdana" size="2">Procedure MUST be
               in a .BAS module </font></li>
               <li><font face="Verdana" size="2">Passed procedure
               must be As Any or As Long </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Passing a null value</b>
        </font>
        <ul>
               <li><font face="Verdana" size="2">To pass a null
               value - zero-length string (&quot;&quot;) will not work
               </font></li>
               <li><font face="Verdana" size="2">To pass a null
               value - use vbNullString </font></li>
               <li><font face="Verdana" size="2">To pass a null
               value - change Type to Long and then use 0&amp; </font>
               </li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Window Handle</b> </font>
        <ul>
               <li><font face="Verdana" size="2">A handle is
               simply a number assigned by Windows to each window
               </font></li>
               <li><font face="Verdana" size="2">In VB, the handle
               is the same as the property hWnd </font></li>
               <li><font face="Verdana" size="2">Handles are
               always Long variable types </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Callbacks</b> </font>
        <ul>
               <li><font face="Verdana" size="2">Some API can run
               one of you own VB functions. Your VB function is
               called a &quot;Callback&quot; </font></li>
               <li><font face="Verdana" size="2">VB supports
               callbacks with a function &quot;AddressOf&quot;, which give
               the API the location of the function to execute
               </font></li>
               <li><font face="Verdana" size="2">Callback
               functions must be in a module. They cannot be in a
               form. </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Subclassing</b> </font>
        <ul>
               <li><font face="Verdana" size="2">All windows work
               by processing messages from the Windows operating
               system </font></li>
               <li><font face="Verdana" size="2">You can change
               how a window responds to a message by intercepting
               the message </font></li>
               <li><font face="Verdana" size="2">To intercept a
               message, use the API SetWindowsLong </font></li>
        </ul>
        </li>
        <li><font face="Verdana" size="2"><b>Miscellaneous</b> </font>
        <ul>
               <li><font face="Verdana" size="2">Control
               properties MUST be passed by value (use
               intermediate value to pass ByRef) </font></li>
               <li><font face="Verdana" size="2">Handles - always
               declare as ByVal Long </font></li>
               <li><font face="Verdana" size="2">Variant - to pass
               Variant to argument that is not a Variant type,
               pass the Variant data ByVal </font></li>
               <li><font face="Verdana" size="2">UDT - cannot be
               passed except as ByRef </font></li>
        </ul>
        </li>
 </ul>
 <p><font face="Verdana" size="2"><b>Which API Should I Use?</b><br>
 Finally we get to the good part. First the bad news, then the good news. In
 this section I do not provide code that you can simply copy into your own
 applications. The good news is that I provide a list of features that you might
 want to incorporate into your own application and then tell you which of the
 API to use. For the purposes of this relatively short tutorial, the best I can
 do is to point you off in the right direction! </font><p>
 <font face="Verdana" size="2">In case you don't know, VB6 comes with a tool to
 help you use API in your own applications. The <b>API Viewer</b> is installed
 automatically with VB, and to use it go to the Start/Programs/VB/Tools menu and
 select &quot;API Viewer&quot;. The viewer actions much like my own <b>VB Information
 Center Code Librarian</b> in that you can browse through the various API,
 select one for copying to the clipboard, and then paste the declaration into
 your own application's code window. You'll definitely want to try this out. The
 data file that comes with the viewer if very extensive, listing 1550 API
 Declarations. </font><p><font face="Verdana" size="2">In my case I use API
 regularly, but I've never come close to using 1550 API. At best, I barely have
 broken the 100 mark. It seems that for the most part I can get VB to do
 whatever task I want without resorting to the API. However, in some cases you
 just can do any better than a few lines of API code to get the job done! So,
 here's my own list of useful tasks and the API needed to perform them: </font>
 <p>&nbsp;<p>
 <table cellSpacing="0" cellPadding="0" width="712" style="border-collapse: collapse" bordercolor="#111111">
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Play
 sound</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function sndPlaySound Lib &quot;winmm.dll&quot; Alias &quot;sndPlaySoundA&quot; (ByVal
 lpszSoundName as string, ByVal uFlags as Long) as Long <br>
 Result = sndPlaySound (SoundFile, 1) </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>
 SubClassing</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function CallWindowProc Lib &quot;user32&quot; Alias &quot;CallWindowProcA&quot; (ByVal
 lpPrevWndFunc as Long, ByVal hwnd as Long, byval msg as long, byval wParam as
 long, byval lParam as Long ) as long <br>
 Declare Function SetWindowLong Lib &quot;user32&quot; Alias &quot;SetWindowLongA&quot; (ByVal hwnd
 As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Run
 associated EXE</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function ShellExecute Lib &quot;shell32.dll&quot; Alias &quot;ShellExecuteA&quot; (ByVal hwnd As
 Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters
 As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long </font>
 </td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>List
 window handles</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function EnumWindows Lib &quot;user32&quot; (ByVal lpEnumFunc As Long, ByVal lParam As
 Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Find
 prior instance of EXE</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function FindWindow Lib &quot;user32&quot; Alias &quot;FindWindowA&quot; (ByVal lpClassName As
 String, ByVal lpWindowName As String) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Draw
 dotted rectangle</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function DrawFocusRect Lib &quot;user32&quot; Alias &quot;DrawFocusRect&quot; (ByVal hdc As Long,
 lpRect As RECT) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Invert
 colors of rectangle</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function InvertRect Lib &quot;user32&quot; Alias &quot;InvertRect&quot; (ByVal hdc As Long, lpRect
 As RECT) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get
 cursor position</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetCursorPos Lib &quot;user32&quot; Alias &quot;GetCursorPos&quot; (lpPoint As POINTAPI)
 As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Always on
 top</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function SetWindowPos Lib &quot;user32&quot; Alias &quot;SetWindowPos&quot; (ByVal hwnd As Long,
 ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As
 Long, ByVal cy As Long, ByVal wFlags As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Send
 messages to a window</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function SendMessage Lib &quot;user32&quot; Alias &quot;SendMessageA&quot; (ByVal hwnd As Long,
 ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Find
 directories</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetWindowsDirectory Lib &quot;kernel32&quot; Alias &quot;GetWindowsDirectoryA&quot; (ByVal
 lpBuffer As String, ByVal nSize As Long) As Long <br>
 Declare Function GetSystemDirectory Lib &quot;kernel32&quot; Alias &quot;GetSystemDirectoryA&quot;
 (ByVal lpBuffer As String, ByVal nSize As Long) As Long <br>
 Declare Function GetTempPath Lib &quot;kernel32&quot; Alias &quot;GetTempPathA&quot; (ByVal
 nBufferLength As Long, ByVal lpBuffer As String) As Long <br>
 Declare Function GetCurrentDirectory Lib &quot;kernel32&quot; Alias &quot;GetCurrentDirectory&quot;
 (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Text
 alignment</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetTextAlign Lib &quot;gdi32&quot; Alias &quot;GetTextAlign&quot; (ByVal hdc As Long) As
 Long <br>
 Declare Function SetTextAlign Lib &quot;gdi32&quot; Alias &quot;SetTextAlign&quot; (ByVal hdc As
 Long, ByVal wFlags As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Flash a
 title bar</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function FlashWindow Lib &quot;user32&quot; Alias &quot;FlashWindow&quot; (ByVal hwnd As Long,
 ByVal bInvert As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>
 Manipulate bitmaps</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function BitBlt Lib &quot;gdi32&quot; Alias &quot;BitBlt&quot; (ByVal hDestDC As Long, ByVal x As
 Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal
 hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
 As Long <br>
 Declare Function PatBlt Lib &quot;gdi32&quot; Alias &quot;PatBlt&quot; (ByVal hdc As Long, ByVal x
 As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal
 dwRop As Long) As Long <br>
 Declare Function StretchBlt Lib &quot;gdi32&quot; Alias &quot;StretchBlt&quot; (ByVal hdc As Long,
 ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long,
 ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth
 As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long <br>
 Declare Function CreateCompatibleBitmap Lib &quot;gdi32&quot; Alias &quot;CreateCompatibleBitmap&quot;
 (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long <br>
 Declare Function CreateCompatibleDC Lib &quot;gdi32&quot; Alias &quot;CreateCompatibleDC&quot; (ByVal
 hdc As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Rotate
 text</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function CreateFontIndirect Lib &quot;gdi32&quot; Alias &quot;CreateFontIndirectA&quot; (lpLogFont
 As LOGFONT) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Timing</b>
 </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetTickCount Lib &quot;kernel32&quot; Alias &quot;GetTickCount&quot; () As Long </font>
 </td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>File
 information</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetFileAttributes Lib &quot;kernel32&quot; Alias &quot;GetFileAttributesA&quot; (ByVal
 lpFileName As String) As Long <br>
 Declare Function GetFileSize Lib &quot;kernel32&quot; Alias &quot;GetFileSize&quot; (ByVal hFile
 As Long, lpFileSizeHigh As Long) As Long <br>
 Declare Function GetFullPathName Lib &quot;kernel32&quot; Alias &quot;GetFullPathNameA&quot; (ByVal
 lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String,
 ByVal lpFilePart As String) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get
 window information</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetClassName Lib &quot;user32&quot; Alias &quot;GetClassNameA&quot; (ByVal hwnd As Long,
 ByVal lpClassName As String, ByVal nMaxCount As Long) As Long <br>
 Declare Function GetWindowText Lib &quot;user32&quot; Alias &quot;GetWindowTextA&quot; (ByVal hwnd
 As Long, ByVal lpString As String, ByVal cch As Long) As Long <br>
 Declare Function GetParent Lib &quot;user32&quot; Alias &quot;GetParent&quot; (ByVal hwnd As Long)
 As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Identify
 window at cursor</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function WindowFromPoint Lib &quot;user32&quot; Alias &quot;WindowFromPoint&quot; (ByVal xPoint As
 Long, ByVal yPoint As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Registry
 editing</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function RegCreateKey Lib &quot;advapi32.dll&quot; Alias &quot;RegCreateKeyA&quot; (ByVal hKey As
 Long, ByVal lpSubKey As String, phkResult As Long) As Long <br>
 Declare Function RegDeleteKey Lib &quot;advapi32.dll&quot; Alias &quot;RegDeleteKeyA&quot; (ByVal
 hKey As Long, ByVal lpSubKey As String) As Long <br>
 Declare Function RegDeleteValue Lib &quot;advapi32.dll&quot; Alias &quot;RegDeleteValueA&quot; (ByVal
 hKey As Long, ByVal lpValueName As String) As Long <br>
 Declare Function RegQueryValueEx Lib &quot;advapi32.dll&quot; Alias &quot;RegQueryValueExA&quot; (ByVal
 hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As
 Long, lpData As Any, lpcbData As Long) As Long <br>
 Declare Function RegSetValueEx Lib &quot;advapi32.dll&quot; Alias &quot;RegSetValueExA&quot; (ByVal
 hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal
 dwType As Long, lpData As Any, ByVal cbData As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Drawing
 functions</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function MoveToEx Lib &quot;gdi32&quot; Alias &quot;MoveToEx&quot; (ByVal hdc As Long, ByVal x As
 Long, ByVal y As Long, lpPoint As POINTAPI) As Long <br>
 Declare Function LineTo Lib &quot;gdi32&quot; Alias &quot;LineTo&quot; (ByVal hdc As Long, ByVal x
 As Long, ByVal y As Long) As Long <br>
 Declare Function Ellipse Lib &quot;gdi32&quot; Alias &quot;Ellipse&quot; (ByVal hdc As Long, ByVal
 X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
 </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get icon
 Declare</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Function
 ExtractIcon Lib &quot;shell32.dll&quot; Alias &quot;ExtractIconA&quot; (ByVal hInst As Long, ByVal
 lpszExeFileName As String, ByVal nIconIndex As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Screen
 capture</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function SetCapture Lib &quot;user32&quot; Alias &quot;SetCapture&quot; (ByVal hwnd As Long) As
 Long <br>
 Declare Function CreateDC Lib &quot;gdi32&quot; Alias &quot;CreateDCA&quot; (ByVal lpDriverName As
 String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As
 DEVMODE) As Long <br>
 Declare Function DeleteDC Lib &quot;gdi32&quot; Alias &quot;DeleteDC&quot; (ByVal hdc As Long) As
 Long <br>
 Declare Function BitBlt Lib &quot;gdi32&quot; Alias &quot;BitBlt&quot; (ByVal hDestDC As Long,
 ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long,
 ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As
 Long) As Long <br>
 Declare Function ReleaseCapture Lib &quot;user32&quot; Alias &quot;ReleaseCapture&quot; () As Long
 <br>
 Declare Function ClientToScreen Lib &quot;user32&quot; Alias &quot;ClientToScreen&quot; (ByVal
 hwnd As Long, lpPoint As POINTAPI) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get user
 name</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetUserName Lib &quot;advapi32.dll&quot; Alias &quot;GetUserNameA&quot; (ByVal lpBuffer
 As String, nSize As Long) As LongDeclare Function GetUserName Lib
 &quot;advapi32.dll&quot; Alias &quot;GetUserNameA&quot; (ByVal lpBuffer As String, nSize As Long)
 As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get
 computer name</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetComputerName Lib &quot;kernel32&quot; Alias &quot;GetComputerNameA&quot; (ByVal
 lpBuffer As String, nSize As Long) As LongDeclare Function GetComputerName Lib
 &quot;kernel32&quot; Alias &quot;GetComputerNameA&quot; (ByVal lpBuffer As String, nSize As Long)
 As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get
 volume name/serial#</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetVolumeInformation Lib &quot;kernel32&quot; Alias &quot;GetVolumeInformationA&quot; (ByVal
 lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal
 nVolumeNameSize As Long, lpVolumeSerialNumber As Long,
 lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal
 lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
 </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Identify
 drive type</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetDriveType Lib &quot;kernel32&quot; Alias &quot;GetDriveTypeA&quot; (ByVal nDrive As
 String) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Get free
 space</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function GetDiskFreeSpace Lib &quot;kernel32&quot; Alias &quot;GetDiskFreeSpaceA&quot; (ByVal
 lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As
 Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
 </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>INI
 editing</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function WritePrivateProfileSection Lib &quot;kernel32&quot; Alias &quot;WritePrivateProfileSectionA&quot;
 (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As
 String) As Long <br>
 Declare Function WritePrivateProfileString Lib &quot;kernel32&quot; Alias &quot;WritePrivateProfileStringA&quot;
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As
 Any, ByVal lpFileName As String) As Long <br>
 Declare Function GetPrivateProfileInt Lib &quot;kernel32&quot; Alias &quot;GetPrivateProfileIntA&quot;
 (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault
 As Long, ByVal lpFileName As String) As Long <br>
 Declare Function GetPrivateProfileSection Lib &quot;kernel32&quot; Alias &quot;GetPrivateProfileSectionA&quot;
 (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As
 Long, ByVal lpFileName As String) As Long <br>
 Declare Function GetPrivateProfileString Lib &quot;kernel32&quot; Alias &quot;GetPrivateProfileStringA&quot;
 (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As
 String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal
 lpFileName As String) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Put icon
 in system tray</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function CallWindowProc Lib &quot;user32&quot; Alias &quot;CallWindowProcA&quot; (ByVal
 lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As
 Long, ByVal lParam As Long) As Long <br>
 Declare Function GetWindowLong Lib &quot;user32&quot; Alias &quot;GetWindowLongA&quot; (ByVal hwnd
 As Long, ByVal nIndex As Long) As Long <br>
 Declare Function SetWindowLong Lib &quot;user32&quot; Alias &quot;SetWindowLongA&quot; (ByVal hwnd
 As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long <br>
 Declare Function Shell_NotifyIcon Lib &quot;shell32.dll&quot; Alias &quot; Shell_NotifyIconA&quot;
 (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long <br>
 Declare Sub CopyMemory Lib &quot;kernel32&quot; Alias &quot;RtlMoveMemory&quot; (Destination As
 Any, Source As Any, ByVal Length As Long) <br>
 Declare Function DrawEdge Lib &quot;user32&quot; Alias &quot;DrawEdge&quot; (ByVal hdc As Long,
 qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Wait for
 program to stop</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function CreateProcess Lib &quot;kernel32&quot; Alias &quot;CreateProcessA&quot; (ByVal
 lpApplicationName As String, ByVal lpCommandLine As String,
 lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As
 SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As
 Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo
 As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long <br>
 Declare Function WaitForSingleObject Lib &quot;kernel32&quot; Alias &quot;WaitForSingleObject&quot;
 (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long </font></td>
 </tr>
 <tr>
 <td width="300">&nbsp;</td>
 </tr>
 <tr>
 <td vAlign="top" noWrap width="300"><font face="Verdana" size="2"><b>Stop
 ctrl-alt-del</b> </font></td>
 <td vAlign="top" noWrap width="712"><font face="Verdana" size="2">Declare
 Function SystemParametersInfo Lib &quot;user32&quot; Alias &quot;SystemParametersInfoA&quot; (ByVal
 uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni
 As Long) As Long </font></td>
 </tr>
 </table>
 <p><font face="Verdana" size="2">Hopefully, this section of the tutorial has
 sparked some excitement! You should now see that a door of tremendous
 proportions has been opened to you. You've begun to leave the limitations of VB
 behind and joined the rest of the programming community who have already been
 using the API for years. I hope to add quite a bit to this tutorial section so
 check back often over the next few weeks.</font></td>
 </tr>
</table>

