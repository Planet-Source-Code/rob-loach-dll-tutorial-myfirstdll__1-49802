<div align="center">

## DLL Tutorial \- MyFirstDLL


</div>

### Description

Learn how to not only create a DLL file, but also add it to another project, and call its functions, subs, and properties. This tutorial is step-by-step and very easy to understand. I made this because I noticed that all the other how-to DLL examples on PSC were pretty poor.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rob Loach](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob-loach.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rob-loach-dll-tutorial-myfirstdll__1-49802/archive/master.zip)





### Source Code

<div class=Section1>
<h1>DLL Tutorial:<span style="mso-spacerun: yes"> </span>MyFirstDLL</h1>
<h4>Author: Rob Loach</h4>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>In this tutorial you will learn:</p>
<p class=MsoNormal style='margin-left:54.0pt;text-indent:-18.0pt;mso-list:l1 level1
lfo4;
tab-stops:list 54.0pt'>[if !supportLists]<span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>
</span></span>[endif]How to make a DLL file.</p>
<p class=MsoNormal style='margin-left:54.0pt;text-indent:-18.0pt;mso-list:l4 level1
lfo5;
tab-stops:list 54.0pt'>[if !supportLists]<span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>
</span></span>[endif]How to call the DLL from a different project.</p>
<p class=MsoNormal style='margin-left:54.0pt;text-indent:-18.0pt;mso-list:l4 level1
lfo5;
tab-stops:list 54.0pt'>[if !supportLists]<span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>
</span></span>[endif]How to make a class with properties, subs, and
functions.</p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<h2>Welcome</h2>
<p class=MsoNormal style='text-align:justify'><b>W</b>elcome to my tutorial,
MyFirstDLL.<span style="mso-spacerun: yes"> </span>If you read through this
tutorial, and do all the coding, it will take you about 3-10 minutes for you to
fully understand how the DLL system in VB works.<span style="mso-spacerun:
yes"> </span>If you want to do it more quickly, <b>all the important
information is bolded</b>.<span style="mso-spacerun: yes"> </span>The goal of
this tutorial is to explain step-by-step how to create and use a DLL file.</p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<h2>What is a DLL?</h2>
<p class=MsoNormal style='text-align:justify'><b>A</b> DLL is a file that you
can have your application use.<span style="mso-spacerun: yes"> </span><b>A
programmer can use the functions in a DLL file, but the code itself cannot be
accessed</b>.<span style="mso-spacerun: yes"> </span>This allows you to make
various things such as game engines.<span style="mso-spacerun: yes">
</span>You can then distribute the engine to the public without actually giving
out the code.<span style="mso-spacerun: yes"> </span>DLLs are very useful
because it <b>allows you to hold a large amount of code in only one file</b>.</p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<h2>So how do I make a DLL file?</h2>
<p class=MsoNormal style='text-align:justify'><b>T</b>o make a DLL, follow
these simple steps:</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]1)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Open
Microsoft Visual Basic. </p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]2)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Goto
File <span style='font-family:Wingdings;mso-ascii-font-family:"Times New Roman";
mso-hansi-font-family:"Times New
Roman";mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>à</span></span>
New Project.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]3)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Start
an <b>ActiveX DLL</b>.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]4)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]This
new window is your DLL.<span style="mso-spacerun: yes"> </span>You currently
only have one object in it, a class.<span style="mso-spacerun: yes">
</span>Now it is time to put in the code that you want your DLL to use.<span
style="mso-spacerun: yes"> </span>In this case, just <b>add in the following
code</b>. </p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 bgcolor="#e6e6e6" style='background:
 #E6E6E6;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;
 mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=590 valign=top style='width:442.8pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier
New";color:green'>'=====================<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Title:<span style="mso-spacerun:
 yes">  </span><span style="mso-spacerun:
yes"> </span>MyFirstDLL<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Purpose:<span style="mso-spacerun:
 yes">  </span>Holds a text string and when the DisplayMsg<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>sub is called, it displays a message box
of<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>the string.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>This is just an example showing how a
class<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>file works.<span style="mso-spacerun: yes">
 </span>Now you can type CLASS and . and<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>a list of properties and subs will
appear.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Author:<span style="mso-spacerun:
 yes">  </span>Rob Loach<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier
New";color:green'>'=====================<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Variables<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'=========</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier
New"'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Private</span><span style='font-size:
 10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";color:#3366FF'>
</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier
New"'>p_Text
 <span style='color:navy'>As String</span><span
style='color:blue'><o:p></o:p></span></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Properties<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'==========</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier
New"'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Public Property Get</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New"'>
 Text() <span style='color:navy'>As String</span><span
style='color:blue'><o:p></o:p></span></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  </span>Text =
 p_Text<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>End Property</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:blue'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Public Property Let</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New"'>
 Text(<span style='color:navy'>ByVal</span><span style='color:blue'> </span>i_Text
 <span style='color:navy'>As String</span>)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  </span>p_Text
 = i_Text<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>End Property</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:blue'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Functions and Subs<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'==================</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier
New"'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Public Sub</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:blue'> </span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'>DisplayMsg()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New"'><span style="mso-spacerun: yes">  </span>MsgBox
 p_Text, vbOKOnly, "DLL Function Called"<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>End Sub</span><span style='font-family:
 "Courier New";color:navy'><o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'><b>This is your class file inside
your new DLL</b>.<span style="mso-spacerun: yes"> </span>All it will do is
hold a string called p_Text.<span style="mso-spacerun: yes"> </span>Calling
the property named Text can change the string.<span style="mso-spacerun: yes">
</span>When the DisplayMsg sub is called, it will make a messagebox saying the
Text string.</p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>You could put any code you want
into this class.<span style="mso-spacerun: yes"> </span>This is just an example
that we will be using.<span style="mso-spacerun: yes"> </span><b>You can put
anything you want into the DLL</b> file (Forms, Classes, Modules, etc).</p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]5)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Now
you have to <b>rename your class</b> file.<span style="mso-spacerun: yes">
</span>For this example, name it clsMyFirstDLL.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]6)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]You
can then <b>rename your DLL</b> to MyFirstDLL (or anything you want).<span
style="mso-spacerun: yes"> </span>Click on the ActiveX Icon in the top left of
the project window.<span style="mso-spacerun: yes">  </span>Next, in the
properties window, change the Name property to MyFirstDLL.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l2 level1 lfo2;tab-stops:list 36.0pt'>[if !supportLists]7)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Now,
once your done making your DLL, your going to have to save it as an actual
file.<span style="mso-spacerun: yes"> </span>Goto <b>File </b><b><span
style='font-family:Wingdings;mso-ascii-font-family:"Times New Roman";
mso-hansi-font-family:"Times New
Roman";mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>à</span></span>
Make MyFirstDLL.dll</b>…<span style="mso-spacerun: yes"> </span>Save it
wherever you want.<span style="mso-spacerun: yes"> </span>Just take note of
where you put it.</p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<h2>Now that I’ve made my DLL, how do I use it?</h2>
<p class=MsoNormal style='text-align:justify'><b>Y</b>ou can call the DLL a
number of ways.<span style="mso-spacerun: yes"> </span>In this tutorial, I
will only show you one.</p>
<p class=MsoNormal style='text-align:justify'><b>O</b>nce you have a DLL file,
and want to make use of it, do the following.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo3;tab-stops:list 36.0pt'>[if !supportLists]1)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Start
Microsoft Visual Basic.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo3;tab-stops:list 36.0pt'>[if !supportLists]2)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Start
a Standard EXE.<span style="mso-spacerun: yes"> </span>This will be the new
project that will use the DLL.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo3;tab-stops:list 36.0pt'>[if !supportLists]3)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]<span
style="mso-spacerun: yes"> </span>Goto <b>Project </b><b><span
style='font-family:Wingdings;mso-ascii-font-family:"Times New Roman";
mso-hansi-font-family:"Times New
Roman";mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>à</span></span>
References</b>.<span style="mso-spacerun: yes"> </span>It will take some time
to load.<span style="mso-spacerun: yes"> </span>This is a list of DLL files
that the application is currently using.<span style="mso-spacerun: yes">
</span>Click on <b>browse and load the DLL</b> that you just made.<span
style="mso-spacerun: yes"> </span>It will then add the DLL to the list.<span
style="mso-spacerun: yes"> </span>Now click on OK.<span style="mso-spacerun:
yes"> </span>Your project has now successfully loaded the DLL information,
functions, and properties into memory.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo3;tab-stops:list 36.0pt'>[if !supportLists]4)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Now
is the time to use the DLL and see how the class works within the DLL.<span
style="mso-spacerun: yes"> </span><b>Make two command buttons, one named
Command1 and the other named Command2.<span style="mso-spacerun: yes">
</span>Next, make a textbox and name it Text1</b>.<span style="mso-spacerun:
yes"> </span>These are going to be used in this example.</p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo3;tab-stops:list 36.0pt'>[if !supportLists]5)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]<span
style="mso-spacerun: yes"> </span>Now view the code of the form and put in the
following code:</p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<table border=1 cellspacing=0 cellpadding=0 bgcolor="#e6e6e6" style='background:
 #E6E6E6;border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;
 mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
 <tr>
 <td width=590 valign=top style='width:442.8pt;border:solid windowtext .5pt;
 padding:0cm 5.4pt 0cm 5.4pt'>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier
New";color:green'>'=====================================<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Title:<span style="mso-spacerun:
 yes">   </span>MyFirstDLL Application Use<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Purpose:<span style="mso-spacerun:
 yes">  </span>An application that uses the DLL that was just
made.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>It requires a form (Form1)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>two commands (Command1 and Command2)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'<span style="mso-spacerun:
 yes">      </span>a text box (Text1).<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Author:<span style="mso-spacerun:
 yes">  </span>Rob Loach<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier
New";color:green'>'=====================================<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Variables<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'=========<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'Make a variable that uses the class
 in the<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>'MyFirstDLL DLL file so that we can
 make use of it.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Dim</span><span style='font-size:10.0pt;
 mso-bidi-font-size:12.0pt;font-family:"Courier New";color:green'> </span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:black'>MyFirstDLL</span><span style='font-size:10.0pt;mso-bidi-font-size:
 12.0pt;font-family:"Courier New";color:green'> </span><span style='font-size:
 10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";color:navy'>As
New</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'> </span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>clsMyFirstDLL</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Private Sub</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'> </span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>Command1_Click()</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'> 'Set text<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span>'Set the property of MyFirstDLL to text1 text<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span>'This shows how to set a property of the class/DLL.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span></span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>MyFirstDLL.Text = Text1.Text</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>End Sub</span><span style='font-size:
 10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier
New";color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Private Sub</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'> </span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>Command2_Click()</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span>'Call the sub DisplayMsg in the DLL.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span>'This shows how to call a sub/function of the
class/DLL.<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span></span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>MyFirstDLL.DisplayMsg</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>End Sub</span><span style='font-size:
 10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier
New";color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'>[if
!supportEmptyParas] [endif]<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>Private Sub</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'> </span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>Form_Load()<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span>'Initialize the Objects<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:green'><span style="mso-spacerun: yes">
 </span></span><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'>Form1.Caption =
"MyFirstDLL"<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'><span style="mso-spacerun: yes">
 </span>Text1.Text = "Enter text to be displayed
here..."<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'><span style="mso-spacerun: yes">
 </span>Command1.Caption = "Set Text"<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:black'><span style="mso-spacerun: yes">
 </span>Command2.Caption = "Display Text"</span><span
 style='font-size:10.0pt;mso-bidi-font-size:12.0pt;font-family:"Courier New";
 color:green'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:10.0pt;mso-bidi-font-size:12.0pt;
 font-family:"Courier New";color:navy'>End Sub</span><span style='font-family:
 "Courier New";color:blue'><o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l0 level1 lfo3;tab-stops:list 36.0pt'>[if !supportLists]6)<span
style='font:7.0pt "Times New Roman"'>
</span>[endif]Now
run the program and play around with it.<span style="mso-spacerun: yes">
</span>As you can see, when you click Set Text, it sets the property “Text” in
MyFirstDLL to whatever you typed in.<span style="mso-spacerun: yes"> </span>When
you click Display Text, it calls the sub DisplayMsg in the DLL.</p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal>[if !supportEmptyParas] [endif]<o:p></o:p></p>
<h2>Quick Things To Remember</h2>
<p class=MsoNormal style='margin-left:54.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l4 level1 lfo5;tab-stops:list 54.0pt'>[if !supportLists]<span
style='font-family:Symbol'>·<span style='font:7.0pt "Times New
Roman"'>
</span></span>[endif]To make a DLL, use ActiveX DLL project.</p>
<p class=MsoNormal style='margin-left:54.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l4 level1 lfo5;tab-stops:list 54.0pt'>[if !supportLists]<span
style='font-family:Symbol'>·<span style='font:7.0pt "Times New
Roman"'>
</span></span>[endif]Project <span
style='font-family:Wingdings;mso-ascii-font-family:
"Times New Roman";mso-hansi-font-family:"Times New Roman";mso-char-type:symbol;
mso-symbol-font-family:Wingdings'><span
style='mso-char-type:symbol;mso-symbol-font-family:
Wingdings'>à</span></span> References</p>
<p class=MsoNormal style='margin-left:54.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l4 level1 lfo5;tab-stops:list 54.0pt'>[if !supportLists]<span
style='font-family:Symbol'>·<span style='font:7.0pt "Times New
Roman"'>
</span></span>[endif]Keep your DLLs in the same directory as the project.</p>
<p class=MsoNormal style='margin-left:54.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l4 level1 lfo5;tab-stops:list 54.0pt'>[if !supportLists]<span
style='font-family:Symbol'>·<span style='font:7.0pt "Times New
Roman"'>
</span></span>[endif]Name everything to keep organization.</p>
<p class=MsoNormal style='margin-left:54.0pt;text-align:justify;text-indent:
-18.0pt;mso-list:l4 level1 lfo5;tab-stops:list 54.0pt'>[if !supportLists]<span
style='font-family:Symbol'>·<span style='font:7.0pt "Times New
Roman"'>
</span></span>[endif]Make sure to use as many variables as possible in every
sub/function in your DLLs to allow diversity of programs.</p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<h2>Conclusion</h2>
<p class=MsoNormal>That concludes this tutorial!<span style="mso-spacerun:
yes"> </span>In it you learned how to make a DLL and use it in a different
program.<span style="mso-spacerun: yes"> </span>Thank you for reading MyFirstDLL
and I hope you have learned the DLL-VB concept. </p>
<p class=MsoNormal style='margin-left:36.0pt;text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
<p class=MsoNormal style='text-align:justify'>[if
!supportEmptyParas] [endif]<o:p></o:p></p>
</div>

