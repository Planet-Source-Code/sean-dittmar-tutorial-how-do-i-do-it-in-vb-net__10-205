<div align="center">

## Tutorial: How Do I Do It In VB\.NET??


</div>

### Description

This tutorial teaches syntax migration from VB6 to VB.NET with simple VB7/VB.NET comparison.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sean Dittmar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-dittmar.md)
**Level**          |Beginner
**User Rating**    |4.7 (141 globes from 30 users)
**Compatibility**  |VB\.NET
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__10-33.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sean-dittmar-tutorial-how-do-i-do-it-in-vb-net__10-205/archive/master.zip)





### Source Code

		<center><STRONG><U>Tutorial: How do I do it in VB.NET?</U></STRONG></center>
		<CENTER> </CENTER>
		<CENTER>This tutorial explain some of the more common migrating problems from VB to
			VB.NET.
		</CENTER>
		<CENTER> </CENTER>
		<UL>
			<LI>
				<DIV align="left"><STRONG>DoEvents</STRONG>
					<BR>
					<BR>
					VB6<BR>
					<FONT face="Courier New Baltic">DoEvents<BR>
					</FONT>
					<BR>
					VB7 <BR>
					<FONT face="Courier New">System.Windows.Forms.Application.DoEvents</FONT>
					<BR>
				</DIV>
			</LI>
		</UL>
<UL>
 <LI>
 <DIV align="left"><STRONG>App Object<BR>
  </STRONG> <BR>
  <FONT color="red"><EM>Get the full application filepath</EM><br>
  <BR>
  </FONT>VB6 <BR>
  <FONT face="Courier New">App.Path & App.EXEName</FONT><BR>
  <BR>
  VB7 <FONT face="Courier New">System.Reflection.Assembly.GetExecutingAssembly.Location.ToString<BR>
  </FONT> <BR>
  <FONT color="red"><EM>Get the app's instance</EM><br>
  <BR>
  </FONT>VB6 <FONT face="Courier New">App.hInstance</FONT><BR>
  <BR>
  VB7<br>
  <FONT face="Courier New">System.Runtime.InteropServices.Marshal.GetHINSTANCE
  _(System.Reflection.Assembly.GetExecutingAssembly.GetModules() _(0)).ToInt32()</FONT><BR>
  <BR>
  <FONT color="red"><EM>Check for a previous instance</EM><br>
  <BR>
  </FONT>VB6 <br>
  <FONT face="Courier New">App.PrevInstance <BR>
  <BR>
  VB7<BR>
  <FONT face="Courier New">Function PrevInstance() As Boolean<BR>
  </FONT><FONT face="Courier New">If Ubound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess)</FONT><FONT face="Courier New">.ProcessName))
  > 0 Then</FONT><FONT face="Courier New"><BR>
  </FONT><FONT face="Courier New">     Return True <BR>
  </FONT><FONT face="Courier New">Else <BR>
       Return False  <BR>
  </FONT><FONT face="Courier New">End If<BR>
  <BR>
  End Sub<BR>
  </FONT></FONT> </DIV>
 </LI>
</UL>
<UL>
 <LI>
 <DIV align="left"><FONT face="Courier New"><FONT face="Times New Roman"><STRONG>Graphics<BR>
  <BR>
  </STRONG></FONT></FONT><FONT face="Courier New"><FONT face="Times New Roman"><EM><FONT color="red">
  Load a picture</FONT><br>
  <BR>
  </EM>VB6 <BR>
  </FONT><FONT face="Courier New">Picture1.Picture = LoadPicture(<EM>path</EM>)<BR>
  </FONT></FONT><FONT face="Courier New"><FONT face="Courier New"><FONT face="Times New Roman">
  <BR>
  VB7 <BR>
  </FONT></FONT><FONT face="Courier New">Dim img As Image = Image.FromFile(<EM>path</EM>)<BR>
  Picture1.Image = img<BR>
  <BR>
  <FONT face="Times New Roman"><EM><FONT color="red">Load a icon<br>
  <BR>
  </FONT></EM>VB6<BR>
  </FONT></FONT></FONT><FONT face="Courier New">Me.Icon = LoadPicture(<EM>path</EM>)<BR>
  </FONT><FONT face="Courier New"><FONT face="Times New Roman"> <BR>
  VB7<BR>
  </FONT></FONT><FONT face="Courier New">Dim ico As New Icon(<EM>path</EM>)<BR>
  Me.Icon = ico<BR>
  </FONT><FONT face="Courier New"> <BR>
  </FONT> </DIV>
 <LI>
 <DIV align="left"><FONT face="Courier New"><STRONG><FONT face="Times New Roman">File
  I/0<BR>
  </FONT></STRONG></FONT><FONT face="Courier New"> <BR>
  <FONT face="Times New Roman"><EM><FONT color="red">Read from a file</FONT><br>
  <BR>
  </EM>VB6<BR>
  </FONT></FONT><FONT face="Courier New">Open <EM>path </EM>For Input As #1<BR>
  Line Input #1, buffer<BR>
  Close #1<BR>
  <BR>
  </FONT><FONT face="Times New Roman">VB7<BR>
  </FONT><FONT face="Courier New">Dim fs As FileStream = File.Open<EM>(path,</EM> FileMode.OpenOrCreate,
  _ FileAccess.Read)<BR>
  </FONT><FONT face="Courier New">Dim sr As New StreamReader(fs) <BR>
  Buffer = sr.ReadLine<BR>
  sr.Close<BR>
  <BR>
  <FONT face="Times New Roman"><EM><FONT color="red">Write to a file</FONT><br>
  <BR>
  </EM>VB6<BR>
  </FONT></FONT><font face="Courier New">Open <EM>path </EM>For Output As
  #1<BR>
  Write #1, buffer<BR>
  Close #1</font><FONT face="Courier New"></FONT><FONT face="Courier New"><br>
  <BR>
  </FONT><FONT face="Courier New"><FONT face="Times New Roman">VB7 <FONT face="Courier New">
  <BR>
  Dim fs As FileStream = File.Open(<EM>path</EM>, FileMode.OpenOrCreate, _<BR>
  FileAccess.Write)<BR>
  Dim sr As New StreamWriter(fs)<BR>
  sr.Write(buffer)<BR>
  sr.Close<br>
  </FONT></FONT></FONT><FONT face="Courier New"><FONT face="Times New Roman"><FONT face="Courier New"><BR>
  </FONT></FONT></FONT> </DIV>
 <LI>
 <DIV align="left"><FONT face="Courier New"><FONT face="Times New Roman"><FONT face="Courier New"><FONT face="Times New Roman"><STRONG>Errors<BR>
  <BR>
  </STRONG><EM><FONT color="red">Check for an error</FONT></EM></FONT></FONT></FONT></FONT><FONT face="Courier New"><FONT face="Times New Roman"><FONT face="Courier New"><FONT face="Times New Roman"><br>
  <BR>
  VB6<BR>
  <FONT face="Courier New">On Error Goto errhandler<BR>
  ...<BR>
  errhandler:<BR>
  MsgBox(err.Description)<BR>
  <BR>
  </FONT></FONT></FONT></FONT></FONT><FONT face="Times New Roman">VB7<BR>
  <FONT face="Courier New">Try<BR>
       ...<BR>
       Throw New Exception("error description goes here")<BR>
       ...<BR>
  Catch e as Exception<BR>
       MsgBox(e.Description)<BR>
  End Try<br>
  <br>
  </FONT></FONT></DIV>
 </LI>
 <LI>
 <DIV align="left"><font face="Times New Roman, Times, serif"><b>Events </b></font><FONT face="Times New Roman"><FONT face="Courier New"><br>
  <br>
  <font face="Times New Roman, Times, serif"><i><font color="#FF0000">Handling
  an event</font><br>
  </i> <br>
  In VB7, there is a new keyword called AddHandler. AddHandler makes handling
  events a snap.<br>
  <br>
  <font face="Courier New, Courier, mono">AddHandler <i>object.event, </i>AddressOf<i>
  procedure</i></font></font></FONT></FONT></DIV>
 </LI>
</UL>

