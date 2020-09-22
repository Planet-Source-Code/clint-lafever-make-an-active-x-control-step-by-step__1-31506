<div align="center">

## Make an Active X Control \(Step by Step\)


</div>

### Description

A step by step article on how to make an Active X control. In my eyes, I am not an expert but somebody asked me to write this (past comment on another submission) so I did. If you follow it, in the end you will have your own Custom PictureBox control that will have a property to assign a URL to an image to use for its picture along with an event to know when it completed its download. Hope you all enjoy. I did not realize it would take me so long to write.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Clint LaFever](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/clint-lafever.md)
**Level**          |Intermediate
**User Rating**    |4.9 (153 globes from 31 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/clint-lafever-make-an-active-x-control-step-by-step__1-31506/archive/master.zip)





### Source Code

<p>&nbsp;</p>
<p><small><b><font face="Verdana">How to build an Active X Control (Basic
Tutorial)</font></b></small></p>
<p><small><font face="Verdana">I cannot believe I am going to try to attempt to
teach since I never thought of myself as a teacher, but in one of my previous
postings, a comment asked if I could post a tutorial on how to make an Active X
Control because that person liked the way I explain things.&nbsp; So here I
go.&nbsp; I would like to request that if I mistake anything or call something
by it's wrong name that you do not flame me.&nbsp; Feel free to comment and let
others know of my mistake, but please, be nice :)&nbsp; Ok, here goes.</font></small></p>
<p><small><font face="Verdana">This tutorial is going to walk you through step
by step of how to create a new PictureBox control that will have a new property
to supply a URL to an image on the web to use as it's picture (without the use
of Winsock or Internet controls).&nbsp; I think everyone would find a use for
this type of control.&nbsp;</font></small></p>
<p><small><font face="Verdana">Instructions assume you are using Visual Basic
6.0</font></small></p>
<p><small><font face="Verdana">Open VB</font></small></p>
<p><small><font face="Verdana">Choose to start a New Active X Control Project:</font></small></p>
<p><small><b><font face="Verdana">Default Naming:</font></b></small></p>
<p><small><font face="Verdana">Easy enough right.&nbsp; Ok, first things first,
we need to name a few things and set some project properties.&nbsp; Click on the
Project Explorer Tree on the Project itself (PROJECT1) and then down in the
properties window, rename it to: WEBPIC.&nbsp; This name is going to become the
name of your OCX (duh).&nbsp; Then click on the user control branch and rename
it to: WebPictureBox.&nbsp; This name is the name it will be known as inside of
VB (the tool tip on the tool when you put it in your available components later
on for new projects that use it)</font></small></p>
<p><small><font face="Verdana">Now go up to your menu and Choose Project.&nbsp;
Then Choose WebPic Properties.&nbsp; In Project Description enter Web Picture
Box.&nbsp; This is the name it will be listed as when you pull up the list of
available controls to add to a project.&nbsp; You want to keep it english like
so you can tell what it is.&nbsp; I hate those who make controls bet never set
this and then it will default to the name of the OCX which most of the time is
some abbreviated name that does not make too much sense when you are just
skimming though.&nbsp; Anyhow, that is a different story.&nbsp; Go ahead and set
all the other properties you want about the project, Company, Copyright
etc.&nbsp;&nbsp; I personally like to set auto increment on the version tab.
Click Ok.</font></small></p>
<p><small><b><font face="Verdana">Start of Coding:</font></b></small></p>
<p><small><font face="Verdana">Ok, before we get started, Save your work
(however you like to save where ever you want)</font></small></p>
<p><small><font face="Verdana">Place a Picture Box on the UserControl (any where
you like, code with handle it's position later).&nbsp; Name it: picBOX.</font></small></p>
<p><font face="Verdana"><small>Double Click on an empty spot on the User </small><small>Control</small><small>.&nbsp;
This should take you to UserControl_Initialize().&nbsp; In that Sub type:</small></font></p>
<p><small><font face="Courier New" color="#000080">Private Sub
UserControl_Initialize()<br>
&nbsp;&nbsp;&nbsp; With UserControl<br>
&nbsp;&nbsp;&nbsp; .picBOX.Move 0, 0, .ScaleWidth, .ScaleHeight<br>
&nbsp;&nbsp;&nbsp; End With<br>
End Sub</font></small></p>
<p><small><font face="Verdana">This code make the picture box match the size of
the control when it is first placed on a form later.</font></small></p>
<p><small><font face="Verdana">Note, instead of ME, you say UserControl when
referring to your object.&nbsp; Me refers it to its exposed methods and
properties that we will put in soon.</font></small></p>
<p><small><font face="Verdana">Now we need to code for when the user control
gets resized.&nbsp; Go to the Resize Event for the UserControl and type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Private Sub UserControl_Resize()<br>
&nbsp;&nbsp;&nbsp; If m_privateResize = False Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; With UserControl<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .picBOX.Move 0, 0, .ScaleWidth, .ScaleHeight<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End With<br>
&nbsp;&nbsp;&nbsp; End If<br>
End Sub</font></small></p>
<p><small><font face="Verdana">Pretty much the same as before, but this just
keeps the picture box the same size as the user control always.&nbsp; However, I
added a bit to check to see if code told it to resize or did the user do
it.&nbsp; Later down I have code resizing the control and I don't want this to
fire.&nbsp; Up in the General Declarations you need to defind m_privateResize as
Boolean, up top type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Private m_privateResize As Boolean</font></small></p>
<p><small><b><font face="Verdana">Adding Events:</font></b></small></p>
<p><small><font face="Verdana">Now lets just put in some basic events that our
control will have.&nbsp; You can add more later after you see how this is
done.&nbsp; I am going to add the Click, DblClick, MouseUp, MouseMove, MouseDown,
and Resize events to my control.</font></small></p>
<p><small><font face="Verdana">At the top of your code for the control in the
General Declarations section type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Event Click()&nbsp;<br>
Event DblClick()&nbsp;<br>
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)<br>
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)&nbsp;<br>
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)&nbsp;<br>
Event Resize()&nbsp;</font></small></p>
<p><small><font face="Verdana">These are now events that you can raise later in
your code.&nbsp; Which we will program now.&nbsp; Pretty much, we want to say in
our control when somebody clicks on the picture box (or moves etc) we want to
raise those events out of our control.&nbsp; So in your code type:</font></small></p>
<p><font face="Courier New" color="#000080"><small>Private Sub picBOX_Click()<br>
&nbsp;&nbsp;&nbsp; RaiseEvent Click<br>
End Sub<br>
Private Sub picBOX_DblClick()<br>
&nbsp;&nbsp;&nbsp; RaiseEvent DblClick<br>
End Sub<br>
Private Sub picBOX_MouseDown(Button As Integer, _<br>
&nbsp;&nbsp;&nbsp; Shift As Integer, X As Single, Y As Single)<br>
&nbsp;&nbsp;&nbsp; RaiseEvent MouseDown(Button, Shift, X, Y)<br>
End Sub<br>
Private Sub picBOX_MouseMove(Button As Integer, _<br>
&nbsp;&nbsp;&nbsp; Shift As Integer, X As Single, Y As Single)<br>
&nbsp;&nbsp;&nbsp; RaiseEvent MouseMove(Button, Shift, X, Y)<br>
End Sub<br>
Private Sub picBOX_MouseUp(Button As Integer, _<br>
&nbsp;&nbsp;&nbsp; Shift As Integer, X As Single, Y As Single)<br>
&nbsp;&nbsp;&nbsp; RaiseEvent MouseUp(Button, Shift, X, Y)<br>
End Sub</small></font></p>
<p><small><font face="Verdana">Pretty simple, really just raising our event when
the corresponding events occur within our control.</font></small></p>
<p><small><font face="Verdana">Now I want to add the code for when the control
itself gets resized.</font></small></p>
<p><small><font face="Verdana">Go back to your UserControl_Resize code you typed
earlier and add the RaiseEvent Resize line so it look like this:</font></small></p>
<p><small><font face="Courier New" color="#000080">Private Sub
UserControl_Resize()<br>
&nbsp;&nbsp;&nbsp;  If m_privateResize = False Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    With UserControl<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .picBOX.Move 0, 0,
.ScaleWidth, .ScaleHeight<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    End With<br>
&nbsp;&nbsp;&nbsp;  End If<br>
&nbsp;&nbsp;&nbsp;  RaiseEvent Resize<br>
End Sub</font></small></p>
<p><small><b><font face="Verdana">Properties:</font></b></small></p>
<p><small><font face="Verdana">Now we need some basic properties, really the
same ones as the picture box,&nbsp; I will skip some just to keep this quick.</font></small></p>
<p><small><font face="Verdana">We are going to add the: Appearance, BackColor, BorderStyle,
AutoRedraw, AutoSize, and Picture property to our control.</font></small></p>
<p><small><font face="Verdana">In your code type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Public Property Get Appearance() As Integer<br>
&nbsp;&nbsp;&nbsp; Appearance = picBOX.Appearance<br>
End Property<br>
Public Property Let Appearance(ByVal New_Appearance As Integer)<br>
&nbsp;&nbsp;&nbsp; picBOX.Appearance() = New_Appearance<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "Appearance"<br>
End Property<br>
Public Property Get BackColor() As OLE_COLOR<br>
&nbsp;&nbsp;&nbsp; BackColor = picBOX.BackColor<br>
End Property<br>
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)<br>
&nbsp;&nbsp;&nbsp; picBOX.BackColor() = New_BackColor<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "BackColor"<br>
End Property<br>
Public Property Get BorderStyle() As Integer<br>
&nbsp;&nbsp;&nbsp; BorderStyle = picBOX.BorderStyle<br>
End Property<br>
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)<br>
&nbsp;&nbsp;&nbsp; picBOX.BorderStyle() = New_BorderStyle<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "BorderStyle"<br>
End Property<br>
Public Property Get AutoRedraw() As Boolean<br>
&nbsp;&nbsp;&nbsp; AutoRedraw = picBOX.AutoRedraw<br>
End Property<br>
Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)<br>
&nbsp;&nbsp;&nbsp; picBOX.AutoRedraw() = New_AutoRedraw<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "AutoRedraw"<br>
End Property<br>
Public Property Get AutoSize() As Boolean<br>
&nbsp;&nbsp;&nbsp; AutoSize = picBOX.AutoSize<br>
End Property<br>
Public Property Let AutoSize(ByVal New_AutoSize As Boolean)<br>
&nbsp;&nbsp;&nbsp; picBOX.AutoSize() = New_AutoSize<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "AutoSize"<br>
End Property<br>
Public Property Get Picture() As Picture<br>
&nbsp;&nbsp;&nbsp; Set Picture = picBOX.Picture<br>
End Property<br>
Public Property Set Picture(ByVal New_Picture As Picture)<br>
&nbsp;&nbsp;&nbsp; Set picBOX.Picture = New_Picture<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "Picture"<br>
End Property<br>
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)<br>
&nbsp;&nbsp;&nbsp; picBOX.Appearance = PropBag.ReadProperty("Appearance", 1)<br>
&nbsp;&nbsp;&nbsp; picBOX.BackColor = PropBag.ReadProperty("BackColor", &amp;H8000000F)<br>
&nbsp;&nbsp;&nbsp; picBOX.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)<br>
&nbsp;&nbsp;&nbsp; picBOX.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)<br>
&nbsp;&nbsp;&nbsp; picBOX.AutoSize = PropBag.ReadProperty("AutoSize", False)<br>
&nbsp;&nbsp;&nbsp; Set Picture = PropBag.ReadProperty("Picture", Nothing)<br>
End Sub<br>
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("Appearance", picBOX.Appearance, 1)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("BackColor", picBOX.BackColor, &amp;H8000000F)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("BorderStyle", picBOX.BorderStyle, 1)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("AutoRedraw", picBOX.AutoRedraw, False)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("AutoSize", picBOX.AutoSize, False)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("Picture", Picture, Nothing)<br>
End Sub</font><font face="Verdana"><br>
</font></small></p>
<p><font face="Verdana"><small>Ok, to explain all that, in general you just made
properties to your </small><small>control</small><small> that when they get set,
will then in turn set properties of the picture box inside of your
control.&nbsp; The ReadProperties and WriteProperties as subs that will save
these properties to the property bag of the control so it remembers what you set
even after you close.&nbsp; As you typed those lines (if you did not copy/paste)
then you would have noticed what each of those </small><small>arguments</small><small>
are, Name, Value, Default.&nbsp; </small><small>Actually</small><small> pretty
easy to understand I think.</small></font></p>
<p><small><b><font face="Verdana">Tweaking:</font></b></small></p>
<p><small><font face="Verdana">Ok, with that added, we need to tweak a few
things now.&nbsp; One thing that stands out is the AutoResize Event.&nbsp; Right
now our control is coded that if the developer resizes the user control, we
resize the picture box to fit.&nbsp; But what happens when AutoSize is set to
true and a new picture gets assigned.&nbsp; The picture box will change
size.&nbsp; Therefore, we need to code to make the usercontrol match back to the
size of the new picture loaded.&nbsp; So, in the &quot;Public Property Set
Picture&quot; sub, we need to add some code to make it look like this:</font></small></p>
<p><small><font face="Courier New" color="#000080">Public Property Set Picture(ByVal New_Picture As Picture)<br>
&nbsp;&nbsp;&nbsp; Set picBOX.Picture = New_Picture<br>
&nbsp;&nbsp;&nbsp; If Me.AutoSize = True Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; With UserControl<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; m_privateResize = True<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .Width = .picBOX.Width<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .Height = .picBOX.Height<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; m_privateResize = False<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End With<br>
&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp; PropertyChanged "Picture"<br>
End Property</font></small></p>
<p><small><font face="Verdana">There, that takes care of that.&nbsp; You may
find other areas to tweak, but I am just building this the same time I am typing
so I have not thought of any yet.</font></small></p>
<p><small><font face="Verdana"><b>Custom Properties:</b></font></small></p>
<p><small><font face="Verdana">Time to add our custom properties which makes our
new version of a picture box different from the default one.&nbsp; We need a new
property named PictureURL that will contain a string to a fully qualified URL to
an image on the web.&nbsp; Because this property does not correspond back to
some other property already of the picture box, we need a place to store it when
it gets set.&nbsp; So up in the General Declarations section type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Const m_def_PictureURL = ""<br>
Private m_PictureURL As String</font></small></p>
<p><small><font face="Verdana">Now in code type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Public Property Get PictureURL() As String<br>
&nbsp;&nbsp;&nbsp; PictureURL = m_PictureURL<br>
End Property<br>
Public Property Let PictureURL(ByVal New_PictureURL As String)<br>
&nbsp;&nbsp;&nbsp; m_PictureURL = New_PictureURL<br>
  PropertyChanged "PictureURL"<br>
End Property<br>
Private Sub UserControl_InitProperties()<br>
&nbsp;&nbsp;&nbsp; m_PictureURL = m_def_PictureURL<br>
End Sub</font></small></p>
<p><small><font face="Verdana">This is the statements to read and write to this
property.&nbsp; It will save and read from the m_PictureURL variable we defined
above and use the m_def_PictureURL constant as default the first time this
control is initialized.</font></small></p>
<p><small><font face="Verdana">However, now we need to go back to our
ReadProperties and WriteProperties to make sure we tell the property bag to
remember what ever gets set here in design time.</font></small></p>
<p><small><font face="Courier New" color="#000080">Private Sub UserControl_ReadProperties(PropBag As PropertyBag)<br>
&nbsp;&nbsp;&nbsp; picBOX.Appearance = PropBag.ReadProperty("Appearance", 1)<br>
&nbsp;&nbsp;&nbsp; picBOX.BackColor = PropBag.ReadProperty("BackColor", &amp;H8000000F)<br>
&nbsp;&nbsp;&nbsp; picBOX.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)<br>
&nbsp;&nbsp;&nbsp; picBOX.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)<br>
&nbsp;&nbsp;&nbsp; picBOX.AutoSize = PropBag.ReadProperty("AutoSize", False)<br>
&nbsp;&nbsp;&nbsp; Set Picture = PropBag.ReadProperty("Picture", Nothing)<br>
&nbsp;&nbsp;&nbsp; m_PictureURL = PropBag.ReadProperty("PictureURL", m_def_PictureURL)<br>
End Sub<br>
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("Appearance", picBOX.Appearance, 1)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("BackColor", picBOX.BackColor, &amp;H8000000F)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("BorderStyle", picBOX.BorderStyle, 1)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("AutoRedraw", picBOX.AutoRedraw, False)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("AutoSize", picBOX.AutoSize, False)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("Picture", Picture, Nothing)<br>
&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty("PictureURL", m_PictureURL, m_def_PictureURL)<br>
End Sub</font></small></p>
<p><small><font face="Verdana">Note the last lines in each sub.&nbsp; We are
saving and reading the private variable we defined and storing that.&nbsp; Once
again, the property bag is what remembers what you set in the property window
when you design a form and place controls on it.&nbsp; If you do not use the
property bag, no matter what you set on the property window later will never be
saved.</font></small></p>
<p><small><font face="Verdana">Ok now, CLICK SAVE.&nbsp; You do not want to lose
what you have done so far.</font></small></p>
<p><small><b><font face="Verdana">Finishing our Custom Properties:</font></b></small></p>
<p><small><font face="Verdana">Ok, now we need to add the bit that gets an image
from the web.&nbsp; We need to go back to our &quot;Public Property Let PictureURL&quot;
sub and add some code.&nbsp; This code will get the image from the web for
us.&nbsp; Type to make our &quot;Public Property Let PictureURL&quot; sub look
like this:</font></small></p>
<p><small><font face="Courier New" color="#000080">Public Property Let PictureURL(ByVal New_PictureURL As String)<br>
&nbsp;&nbsp;&nbsp; m_PictureURL = New_PictureURL<br>
&nbsp;&nbsp;&nbsp; If (New_PictureURL &lt;> "") Then<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; AsyncRead m_PictureURL, vbAsyncTypePicture, "PictureURL", vbAsyncReadForceUpdate<br>
&nbsp;&nbsp;&nbsp; End If<br>
&nbsp;&nbsp;&nbsp; PropertyChanged &quot;PictureURL&quot;<br>
End Property</font></small></p>
<p><small><font face="Verdana">This uses the AsyncRead method to get an image
from the web.&nbsp; Pretty simple huh :)&nbsp; The vbAsyncReadForceUpdate
argument tells the AsyncRead to always get the picture from the web and ignore
any cached copy.&nbsp; Maybe later you can change this and provide some new
property to have this as a setting.&nbsp; (nice upgrade to practice with)</font></small></p>
<p><small><font face="Verdana">Ok almost done.&nbsp; The code above just starts
the download.&nbsp; Now we need to get the picture when it is done.&nbsp; For
this we use the AsyncReadComplete event of our user control.&nbsp; Go ahead and
type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)<br>
&nbsp;&nbsp;&nbsp; On Error Resume Next<br>
&nbsp;&nbsp;&nbsp; Select Case AsyncProp.PropertyName<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Case "PictureURL"<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Set Me.Picture = AsyncProp.Value<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Case Else<br>
&nbsp;&nbsp;&nbsp; End Select<br>
End Sub</font></small></p>
<p><small><font face="Verdana">The user controls AsyncReadComplete event is
fired when the download is done.&nbsp; So we ready the AsyncProp object to
determine what was just downloaded.&nbsp; In the earlier code when we started
the download, we supplied a name &quot;PictureURL&quot; as the name of the
download (not the file name, but the name associated with the download.&nbsp;
Just like you name controls when you code).&nbsp; We check to see if this
downloaded file is the one we requested, if it is, assign it to the picture
property of the picture box.&nbsp; It is written this way to help you see that
you can add more capability to this if you wish and provide multiple downloads
and what not.</font></small></p>
<p><small><b><font face="Verdana">Last Bit Of Code:</font></b></small></p>
<p><small><font face="Verdana">Ok, we are pretty much done, but I think it would
be nice to add an event to our control to tell the user/developer using it, that
the download is done.&nbsp; So back up in the General Declarations type:</font></small></p>
<p><small><font face="Courier New" color="#000080">Event DownloadComplete()</font></small></p>
<p><small><font face="Verdana">Then back in the&nbsp; &quot;Private Sub
UserControl_AsyncReadComplete&quot; sub, add a line to make it look like this:</font></small></p>
<p><small><font face="Courier New" color="#000080">Private Sub
UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)<br>
&nbsp;&nbsp;&nbsp;  On Error Resume Next<br>
&nbsp;&nbsp;&nbsp;  Select Case AsyncProp.PropertyName<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    Case &quot;PictureURL&quot;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;      Set Me.Picture = AsyncProp.Value<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    Case Else<br>
&nbsp;&nbsp;&nbsp;  End Select<br>
&nbsp;&nbsp;&nbsp; RaiseEvent DownloadComplete<br>
End Sub</font></small></p>
<p><small><font face="Verdana">There, the code side of things is done.&nbsp;
Click Save.&nbsp; Ok, go for the first compile.&nbsp; Fix any typos and try to
compile again until you get a good compile.&nbsp; Now go back to your menu,
choose Project, Choose WebPic Properties.&nbsp; Click the component tab.&nbsp;
Turn on the option for Binary Compatibility (it should be defaulted pointing to
the ocx you just made)</font></small></p>
<p><small><font face="Verdana">The reason for this is to make it that when you
make changes to your control it will make it so programs already compiled with
earlier versions of your control will still work.&nbsp; However, depending on
your changes, it may warn you that you are breaking compatibility.&nbsp; If you
break it, you should consider compiling under a new name if other programs exist
using your older version that are already compiled and released.&nbsp; If there
are no other programs released, you can break it and then try not to
again.&nbsp; Breakage occurs when you alter the declaration of an exposed method
or property that already existed in the older version.&nbsp; For example, if you
right now have the PictureURL property but later decided to call it just URL, it
will break compatibility.&nbsp; But if you add new properties or methods, or
just alter code inside existing functions or subs, it will not break.&nbsp;</font></small></p>
<p><small><b><font face="Verdana">Finishing up:</font></b></small></p>
<p><small><font face="Verdana">Now go back to your design view of your user
control. I suggest sizing the user control down to a better size.&nbsp;
Remember, the size you set here will become the default size of the control when
it later gets placed on a form.&nbsp; Also, for the properties of the user
control, there is a property named: ToolBoxBitmap.&nbsp; Here is where you
assign an image to use as the image that will appear in the toolbox later
on.&nbsp; For best results make a BMP 16x15.&nbsp; Note it will attempt to read
(I believe the bottom left pixel, or top left I forget) to determine what color
it will use as its transparent color.&nbsp; I normally just keep a one pixel
border around the image I make and have the background color set to LIME green
or something to stay out of trouble.&nbsp; Feel free to make it what ever you
want or you can do it later.</font></small></p>
<p><small><b><font face="Verdana">Testing:</font></b></small></p>
<p><small><font face="Verdana">Ok, time to test.&nbsp; Yeah.&nbsp; Don't close
this project but I do suggest closing all design and code windows of the control
(in your playing around later you will find out why), just go up to File, then
choose Add Project.&nbsp; Choose Standard EXE.&nbsp; Click Ok.</font></small></p>
<p><small><font face="Verdana">Now over in your project explorer tree right
click on the project that just got added and choose Set As Start Up.</font></small></p>
<p><small><font face="Verdana">Rename the project to whatever you want.&nbsp;
Like WebPicTest.&nbsp; Then go to the form and rename it to something like
frmTEST.</font></small></p>
<p><small><font face="Verdana">In the toolbox you should see either the image
you made for your control, or the default generic image if you did not.&nbsp; If
you cannot tell, just mouse move over the controls listed as the bottom until
the tooltip of one reads the name of your control.&nbsp; Go ahead and click it
and add it to your form.</font></small></p>
<p><small><font face="Verdana">Presto, your control is on a form.&nbsp; Go ahead
and resize it a bit to make sure our resize code works.</font></small></p>
<p><small><font face="Verdana">Yeah it does (at least for me).&nbsp; Over in the
properties for it, check AutoResize to true.</font></small></p>
<p><small><font face="Verdana">Then in the picture property go and browse for an
image from your hard drive to test the Picture Property.</font></small></p>
<p><small><font face="Verdana">Yeah, it worked and it resized right.</font></small></p>
<p><small><font face="Verdana">Ok, now the real test.&nbsp; Remove the image
from the Picture Property.&nbsp; What we coded really is setup for us to use
either Picutre, or PictureURL but not really both, it won't crash, but just adds
a little confusion.&nbsp; Anyhow, delete the previous image from the picture
property then in the PictureURL property type: <a href="http://microsoft.com/library/homepage/images/init_windows.gif">http://microsoft.com/library/homepage/images/init_windows.gif</a>
and press enter.&nbsp; If you left the other picture there, all that would
happen is the new web downloaded image would replace it.&nbsp; Ok, time for
testing of the events and calling properties in code.&nbsp; First lets delete
what we typed in the PictureURL property.&nbsp; Then for the Picture Property go
ahead and browse and choose an image from your hard drive.&nbsp; The on the code
for this test form have it say:</font></small></p>
<p><font face="Courier New" color="#000080"><small>Option Explicit<br>
Private Sub WebPictureBox1_Click()<br>
&nbsp;&nbsp;&nbsp; Me.WebPictureBox1.PictureURL = _<br>
&nbsp;&nbsp;&nbsp; "http://microsoft.com/library/homepage/images/init_windows.gif"<br>
&nbsp;&nbsp;&nbsp; Debug.Print "Click"<br>
End Sub<br>
Private Sub WebPictureBox1_DblClick()<br>
&nbsp;&nbsp;&nbsp; Debug.Print "DblClick"<br>
End Sub<br>
Private Sub WebPictureBox1_DownloadComplete()<br>
&nbsp;&nbsp;&nbsp; Debug.Print "Download Complete"<br>
End Sub<br>
Private Sub WebPictureBox1_MouseDown(Button As Integer, _<br>
&nbsp;&nbsp;&nbsp; Shift As Integer, X As Single, Y As Single)<br>
&nbsp;&nbsp;&nbsp; Debug.Print "MouseDown"<br>
End Sub<br>
Private Sub WebPictureBox1_MouseMove(Button As Integer, _<br>
&nbsp;&nbsp;&nbsp; Shift As Integer, X As Single, Y As Single)<br>
&nbsp;&nbsp;&nbsp; Me.Caption = X &amp; " : " &amp; Y<br>
End Sub<br>
Private Sub WebPictureBox1_MouseUp(Button As Integer, _<br>
&nbsp;&nbsp;&nbsp; Shift As Integer, X As Single, Y As Single)<br>
&nbsp;&nbsp;&nbsp; Debug.Print "MouseUp"<br>
End Sub<br>
Private Sub WebPictureBox1_Resize()<br>
&nbsp;&nbsp;&nbsp; Debug.Print "Resize"<br>
End Sub</small></font></p>
<p><small><font face="Verdana">Run your test project.</font></small></p>
<p><small><font face="Verdana">Click on your control</font></small></p>
<p><small><font face="Verdana">Did it download the image.&nbsp; Did you get all
your debug.prints?</font></small></p>
<p><small><font face="Verdana">I did.&nbsp; yeah.&nbsp; </font></small></p>
<p><small><b><font face="Verdana">Code Listing Reference:</font></b></small></p>
<p><small><font face="Verdana">Ok, here at the end is a listing of all the code
for the user control so you can just copy and paste from here if you had any
problems:</font></small></p>
<hr>
<b><font FACE="Verdana" SIZE="2" COLOR="#4d8080">
<p ALIGN="LEFT">WebPictureBox (Code)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">Option Explicit</p>
<p ALIGN="LEFT">Event Click()</p>
<p ALIGN="LEFT">Event DblClick()</p>
<p ALIGN="LEFT">Event MouseDown(Button As Integer, Shift As Integer, X As
Single, Y As Single)</p>
<p ALIGN="LEFT">Event MouseMove(Button As Integer, Shift As Integer, X As
Single, Y As Single)</p>
<p ALIGN="LEFT">Event MouseUp(Button As Integer, Shift As Integer, X As Single,
Y As Single)</p>
<p ALIGN="LEFT">Event Resize()</p>
<p ALIGN="LEFT">Event DownloadComplete()</p>
<p ALIGN="LEFT">Const m_def_PictureURL = &quot;&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">Private m_PictureURL As String</p>
<p ALIGN="LEFT">Private m_privateResize As Boolean</p>
<p ALIGN="LEFT">Private Sub picBOX_Click()</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent Click</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub picBOX_DblClick()</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent DblClick</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub picBOX_MouseDown(Button As Integer, Shift As
Integer, X As Single, Y As Single)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent MouseDown(Button, Shift, X, Y)</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub picBOX_MouseMove(Button As Integer, Shift As
Integer, X As Single, Y As Single)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent MouseMove(Button, Shift, X, Y)</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub picBOX_MouseUp(Button As Integer, Shift As Integer,
X As Single, Y As Single)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent MouseUp(Button, Shift, X, Y)</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub UserControl_Initialize()</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; With UserControl</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; .picBOX.Move 0, 0, .ScaleWidth,
.ScaleHeight</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; End With</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub UserControl_Resize()</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; If m_privateResize = False Then</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; With UserControl</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
.picBOX.Move 0, 0, .ScaleWidth, .ScaleHeight</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End With</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; End If</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent Resize</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Public Property Get Appearance() As Integer</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Appearance = picBOX.Appearance</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Let Appearance(ByVal New_Appearance As Integer)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.Appearance() = New_Appearance</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;Appearance&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Get BackColor() As OLE_COLOR</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; BackColor = picBOX.BackColor</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.BackColor() = New_BackColor</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;BackColor&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Get BorderStyle() As Integer</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; BorderStyle = picBOX.BorderStyle</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Let BorderStyle(ByVal New_BorderStyle As
Integer)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.BorderStyle() = New_BorderStyle</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;BorderStyle&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Get AutoRedraw() As Boolean</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; AutoRedraw = picBOX.AutoRedraw</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.AutoRedraw() = New_AutoRedraw</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;AutoRedraw&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Get AutoSize() As Boolean</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; AutoSize = picBOX.AutoSize</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Let AutoSize(ByVal New_AutoSize As Boolean)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.AutoSize() = New_AutoSize</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;AutoSize&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Get Picture() As Picture</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Set Picture = picBOX.Picture</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Set Picture(ByVal New_Picture As Picture)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Set picBOX.Picture = New_Picture</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; If Me.AutoSize = True Then</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; With UserControl</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
m_privateResize = True</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
.Width = .picBOX.Width</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
.Height = .picBOX.Height</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
m_privateResize = False</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; End With</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; End If</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;Picture&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Private Sub UserControl_ReadProperties(PropBag As PropertyBag)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.Appearance = PropBag.ReadProperty(&quot;Appearance&quot;,
1)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.BackColor = PropBag.ReadProperty(&quot;BackColor&quot;,
&amp;H8000000F)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.BorderStyle = PropBag.ReadProperty(&quot;BorderStyle&quot;,
1)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.AutoRedraw = PropBag.ReadProperty(&quot;AutoRedraw&quot;,
False)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; picBOX.AutoSize = PropBag.ReadProperty(&quot;AutoSize&quot;,
False)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Set Picture = PropBag.ReadProperty(&quot;Picture&quot;,
Nothing)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; m_PictureURL = PropBag.ReadProperty(&quot;PictureURL&quot;,
m_def_PictureURL)</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub UserControl_WriteProperties(PropBag As PropertyBag)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;Appearance&quot;,
picBOX.Appearance, 1)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;BackColor&quot;,
picBOX.BackColor, &amp;H8000000F)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;BorderStyle&quot;,
picBOX.BorderStyle, 1)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;AutoRedraw&quot;,
picBOX.AutoRedraw, False)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;AutoSize&quot;,
picBOX.AutoSize, False)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;Picture&quot;,
Picture, Nothing)</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Call PropBag.WriteProperty(&quot;PictureURL&quot;,
m_PictureURL, m_def_PictureURL)</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Public Property Get PictureURL() As String</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PictureURL = m_PictureURL</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Public Property Let PictureURL(ByVal New_PictureURL As String)</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; m_PictureURL = New_PictureURL</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; If (New_PictureURL &lt;&gt; &quot;&quot;)
Then</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; AsyncRead
m_PictureURL, vbAsyncTypePicture, &quot;PictureURL&quot;, vbAsyncReadForceUpdate</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; End If</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; PropertyChanged &quot;PictureURL&quot;</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Property</p>
<p ALIGN="LEFT">Private Sub UserControl_InitProperties()</p>
</font></b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; m_PictureURL = m_def_PictureURL</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
<p ALIGN="LEFT">Private Sub UserControl_AsyncReadComplete(AsyncProp As
AsyncProperty)</p>
</font></b><font FACE="Verdana" SIZE="1" COLOR="#800000">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; On Error Resume Next</p>
</font><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; Select Case AsyncProp.PropertyName</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Case &quot;PictureURL&quot;</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Set Me.Picture = AsyncProp.Value</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Case Else</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; End Select</p>
<p ALIGN="LEFT">&nbsp;&nbsp;&nbsp; RaiseEvent DownloadComplete</p>
</font><b><font FACE="Verdana" SIZE="1">
<p ALIGN="LEFT">End Sub</p>
</font></b>
<hr>
<p><font face="Verdana"><small>I hope this all works out.&nbsp; I also hope you
learned something.&nbsp; At least you got a new control.&nbsp; One final thing I
would do is back in the WebPic project.&nbsp; I would open the Object Browser,
then for each of the properties of our new control, define tips to display in
the property window (the area on the bottom of the property window that tells
you what a property does) and also pick what event I would want as my default
event.&nbsp; While this option is nice and make a control more professional, it
is just too much to explain here and requires another lesson.</small></font></p>
<p><small><font face="Verdana">Feel free to email if you have any questions:</font></small></p>
<p><small><a href="mailto:lafeverc@hotmail.com"><font face="Verdana">lafeverc@hotmail.com</font></a></small></p>
<p><small><font face="Verdana">-Clint LaFever<br>
<a href="http://lafever.iscool.net">http://lafever.iscool.net</a> or </font><a href="http://vbaisc.iscool.net"><font face="Verdana">http://vbaisc.iscool.net</font></a></small></p>
<p>&nbsp;</p>

