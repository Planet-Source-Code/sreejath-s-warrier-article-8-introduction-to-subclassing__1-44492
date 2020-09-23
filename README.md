<div align="center">

## Article \#8 Introduction to Subclassing


</div>

### Description

The API programming series is a set of articles dealing with a common theme: API

programming in Visual Basic. In this article we will look at the concept of subclassing.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sreejath S\. Warrier](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sreejath-s-warrier.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sreejath-s-warrier-article-8-introduction-to-subclassing__1-44492/archive/master.zip)





### Source Code

<strong>Foreword</strong><br>
The API programming series is a set of articles dealing with a common theme: API
programming in Visual Basic. Though there are no hard and fast rules regarding
the content of these articles, generally one article can be expected to contain
issues related to API programming, explanation of one or more API calls with
generously commented code snippets or bug reports. Depending on the subject,
these code samples may expand to become a full-fledged application. <br>
In this article we will look at the concept of subclassing.<br>
<strong>Introduction</strong><br>
Quite often you may here experienced (and even not-so-experienced) VB developers
singing the praises of subclassing. Quite often too, you see the same developers
cursing themselves for &quot;being so foolish as to use that Devil's tool&quot; in their
applications. So is subclassing a dream come true, or a nightmare? For that
matter, what is this subclassing anyway? Let's take a look.<br>
<strong>Windows internals - The basics</strong><br>
If you choose to delve into the architecture of the Windows OS, you will come
across a term called &quot;message&quot;. You will hear such sweeping generalisations like
&quot;Windows runs on messages.&quot; &quot;If Windows is the human body, then messages form
its life blood.&quot; etc, etc. So what exactly are these messages?<br>
Put quite simply, messages are the primary means by which an application informs
the Windows (or vice versa) that some particular event has occurred and/or that
a particular action needs to be taken (which pretty much amounts to the same
thing in most cases). A message has an ID and may have one or more parameters.
The OS (or the App) uses the message ID to identify which event has occurred.
The parameters provide more info re: the event. The OS or the app then takes
appropriate measures to respond suitably to this message. For this purpose it
uses a message handler or Windows procedure (WinProc). Confused? OK, let us take
an example. Suppose the user clicks the mouse on a form in your app. This
generates a message with a unique ID and with parameters indicating the location
of the mouse click, the handle of the Window etc. Depending upon the nature of
the message Windows may pass it along to the Application or handle the message
itself. In either case the message handling function is called (Surprise!)
message handler (or a WinProc short form for Windows Procedure). It is this
function that takes appropriate action to respond to this message. Needless to
say, each and every window has a default message handler. And in this case, a
window can be a button, text box, form etc. Windows keeps track of the various
message handlers using a Class structure associated with each window handle.<br>
In the case of an App written using VB, the WinProc presents the message to the
corresponding event handler after &quot;massaging&quot; it. I.e. it alters the parameters
into a form understood by the App. The event handler then performs the actions
dictated by the code written in it. Thus, the VB almost completely masks the
inner workings and presents a friendly interface to the programmer. This is not
necessarily a bad thing. But if we need to obtain more control over our app or
provide additional functionality than is provided by the default WinProc we
cannot do so from within VB. We need to enter the shadowy realm of subclassing.
Which brings us to the topic at hand.<br>
<strong>Subclassing</strong><br>
As we saw above, each window has an associated WinProc. Subclassing refers to
that method of programming in which we insert our own WinProc between the
message sender and the default WinProc. This enables us to handle the messages
in the way we choose, rather than depend on the default message handler. Of
course, we need not handle all the messages within our WinProc. We can handle
only those that we need to exhibit modified functionality and pass the rest on
to the default message handler. This enables us to add additional functionality
where we want without duplicating the rest of the features using our code. <br>
So subclassing can be illustrated as below:</p>
<div class="Code">
 <pre><b>Message Source --&gt; Our WinProc --&gt; Default WinProc</b></pre>
</div>
<p>In this sense our WinProc acts as a front office, which handles any message
we choose in a manner chosen by us, and passes the rest to the default message
handler. <br>
Let us see a simple subclassing module:</p>
<div class="Code">
 <pre><b>Public Function WindowProc(ByVal hWnd, ByVal etc....)
' This is the WinProc we insert before the default WinProc.
'In the main App we must take control of the message handling by installing our WinProc
'as the default message handler
'For this purpose we must use the SetWindowLong API call encapsulated in the user32.dll
'Also, we must hand the control back to the default message handler after we are done (again using 'the SetWindowLong API call) or the App may crash
'SetWindowLong API call has the following syntax
'procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf OurProc)
  Select Case iMsg
   Case SOME_MESSAGE
    DoSomething 'i.e. write code to accomplish something here.
  End Select
  ' pass all messages on to VB and then return the value to windows
  WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
End Function
</b></pre>
</div>
<p>In the above code we are subclassing the message SOME_MESSAGE. Whenever this
message is encountered the code we write in the Select Case block executes
before the default handler gets to see it. All the other messages are passed
unmodified to the default WinProc.<br>
Subclassing is not limited to one level either. A window can be (and in many
cases is) subclassed multiple times. This can be illustrated as below:</p>
<div class="Code">
 <pre><b>Message Source --&gt; Our WinProc#1 --&gt; Our WinProc#2 --&gt; Our WinProc#3--&gt; Default WinProc </b></pre>
</div>
<p>At each level we can select the messages we want to subclass handle them
appropriately with our code and pass the rest on to the next level. We can even
pass on the messages that we've handled to the next level, in which case the
message handler in the next level will see the modified message only. Moreover,
we can change the order in which we respond to the message by modifying the
manner in which we pass on the message to the default WinProc.<br>
I.e. if we want our code to execute after the Default handler has handled the
message we can achieve it as shown below:</p>
<div class="Code">
 <pre><b>Public Function WindowProc(ByVal hWnd, ByVal etc....)
  Select Case iMsg
   Case SOME_MESSAGE
    DoSomething
   Case WM_PAINT
    ' Here we pass the message to the default WinProc first.
    WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
'And after the default WinProc has seen the message we handle it using our code.
    Execute_Our_Code
    Exit Function 'Here we must exit the function, since we already passed the message to the
'Default WinProc. Or the message is again passed to the Default WinProc, which might not be what
'we require
  End Select
'  pass all the remaining messages on to default WinProc unmodified and then return the value to windows
  WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
 End Function
</b></pre>
</div>
<p>OK, so that is that. &quot;This seems pretty straightforward&quot; I hear you say, &quot;Why
all the initial hoopla about subclassing being tough esoteric, etc?&quot; Well, quite
unfortunately, it isn't quite as simple as that. Some of the issues are a bit
esoteric and I'd rather wait until we've discussed some more advanced concepts
before I explain them. But we'll deal with some of them here.<br>
<br>
First of all subclassing goes to the very heart of windows and hence all the
cute error-handling features are rendered useless here. If you subclass a window
and there is an error in your code, then your app <strong>WILL </strong>crash.
And it will probably take windows with it too. A GPF is a near certainty anyway.<br>
<br>
Secondly, we cannot debug subclassing code from VB. If you try that VB <strong>
WILL </strong>crash. Of course there are ways to do this. And we will deal with
them in a later article, but don't do it directly.<br>
<br>
Thirdly, if there is an error in your subclassing code and you run it from
within the IDE, VB will enter into the break mode when it encounters the error
and will very obediently <strong>crash </strong><br>
<br>
Also, writing subclassing code is nowhere near as straightforward as programming
in VB. It is much more challenging and you have to keep a sharp eye for
interdependencies, synchronisation etc which can be a regular headache.<br>
<br>
Now that doesn't mean that you shouldn't touch subclassing with barge pole. But
it does mean that you should be very, very careful when venturing into this
area. For the rewards are high, but so are the risks. <br>
For starters, keep the following things in mind:</p>
<div class="Code">
 <pre><b>
1. Always save your project before running it. So even if an error crashes VB, you won't have to retype the entire code you wrote since the last save.
2. Do not break in subclassing code. This WILL crash VB. See rule 1
3. Double triple check your subclassing code. Remember, any error here will crash your App and may even crash Windows.
4. If you get into the deep end, be aware of the interdependency and other such issues (to be dealt in a later article)
5. <strong>Most important:</strong> Don't let some crackpot author (like me) scare you away from exploring the wonderful world of subclassing.
</b></pre>
</div>
<p>Well I guess that's it for now. In the next article we will see how we can
use subclassing to modify the system menu of a window and pick up some useful
things along the way.<br>
As always, if you have any questions, comments or criticism do feel free to mail
me.<br>
Good-bye, Good luck and happy coding!</p>
<p><b>Copyright © 2001-2003 Sreejath S. Warrier<br>
&nbsp;</b></p>

