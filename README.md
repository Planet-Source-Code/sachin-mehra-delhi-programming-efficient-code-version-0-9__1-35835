<div align="center">

## Programming Efficient Code Version 0\.9


</div>

### Description

Our programs spend a lot of time in loops, and unfortunately loose a lot of their performance in them as well.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sachin Mehra \(Delhi\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sachin-mehra-delhi.md)
**Level**          |Intermediate
**User Rating**    |4.3 (43 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sachin-mehra-delhi-programming-efficient-code-version-0-9__1-35835/archive/master.zip)





### Source Code

<html>
<head>
<title>Programming Efficient Code</title>
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Courier;
	panose-1:2 7 4 9 2 2 5 2 4 4;
	mso-font-alt:"Courier New";
	mso-font-charset:0;
	mso-generic-font-family:modern;
	mso-font-format:other;
	mso-font-pitch:fixed;
	mso-font-signature:3 0 0 0 1 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;
	text-underline:single;}
code
	{font-family:Courier;
	mso-ascii-font-family:Courier;
	mso-fareast-font-family:"Times New Roman";
	mso-hansi-font-family:Courier;
	mso-bidi-font-family:"Courier New";}
pre
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Courier New";
	mso-fareast-font-family:"Times New Roman";}
span.SpellE
	{mso-style-name:"";
	mso-spl-e:yes;}
span.GramE
	{mso-style-name:"";
	mso-gram-e:yes;}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-US link=blue vlink=purple style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoNormal><b><span style='font-size:16.0pt;font-family:Arial'>Programming
Efficient Code Version 0.9</span></b></p>
<p class=MsoNormal><b><span style='font-size:11.0pt;font-family:Arial'>Focusing
Loops <o:p></o:p></span></b></p>
<p class=MsoNormal><span style='font-size:9.0pt;font-family:Arial'>Article by </span><i><span
style='font-size:10.0pt;font-family:Arial'>Sachin Mehra (<a
href="mailto:sachinweb@hotmail.com">sachinweb@hotmail.com</a>)<o:p></o:p></span></i></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>One of the
worst things that a programmer can assume is that the compiler and middleware
will do the optimizations for you!&nbsp; Most applications are being targeted
for 50-200 concurrent users, which is why we need to constantly be worrying
about the performance of our code. </span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>Say you
have a search component (which will be on everybody’s desktop) which takes 3-5
seconds to load; and consequently ties down the database.&nbsp; Imagine what
will happen when 200 people try to load this at the same time.&nbsp; Simple
math would suggest (say using an average of 4 seconds): (4 x 200) / 60 = 13 –
THAT’S 13 MINUTES!&nbsp; And actually, when dealing with situations of high
contention, you cannot assume 100% efficiency and could be realistically
dealing with something in the range of 20-25 minutes of processing time
required.&nbsp; </span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>There is no
‘silver bullet’ to making fast and efficient code.&nbsp; Middleware will not
solve the problem for you, databases will not solve the problem for you, it is
up to you as a computational process engineer (how’s that for a title?) to
understand and deal with the underlying inefficiencies in the software you design.&nbsp;
There are many things which you need to consider, and you need to think your
logic out carefully.&nbsp; One thing you should always be asking yourself is
“could this be done better?<span class=GramE>”.</span></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;In
programming, we find ourselves in loops a lot. &nbsp;In VB, ASP and COM, we
especially find ourselves looping through </span><span style='font-size:10.0pt;
font-family:"Courier New"'>Collection</span><span style='font-size:10.0pt;
font-family:Arial'> objects an awful lot.&nbsp; This is one of the particular
areas where many of us need some improvement.&nbsp; </span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>Our
programs spend a lot of time in loops, and unfortunately loose a lot of their
performance in them as well.&nbsp; I will try to cover a few pointers which may
help you in certain situations save some unnecessary computational cycles off
you’re code.</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>People tend
to think from beginning to end, and they tend to program in this forward
lineage as well.&nbsp; But this can often be inefficient. &nbsp;Sometimes the
computer can find its way from the end to the beginning much faster.&nbsp; </span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>Consider
the code in Example 1.</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;font-family:Arial'>Example
1</span></b><span style='font-size:10.0pt;font-family:Arial'>:</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>&nbsp;<span class=SpellE>i</span> =0</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span class=GramE><span
style='font-size:9.0pt;font-family:Arial'>While <span
style='mso-spacerun:yes'> </span><span class=SpellE>i</span></span></span><span
style='font-size:9.0pt;font-family:Arial'> &lt;&gt; <span class=SpellE>ubound</span>(<span
class=SpellE>arrayName</span>)</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<span class=SpellE><span class=GramE>doSomething</span></span><o:p></o:p></span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'><span style='mso-tab-count:1'>                </span><span
style='mso-spacerun:yes'>  </span><span class=SpellE>i</span> = <span
class=SpellE>i</span> + 1<o:p></o:p></span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>Wend</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>This is a
fairly straight-forward while-loop to iterate that iterates through an entire array
collection to do something.&nbsp; But consider Example 2:</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><b><span style='font-size:10.0pt;font-family:Arial'>Example
2</span></b><span style='font-size:10.0pt;font-family:Arial'>:</span></p>
<p class=MsoNormal style='background:#E6E6E6'><span style='font-size:10.0pt;
font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>&nbsp;<span class=SpellE>i</span> = <span class=SpellE><span
class=GramE>ubound</span></span><span class=GramE>(</span><span class=SpellE>arrayName</span>)</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span class=GramE><span
style='font-size:9.0pt;font-family:Arial'>While<span style='mso-spacerun:yes'> 
</span><span class=SpellE>i</span></span></span><span style='font-size:9.0pt;
font-family:Arial'> &lt;&gt; 0</span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<span class=SpellE><span class=GramE>doSomething</span></span><o:p></o:p></span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'><span style='mso-tab-count:1'>                </span><span
style='mso-spacerun:yes'>  </span><span class=SpellE>i</span> = <span
class=SpellE>i</span> - 1<o:p></o:p></span></p>
<p class=MsoNormal style='background:#E0E0E0'><span style='font-size:9.0pt;
font-family:Arial'>Wend</span></p>
<p class=MsoNormal style='background:#E6E6E6'><span style='font-size:10.0pt;
font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>This example
is many times more efficient than Example 1.&nbsp; In example one, we are
making a call to </span><span class=SpellE><span class=GramE><i
style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;font-family:
Arial'>ubound</span></i></span></span><span class=GramE><i style='mso-bidi-font-style:
normal'><span style='font-size:9.0pt;font-family:Arial'>(</span></i></span><span
class=SpellE><i style='mso-bidi-font-style:normal'><span style='font-size:9.0pt;
font-family:Arial'>arrayName</span></i></span><i style='mso-bidi-font-style:
normal'><span style='font-size:9.0pt;font-family:Arial'>)</span></i><span
style='font-size:10.0pt;font-family:Arial'> for every iteration through the
loop which is unnecessary, and also we are doing a direct X AND comparison to
determine if the loop should continue which is also more efficient.&nbsp; By
looping backwards through the </span><span style='font-size:10.0pt;font-family:
"Courier New"'>Array</span><span style='font-size:10.0pt;font-family:Arial'> we
manage to increase processing efficiency but 50% or more! &nbsp;</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><b><span style='font-family:Arial'>Final Words</span></b></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>Programming
is all about problem solving.&nbsp; And as with other kinds of problem solving,
there are always many different ways to solve the problem.&nbsp; However, some
ways are more certainly better than others.&nbsp; You should take the time to
understand how the underlying components you are using actually work, and why
they work <span class=GramE>they</span> way they do.</span></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><o:p>Version 1.0 of this article shall be comming soon, I need my friend Romi's help to work out and try a few application tests before comming up with the final version. You may also send in your inputs to me.</o:p></p>
<p class=MsoNormal><span style='font-size:10.0pt;font-family:Arial'>&nbsp;</span></p>
<p class=MsoNormal><i><span style='font-size:10.0pt;font-family:Arial'>Article <span
class=GramE>By</span> Sachin Mehra (sachinweb@hotmail.com)&nbsp; </span></i></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
</div>
</body>
</html>

