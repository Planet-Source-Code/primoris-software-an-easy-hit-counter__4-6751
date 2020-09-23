<div align="center">

## An Easy Hit Counter


</div>

### Description

This will give you the absolute basics on adding a hit counter to your web pages that reads information from a database table--not a file. This doesn't use pre-written objects, and is handled completely by the server.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Primoris Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/primoris-software.md)
**Level**          |Beginner
**User Rating**    |4.2 (25 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/primoris-software-an-easy-hit-counter__4-6751/archive/master.zip)





### Source Code

```
<H3>Adding an Easy Hit Counter</H3>
<P>This is the easiest way to add a hit counter that I know of that doesn't use
the a premade component (like Microsoft's).  This hit counter saves the hit
information to a database instead of a file.  This relies heavily on the
global.asa file.  If you're not familiar with the file, I suggest reading
up on it a bit (I'm sure there's some good tutorials on the web).  For our
purposes here, just make sure you put this code in the global.asa in the root of
your web, and if you have no global.asa file, it should be okay to create a new
one with these two subprocedures in it.  Be sure to include the
<SCRIPT> tags.</P>
<P>As for the counter, the idea is simple.  At the beginning of your
application, load the previously saved hit values from the database.  Do
this by creating a quick connection, loading the values from a recordset, and
then putting the values into application-wide variables.  So your code
might look like the following:</P>
<P><FONT color="#0000FF">Sub </FONT>Application_OnStart</P>
<BLOCKQUOTE>
 <P><FONT color="#0000FF">Dim</FONT> adoApp, adoSessConn <BR>
 <FONT color="#0000FF">Set</FONT> adoSessConn =
 Server.CreateObject("ADODB.Connection") <BR>
 <FONT color="#0000FF">Set</FONT> adoApp =
 Server.CreateObject("ADODB.Recordset") <BR>
 adoSessConn.Open "DSN"<BR>
 adoApp.Open "hits", adoSessConn, 2, 2, 2 <BR>
 adoApp.MoveFirst <BR>
 Application.Lock<BR>
 Application("Home_Hits")= adoApp("homehits") <BR>
 Application.Unlock<BR>
 adoApp.Close <BR>
 adoAppConn.Close <BR>
 <FONT color="#0000FF">Set</FONT> adoApp = <FONT color="#0000FF">Nothing</FONT>
 <BR>
 <FONT color="#0000FF">Set </FONT>adoAppConn = <FONT color="#0000FF">Nothing</FONT>
 </P>
</BLOCKQUOTE>
<P><FONT color="#0000FF">End Sub</FONT></P>
<P>Now you've got one application variable loaded with the value from the last
time it was saved--which, if you coded it right, will be the number of hits that
page had received at the time the application ended.  On our web site we
keep track of about 30 pages this way, though only one is shown here.</P>
<P>Now every time a page is hit (or in this case we'll keep track of when it's
refreshed, too) you will have a chunk of code at the beginning to increment that
variable.  Notice we do <I>not</I> have code to save this incremented value
on every page.  Why bog the server down when we only need it saved every so
often?  You can put code to write the values to a table whereever you want,
naturally.  Our system has been running for quite some time like this,
however.</P>
<P>So now for the chunk of code on each page.  It should look somewhat like
the following:</P>
<P><SPAN style="background-color: #FFFF00">&lt;%</SPAN></P>
<BLOCKQUOTE>
 <P>Application.Lock<BR>
 Application("Home_Hits") = Application("Home_Hits") + 1<BR>
 Application.Unlock</P>
</BLOCKQUOTE>
<P><SPAN style="background-color: #FFFF00">%&gt;</SPAN></P>
<P>Or for the adventurous, you can change the language of the page to JavaScript
and use the trusty 'Application("Home_Hits") += 1;'.  But you get
the point.  Now the only thing left to do is save the values to your
database when the application ends.  This would look like:</P>
<P><FONT color="#0000FF">Sub </FONT>Application_OnEnd</P>
<BLOCKQUOTE>
 <P><FONT color="#0000FF">Dim</FONT> adoApp, adoSessConn <BR>
 <FONT color="#0000FF">Set</FONT> adoSessConn =
 Server.CreateObject("ADODB.Connection") <BR>
 <FONT color="#0000FF">Set</FONT> adoApp =
 Server.CreateObject("ADODB.Recordset") <BR>
 adoSessConn.Open "DSN"<BR>
 adoApp.Open "hits", adoSessConn, 2, 2, 2 <BR>
 adoApp.MoveFirst <BR>
 adoApp("homehits") = Application("Home_Hits")<BR>
 adoApp.Update<BR>
 adoApp.Close <BR>
 adoAppConn.Close <BR>
 <FONT color="#0000FF">Set</FONT> adoApp = <FONT color="#0000FF">Nothing</FONT>
 <BR>
 <FONT color="#0000FF">Set </FONT>adoAppConn = <FONT color="#0000FF">Nothing</FONT>
 </P>
</BLOCKQUOTE>
<P><FONT color="#0000FF">End Sub</FONT></P>
<P>It's as easy as that!  Yes, I know you can use a simple connection
object and execute statement, but we're going for uniformity here, and I try to
avoid using SQL statements except for selecting queries to be on the safe
side.  Notice that we always locked the application before updating
it--this is standard practice for obvious reasons.  It would by no means be
disastrous not to lock this on such a simple procedure, but you might as well
get into the habit.</P>
```

