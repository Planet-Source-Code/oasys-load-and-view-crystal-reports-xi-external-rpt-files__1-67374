<div align="center">

## Load and view Crystal Reports XI external RPT files\.


</div>

### Description

Allows you to load and view any outside XI report file from your VB6 code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[OASyS](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/oasys.md)
**Level**          |Beginner
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/oasys-load-and-view-crystal-reports-xi-external-rpt-files__1-67374/archive/master.zip)





### Source Code

<p><font face="Arial" size="2">I had decided to publish the following code into
PSC since I'm having looking for it for weeks and I had not found any similar
code and explanation in whole Internet. The following code is really a
compilation of dozens of manuals, codes, forum messages and so forth. I hope it
could be usefull to you. If you agree, please vote...<br>
<br>
====================<br>
HOW TO OPEN AN EXTERNAL CRYSTAL XI REPORT IN VB6<br>
====================<br>
<br>
<br>
<b>1) You must enable the &quot;Crystal ActiveX Report Viewer Library 11.0&quot; in the
CONTROLS tab of PROJECT - COMPONENTS. A new object will appear in the VB6
TOOLBOX section (the REPORT VIEWER ocx).<br>
<br>
2) Create a new form named &quot;FRMCRYSTAL&quot; (frmCrystal.frm) and add the Crystal
object into it. The report viewer window will appear inside your form. Put the
following lines in its LOAD event:<br>
</b><br>
<font color="#0000FF"><i>CRViewer1.Top = 0<br>
CRViewer1.Left = 0<br>
CRViewer1.Height = ScaleHeight<br>
CRViewer1.Width = ScaleWidth<br>
</i></font><br>
<b>Set its &quot;VISIBLE&quot; property to FALSE and close the form.<br>
</b><br>
<br>
<br>
<b>2) To open and display an external Crystal XI report, you may use these lines
in any place of your program:<br>
I'll divide the task in 4:<br>
</b><br>
2.1) point and open the report;<br>
2.2) change the report parameters;<br>
2.3) change the report SQL query; and<br>
2.4) view the report.<br>
<br>
<br>
<b>2.1)&nbsp; Let's open the external file.The second parameter (&quot;1&quot;) means
NO-EXCLUSIVE open or TEMP-COPY open.<br>
If you change for zero (EXCLUSIVE), the report will open just once per session.<br>
<br>
</b><font color="#0000FF"><i>Dim MyApp As New CRAXDRT.Application<br>
Dim MyRpt As New CRAXDRT.Report<br>
<br>
Set MyRpt = MyApp.OpenReport(&quot;c:\windows\sample.rpt&quot;, 1)<br>
</i></font><br>
<br>
<br>
<b>2.2) Let's change some parameters. Interesting to note the FORMULAFIELD
syntax, where I can name my text-fields instead to utilize their indexes (easier
to work on than the Business Solution's Crystal manual).&nbsp; Obviously, the
name between parenteses must reflect the exact formula/text/field name. In that
example, the formulas are STRING type.<br>
</b><br>
<i><font color="#0000FF">MyRpt.ReportTitle = &quot;That's my Title!&quot;<br>
<br>
MyRpt.FormulaFields.GetItemByName(&quot;InitialDate&quot;).Text = &quot;'Initial Day: &quot; &amp;
dr1(0).Value &amp; &quot;' &quot;<br>
MyRpt.FormulaFields.GetItemByName(&quot;FinalDate&quot;).Text = &quot;'Final Day: &quot; &amp; dr1(1).Value
&amp; &quot;' &quot;<br>
</font></i><br>
<br>
<br>
<br>
<b>2.3) Change the report SQL query...<br>
</b><br>
<i><font color="#0000FF">x = 100 : y = 1000<br>
MyRpt.SQLQueryString = &quot;select * from dropouts where InitialIssue &lt;= &quot; &amp; x &amp; &quot;
and FinalIssue &gt;= &quot; &amp; y<br>
</font></i><br>
<br>
<br>
<br>
<b>2.4) Finally, log on your server and database. Since your report already
exist and has the appropriated fields and layout, you must use the same database
and login info present in the report. In my example, you must change the names
for your correct ones.<br>
</b><br>
<i><font color="#0000FF">MyRpt.Database.Tables(1).SetLogOnInfo &quot;&lt;server&gt;&quot;,
&quot;&lt;database&gt;&quot;, &quot;&lt;login&gt;&quot;, &quot;&lt;password&gt;&quot;<br>
MyRpt.Database.Verify</font></i>&nbsp;&nbsp;&nbsp;&nbsp; '<font color="#FF0000">required!!!</font><br>
<br>
<br>
<b>' And voilá !!! Show the report!!!<br>
</b><br>
<i><font color="#0000FF">frmCrystal.CRViewer1.ReportSource = MyRpt<br>
frmCrystal.Show<br>
frmCrystal.CRViewer1.ViewReport<br>
<br>
</font></i><br>
' Clear memo...<br>
<br>
<font color="#0000FF"><i>Set MyRpt = Nothing<br>
Set MyApp = Nothing<br>
</i></font><br>
&nbsp;</font></p>

