


<!-- #include virtual="/kjasp/func/DB.fun" -->
<%
 SECTM = Request("SECTM")
 'TNUM = Request("TNUM")
''Lesson = Request("Lesson")
''No =  Request("No")
''Name = Request("Name")
  'SECNUM= CLng(SECTM)+CLng(TNUM)
  'crsTM =Trim(CStr(SECNUM))&".html"
    'crsTM ="math/"&Trim(CStr(rs("�D�ؽX")))&".html"
    ''crsTM ="1010104.html" & '�ƾǹϧ��ɶ��b�P�ؿ��~�|���
   crsTM ="/Hmath-1/TBKIN/"&Trim(SECTM)
   ' Response.Write crsTM
 
%>

<HTML>
 <BODY BgColor=White Background="B01.jpg">
'<BODY BgColor=White >
<!--<H2>�Ҹլ�� <HR></H2>
<FORM Action=ScoreKac-1c.asp Method=POST>
<FORM >   </FORM>-->
  <% =HMCLD(crsTM) %> 
<HR>
</BODY>
   <script>
    document.onselectionchange=__OnSelectionChange;
       var running=false;
     function __OnSelectionChange()
       { 
       if (running==true) return;
          running=true;
       document.selection.empty();
       running=false;       
        }
  </script>
</HTML>

 <% '���հƵ{��2 
  FUNCTION HMCLD(rsTM) 
   Set fs = Server.CreateObject("Scripting.FileSystemObject")
     File = Server.MapPath(rsTM)
  Set txtf = fs.OpenTextFile( File )
If Not txtf.atEndOfStream Then	' ���T�w�٨S����F��������m
    Content = txtf.ReadAll	' Ū������ɮת����
   '' Lines = Replace(Content, vbCrLf, "<BR>" )
    ''Response.Write Lines
   ' Response.Write Content
End If

  HMCLD=Content
END FUNCTION

%>
   
<% '���հƵ{��1
   FUNCTION HMCOD(rsTM) 
    '����WORD�إ�
    gher = ""
     ' mypos1 = 1
     ' mypos2 = 1
      mysear1 = "{"
      mysear2 = "}"      
     mytext=rsTM
     mylen = Len(mytext)
    For j= 0 TO mylen
      mypos1 = InStr(1, mytext, mysear1, 1)  
      mypos2 = InStr(1, mytext, mysear2, 1)  
     If mypos1 <> 0 Then
          textf = Mid(mytext, 1, mypos1-1)
          textm = Mid(mytext,mypos1+1, mypos2-1-mypos1)
          textb = Mid(mytext,mypos2+1, mylen)
        gher=gher+textf
          'RESPONSE.WRITE textf
          If textm <> "" Then
               textmf = Rs(textm)
              gher=gher+textmf
              'RESPONSE.WRITE textmf
           End If
             textm = ""
             mytext = Textb
               mylen = Len(mytext)
       Else
          Textb = mytext
             mylen = Len(Textb)
         gher=gher+textb
           'RESPONSE.WRITE textb
          Exit For
      End If     
   Next
  HMCOD=gher
END FUNCTION
%>      
    

    
 


















































