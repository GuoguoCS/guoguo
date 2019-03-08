<%
startime=timer()
set conn=server.CreateObject("adodb.connection")
DBPath = Server.MapPath("#2008bictsrxh.mdb")
conn.open "provider=microsoft.jet.oledb.4.0; data source="&DBpath

function CloseDB

Conn.Close
set Conn=Nothing

End Function

const ProgramCName="日记"
const ProgramEName="鱼人码头"
const ProgramVer="fish"
%>