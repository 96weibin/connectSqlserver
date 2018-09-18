<%@ language="vbscript" codepage="65001"%>
<%
set Conn = server.createObject("adodb.connection")
Conn.connectionString = "dirver={sql server}; server=iZ2ze2ew3wajpZ\WEIBINSERVER;uid=dev;pwd=dev;dataBase=dev"
Conn.open
response.write(Conn.state)
Response.write("世界真美好")
Conn.close
%>