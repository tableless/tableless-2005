<!-- #include file="cab.asp" -->
<%
sub fazBusca(q)

	q=trim(q)
	if q<>"" then

		%><h2>Procurando no Blog por <br><em>"<%=q%>"</em></h2><%

		buscar=split(lcase(q)," ")
		achou="<ul class=""busca"">"
		cont=0

		set docXML=CreateObject("Microsoft.XMLDOM")
		docXML.load("d:\http\tableless\web\blog.xml")

		set posts=docXML.documentElement.selectNodes("/blog/mensagem")
		for each p in posts
			tem=true
			tp=lcase(p.Text)
			for x=0 to ubound(buscar)
				if instr(tp,buscar(x))=0 then
					tem=false
				end if
			next
			if tem then 
				achou=achou & "<li><a href=""http://www.tableless.com.br/mensagem.asp?id=" & p.getAttribute("id") & """>" & p.childNodes(0).Text & " - " & RegReplace(p.childNodes(3).Text,"([0-9]+)/([0-9]+)/([0-9]+)", "$2/$1/$3") & "</a></li>"
				cont=cont+1
			end if
		next

		if cont>0 then
			if cont=1 then
				achou="<p>Um resultado encontrado.</p>" & achou & "</ul>"
			else
				achou="<p>" & cont & " resultados encontrados.</p>" & achou & "</ul>"
			end if
		else
			achou="<p>Nenhum resultado encontrado.</p>"
		end if

		Response.Write achou
		%>
		<p>Por enquanto, essa busca procura apenas no blog. Tente uma <a href="http://www.google.com.br/search?q=<%=Server.URLEncode(q)%>+site:www.tableless.com.br">busca no site inteiro</a> pelo Google.</p>
		<%

	end if

end sub

function util_RegReplace(txtin,regex,txtrpl,ignorecase,global)
	Dim objRE
	Set objRE = New RegExp
	objRE.pattern=regex
	objRE.ignorecase=ignorecase
	objRE.global=global
	util_RegReplace=objRE.Replace(txtin,txtrpl)
	Set objRE = Nothing
end function

function RegReplace(txtin,regex,txtrpl)
	RegReplace=util_RegReplace(txtin,regex,txtrpl,true,true)
end function
%>
<script runat="server" language="jscript">
function makeURL(txt){
	txt=txt+""
	txt=txt.replace(/(http:\/\/[^\n ]*)/gi,"<a href=\"$1\" title=\"$1\">¹$1¹</a>")
	Arrtxt=txt.split("¹")
	for(i=1;i<Arrtxt.length;i+=2)Arrtxt[i]=Arrtxt[i].substr(0,17)+"..."
	return (Arrtxt+"").replace(/\>,http/gi,">http").replace(/\.\.\.,/g,"...")
}
</script>


	<%fazBusca(request("q"))%>

	<!-- #include file="direita.asp" -->
	<!-- #include file="rodape.asp" -->