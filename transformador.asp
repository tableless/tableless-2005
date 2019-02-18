<%
Session.LCID=1046

function transforma(arquivoxml,arquivoxsl) 

	'Cria os objetos necessários
	set xsl=create_processor(arquivoxsl)
	set xmldoc = Server.CreateObject("microsoft.xmldom") 

	'Carrega o XML
	if not xmldoc.Load(arquivoxml) then
	 Response.Write("ERRO XML:"+xmldoc.parseError)
	end if

	mes=CLng("0" & request("mes"))
	if mes=0 then
		mes=month(now) & year(now)
	end if
	mes=right("0" & mes, 6)

	'Executa a transformação
	xsl.input = xmldoc
	xsl.addParameter "Mes", mes
	xsl.addParameter "Random", CInt(rnd()*1000)
	xsl.transform()

	transforma=replace(xsl.output,"<?xml version=""1.0"" encoding=""UTF-16""?>","")

	transforma=replace(replace(transforma, "[", "<"), "]", ">")


end function 

function create_processor (fic_xsl) 

	Dim xslt '(New Msxml2.XSLTemplate) 
	Dim xslDoc '(New Msxml2.FreeThreadedDOMDocument.3.0) 
	Dim xmlDoc '(New Msxml2.DOMDocument) 
	Dim xslProc '(IXSLProcessor) 

	'creation d'un freethreaded à partir d'un xsl 
	'le template xsl est ensuite créé à partir de cette référenc 
	set xsldoc = Server.CreateObject("MSXML2.FreeThreadedDOMDocument.3.0") 
	set xslt = Server.CreateObject("MSXML2.XSLTemplate") 

	xslDoc.async = False 
	if not xslDoc.Load (fic_xsl) then
	 Response.Write("ERRO XSL:"+fic_xsl.parseError)
	end if
	Set xslt.stylesheet = xslDoc 
	Set xslProc = xslt.createProcessor() 
	set create_processor = xslProc 

end function 
%>

<%
dim nome_mes(12)
nome_mes(1)="Janeiro"
nome_mes(2)="Fevereiro"
nome_mes(3)="Março"
nome_mes(4)="Abril"
nome_mes(5)="Maio"
nome_mes(6)="Junho"
nome_mes(7)="Julho"
nome_mes(8)="Agosto"
nome_mes(9)="Setembro"
nome_mes(10)="Outubro"
nome_mes(11)="Novembro"
nome_mes(12)="Dezembro"
%>

<%
function transformaData(d)
	sd=split(d, "/")
	mespost=CInt(sd(0))
	if mespost<=12 then mespost=nome_mes(mespost)
	transformaData=sd(1) &  " de " & mespost & " de " & sd(2)
end function

function getConteudo
	id=request.QueryString("id")
	set docXML=CreateObject("Microsoft.XMLDOM")
	docXML.load("d:\http\tableless\web\blog.xml")
	set datas=docXML.getElementsByTagName("data")
	for each i in datas
		i.text=transformaData(i.text)
	next
	set posts=docXML.documentElement.selectNodes("/blog/mensagem[@id=" & id & "]")
	html=""
	for each p in posts
		set xsl=create_processor(server.mappath("blog.xsl"))
		xsl.input=p
		xsl.transform()
		html=html & xsl.output
		html=replace(replace(html,"[","<"),"]",">")
		set coms=p.getElementsByTagName("comentario")
		if coms.length>0 then
			html=html & "<h2>Comentários:</h2>"
		end if
		for each c in coms
			html=html & "<div class=""comentario""><h3>" & c.getAttribute("nome") & ""
			if c.getAttribute("email")<>"" then 
				html=html & " [<a href=""mailto:" & c.getAttribute("email") & """>email</a>]"
			end if
			if c.getAttribute("url")<>"" then 
				html=html & " [<a href=""http://" & c.getAttribute("url") & """>site</a>]"
			end if
			html=html & "</h3><p>" & makeURL(Replace(c.text,vbcrlf,"<br />")) & "</p></div>"
		next
	next
	getConteudo=html
end function

sub meteComentario

	id=request("id")
	nome=request("nome")
	email=request("email")
	url=request("url")
	watch=request("watch")
	mensagem=request("mensagem")
	set docXML=CreateObject("Microsoft.XMLDOM")
	docXML.load("d:\http\tableless\web\blog.xml")

	set comment=docXML.createElement("comentario")
	comment.setAttribute "nome",Server.HTMLEncode(nome)
	comment.setAttribute "email",Server.HTMLEncode(email)
	comment.setAttribute "url",Server.HTMLEncode(url)
	if watch="on" then comment.setAttribute "watch",int(rnd*1000000+rnd*1000)
	comment.text=Server.HTMLEncode(mensagem)

	set posts=docXML.documentElement.selectNodes("/blog/mensagem[@id=" & id & "]")
	for each p in posts
		set watches=p.childNodes
		for each w in watches
			if w.getAttribute("watch")<>"" then
				sendMail "Novo post nos comentários do tableless","O post no Tableless que você pediu para observar: <http://www.tableless.com.br/mensagem.asp?id=" & id & "> recebeu um novo comentário: " & vbcrlf & vbcrlf & "Nome: " & nome & vbcrlf & "Email: " & email & vbcrlf & "Site: http://" & url & vbcrlf & vbcrlf & "== Mensagem: ==" & vbcrlf & mensagem & vbcrlf &  vbcrlf & "===============" & vbcrlf & "Para parar de observar esse post clique em: <http://www.tableless.com.br/rw.asp?w=" & w.getAttribute("watch") & "&p=" & id & ">",w.getAttribute("email")
			end if
		next
		p.appendChild comment
	next

	docXML.save("d:\http\tableless\web\blog.xml")

	sendMail "Novo post nos comentários do tableless","<http://www.tableless.com.br/mensagem.asp?id=" & id & ">" & vbcrlf & vbcrlf & "Nome: " & nome & vbcrlf & "Email: " & email & vbcrlf & "Site: http://" & url & vbcrlf & vbcrlf & "== Mensagem: ==" & vbcrlf & mensagem,"tableless@tableless.com.br"

end sub


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

sub sendMail(tit,body,para)
	set mail=createobject("cdonts.newmail")
	mail.from="robo@tableless.com.br"
	mail.to=para
	mail.body=body & vbcrlf & vbcrlf & "=======" & vbcrlf & "Enviado automaticamente pelo robô tableless. Não responda, o robô não lê e-mails."
	mail.subject=tit
	mail.send
	set mail=nothing
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