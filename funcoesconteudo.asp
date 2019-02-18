<%

paginacao=""

XoffSet=Request.QueryString("p")
if XoffSet="" then XoffSet="-1"
XoffSet=CInt(XoffSet)
arquivo=Request.QueryString("arq")
mesPrimeiroPost="012003"

function doisDigitos(numero)
	doisDigitos=right("00" & numero, 2)
end function

function pegaMes(data)
	pegaMes=left(data, 2)
end function

function pegaAno(data)
	pegaAno=right(data, 4)
end function

function geraBlog()

	paginacao="<h2>Arquivo:</h2> <ul>"
	set docXML=CreateObject("Microsoft.XMLDOM")
	docXML.load("d:\http\tableless\web\blog.xml")

	set docXML2=CreateObject("Microsoft.XMLDOM")
	docXML2.loadXML("<blog></blog>")
 
	primeiro=0
	ultimo=9
	jah_primeiro=false

	esteMes=pegaMes(arquivo)
	esteAno=pegaAno(arquivo)

	set mensagens=docXML.getElementsByTagName("mensagem")
	codigo=mensagens.length
	mesUltimoPost=""
	for each L in mensagens
		codigo=codigo-1
		dataPost=L.childNodes(3).text
		if mesUltimoPost="" then mesUltimoPost=dataPost
		if right(dataPost, 7)=esteMes&"/"&esteAno then
			if not jah_primeiro then 
				primeiro=mensagens.length - codigo - 1
				jah_primeiro=true
			end if
			ultimo=mensagens.length - codigo - 1
		end if
	next

	set datas=docXML.getElementsByTagName("data")
	
	for each i in datas
		i.text=transformaData(i.text)
	next

	'if XoffSet=-1 then
		for x=primeiro to ultimo '0 to 9
			if x<mensagens.length then
				docXML2.childNodes.item(0).appendChild(mensagens.item(x))
			end if
		next
	'else
	'	for x=10+(XoffSet*10) to 1+(XoffSet*10) step -1
	'		if x<=mensagens.length then
	'			docXML2.childNodes.item(0).appendChild(mensagens.item(mensagens.length-x))
	'		end if
	'	next
	'end if

	'if mensagens.length>10 then
	'for x=Int((mensagens.length-1)/10) to 0 step -1
	'	paginacao=paginacao & " <li><a href=""default.asp?p=" & x & """>Arquivo " & x & "</a></li>"
	'next
	'else
	'	paginacao=""
	'end if

	for ano=CInt(CDbl(pegaAno(mesUltimoPost))-CDbl(pegaAno(mesPrimeiroPost))) to 0 step -1
		arqAno=( CDbl(pegaAno(mesPrimeiroPost))+ano )
		pMes=1
		uMes=12
		if arqAno=CDbl(pegaAno(mesUltimoPost)) then
			uMes=CInt(pegaMes(right(mesUltimoPost, 7)))
		elseif arqAno=CDbl(pegaAno(mesPrimeiroPost)) then
			pMes=CInt(pegaMes(mesPrimeiroPost))
		end if
		for mes=uMes to pMes step -1
			paginacao=paginacao&"<li><a href=""default.asp?arq="&doisDigitos(mes)&""&arqAno&""">"&meses(mes)&" de "&arqAno&"</a></li>"
		next
	next
	paginacao=paginacao&"</ul>"
    
	set docXSL=CreateObject("Microsoft.XMLDOM")
	docXSL.load("d:\http\tableless\web\blog.xsl")

	html=docXML2.transformNode(docXSL)

	html=replace(replace(html,"[","<"),"]",">")

	geraBlog=html

end function

dim meses(12)
meses(1)="Janeiro"
meses(2)="Fevereiro"
meses(3)="Março"
meses(4)="Abril"
meses(5)="Maio"
meses(6)="Junho"
meses(7)="Julho"
meses(8)="Agosto"
meses(9)="Setembro"
meses(10)="Outubro"
meses(11)="Novembro"
meses(12)="Dezembro"


function transformaData(d)
	'd=CDate(d)
	'transformaData= Day(d) & " de " & meses(Month(d)) & " de " & Year(d)
	if isdate(d) then
		dia=left(d, instr(d, "/")-1)
		mes=mid(d, len(dia)+2, instrrev(d, "/")-(len(dia)+2))
		mes=CInt(mes)
		ano=right(d, len(d)-instrrev(d, "/"))
		ano=CLng(ano)

		if mes>12 then 
			inv=mes
			mes=dia
			dia=inv
		end if
		transformaData= dia & " de " & meses(mes) & " de " & ano
	else
		transformaData= "?? de ?? de ????"
	end if
end function

%>
