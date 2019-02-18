<%
	email = request.form("email")
	url = request.form("url")
	mensagem = request.form("mensagem")
	if email="" then email="tableless@tableless.com.br"
	Set objmail = Server.CreateObject("CDONTS.NewMail")
	objmail.from = email
	objmail.to = "tableless@tableless.com.br"
	objmail.subject = "Alguém está entrando em contato pelo Tableless"
	objmail.body = "URL: " & url & vbcrlf & mensagem
	objmail.send
	set objmail=nothing
%>
	<!-- #include file="cab.asp" -->

		<h2>Enviado com sucesso</h2>
		<p>Sua mensagem foi enviada!</p>

	<!-- #include file="direita.asp" -->
	<!-- #include file="rodape.asp" -->