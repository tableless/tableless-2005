<!-- #include file="cab.asp" -->

		<%if session("cf")="1" then
		%><script>alert("Seu email foi removido com sucesso da observação desse post.")</script><%
		session("cf")=""
		end if%>
		<%if request("nome")<>"" then
            if session("tmsg")<>request("mensagem") then
                session("tmsg")=request("mensagem")
                meteComentario
            end if
		end if%>
		<%=getConteudo%>


	<h2>Comente:</h2>

	<script type="text/javascript" src="../valida.js"></script>
	<form id="comentario" method="post" onsubmit="return validaMensagem(this);">

		mensagem:<br />
		<textarea name="mensagem" rows="6" cols="36"></textarea><br />
		nome:<br />
		<input type="text" name="nome" value="<%=request.form("nome")%>" class="inputtexto" /><br />
		email:<br />
		<input type="text" name="email" value="<%=request.form("email")%>" class="inputtexto" /><br />
		url:<br />
		<input type="text" name="url" value="<%=request.form("url")%>" class="inputtexto" /><br />
		<input type="checkbox" name="watch"> Observar (avisar por email quando houverem novos comentários)
		<div class="botao">
			<input type="image" src="skins/padrao/imagens/enviar.gif" value="Enviar" alt="Enviar" />
		</div>
	</form>

	<!-- #include file="direita.asp" -->
	<!-- #include file="rodape.asp" -->