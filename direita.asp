		</div> <!-- /div do #esquerda -->

		<div id="direita">
			<h3>Busca</h3>
			<form action="">
				<input type="text" class="busca" />
				<input type="image" src="imagens/bot_buscar.gif" class="botbusca" value="Buscar" />
			</form>
			<h3>Histórico</h3>
			<%for link_ano=year(now)-2000 to 3 step -1%>
			<ul>
				<%for link_mes=12 to 1 step -1
					classe=""
					if right("0" & request("mes"), 6)=right("0" & link_mes & (2000+link_ano), 6) then classe=" class=""selec"""
					if link_ano<>year(now)-2000 or link_mes<=month(now) then
					%>
					
					<li<%=classe%>><a href="default.asp?mes=<%=link_mes%><%=(2000+link_ano)%>"><%=nome_mes(link_mes)%> de <%=(2000+link_ano)%></a></li>

					<%
					end if
				next%>
			</ul>
			<%next%>

			<!--
			<ul>
				<li><a href="#">Janeiro 2003</a></li>
				<li><a href="#">Fevereiro 2003</a></li>
				<li><a href="#">Março 2003</a></li>
				<li><a href="#">Abril 2003</a></li>
				<li><a href="#">Maio 2003</a></li>
				<li><a href="#">Junho 2003</a></li>
				<li><a href="#">Julho 2003</a></li>
				<li><a href="#">Agosto 2004</a></li>
				<li><a href="#">Setembro 2003</a></li>
				<li><a href="#">Outubro 2003</a></li>
				<li><a href="#">Novembro 2003</a></li>
				<li><a href="#">Dezembro 2003</a></li>
				<li><a href="#">Janeiro 2004</a></li>
			</ul>
			-->

		</div><!-- /div do #direita -->