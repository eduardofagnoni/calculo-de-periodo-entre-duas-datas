<%
dataV = request.form("txtdata")

if dataV<>"" then
    
    data_inicial = date()
    data_da_viagem = dataV
    
    function formatData(data) 
            data = replace(data,"/",".")
            data = replace(data,"-",".")
            data = split(data,".")
            dia=data(0)
            mes=data(1)
            ano=data(2)
            data = ano+"-"+mes+"-"+dia  

        formatData = data
    end function
     
    function calcula_data(data_pass_ini, data_pass_final)            

            diferencaAno = DateDiff("yyyy",data_pass_ini,data_pass_final)
            diferencaMes = DateDiff("m",data_pass_ini,data_pass_final)
            diferencaDia = DateDiff("d",data_pass_ini,data_pass_final)
            
            dias = DatePart("d", DateSerial(DatePart("yyyy", data_pass_final) + diferencaAno, DatePart("m", data_pass_final) + diferencaMes, Day(data_pass_final)), Now)
                
        calcula_data = array(diferencaAno,diferencaMes,diferencaDia,dias)

    end function
    
    dataInicial = formatData(data_inicial)
    dataFinal = formatData(data_da_viagem)

    calculo = calcula_data(dataInicial, dataFinal)

    if calculo(0) > 1 then
        txtFalta = "Faltam "
        txtAno = " anos"
    else
        txtFalta = "Faltam "
        txtAno = " ano"
    end if

    if calculo(1) > 1 then
        txtFalta = "Faltam "
        txtMes = " meses"
    else
        txtFalta = "Faltam "
        txtMes = " m&ecirc;s"
    end if

    if calculo(2) > 1 then
        txtFalta = "Faltam "
        txtDia = " dias"
    else
        txtFalta = "Falta "
        txtDia = " dia"
    end if

    if calculo(0) < 1 then

        if calculo(1) < 1 then
            
            if calculo(2)<28 then
                mensagemMontada = txtFalta & calculo(2) & txtDia &" para sua viagem"
            else
                mensagemMontada = txtFalta & calculo(2) & txtDia &" para sua viagem <br>S&atilde;o "&calculo(2)&" dias para a viagem"
            end if

        else                
            mensagemMontada = txtFalta & calculo(1) & txtMes &" e "& calculo(3) & txtDia &" para sua viagem <br>S&atilde;o "&calculo(2)&" dias para a viagem"
        end if
        
    else
        mensagemMontada = txtFalta & calculo(1) & txtMes &" e "& calculo(3) & txtDia &" para sua viagem <br>S&atilde;o "&calculo(2)&" dias para a viagem"
    end if

    response.write("Hoje é: "&data_inicial&"<br>")    
    response.write("Data selecionada: "&dataV&"<br>")
    response.write("<hr>")
    response.write("<br>")
    response.write(mensagemMontada&"<br>")
    response.write("<br>")

    response.write("Fim do arquivo...")

end if
%>

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Calculo de data</title>
</head>
<body>

    <form method="POST" action="calculo-de-datas.asp">
        <input type="date" name="txtdata" id="txtdata">
        <button type="submit">Enviar data</button>
    </form>

</body>
</html>

