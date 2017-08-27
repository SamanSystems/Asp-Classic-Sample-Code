<html>
<head>
<title>Calling Zarinpal web service from classic ASP</title>
<body>

<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
Dim oSOAP 
    Set oSOAP = Server.CreateObject("MSSOAP.SoapClient")
    oSOAP.ClientProperty("ServerHTTPRequest") = True
    oSOAP.mssoapinit("https://www.zarinpal.com/pg/services/WebGate/wsdl")
    
    
	Dim orderId
	Dim amount

	orderId = Request("ordertext") 
	amount = Request("amounttext") 

    Dim authority
	Dim desc
	
	desc = "سفارش شماره: " &  orderId;
    
	authority = 0

    CALL oSOAP.PaymentRequest( "YOUR MERCHENT CODE", amount, "http://www.yoursite.com/callbackpage.asp", desc, authority)
    
    IF Len(authority) = 36 THEN
		Response.Redirect( "https://www.zarinpal.com/pg/StartPay/" & authority )     
    End IF
End If
%>
<FORM method=POST name="form1">

	Order Id :
	<INPUT type="text" name="ordertext">
		
	<br>	

	Amount :
	<INPUT type="text" name="amounttext" ID="Text1">

	<br>
	<br>


	<INPUT type="submit" value="Start Pay" name="submitPay">
</form>
</body>
</html>
