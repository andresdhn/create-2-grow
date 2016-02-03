<html>
<head>
	<title>Create 2 Grow</title>
     <style>
		@font-face {
			font-family: "Myfont";
			src: url('Resources/Fonts/Avant Garde Book Bt.ttf') format("truetype");
		}
		
		body{
			background-color:F0F0F0;
			font-family:"Myfont";
			font-size:16px;
			color:#000;	

		}
		
		a{
			color:inherit;
		}
		
		#confirmation{
			width:100%;
			height:auto;
			text-align:center;
		}
		
	
	</style>
</head>

<body>
    <div id="confirmation">
		<%
		sendUrl="http://schemas.microsoft.com/cdo/configuration/sendusing"
		smtpUrl="http://schemas.microsoft.com/cdo/configuration/smtpserver"
		
		
		' Set the mail server configuration
		Set objConfig=CreateObject("CDO.Configuration")
		objConfig.Fields.Item(sendUrl)=2 ' cdoSendUsingPort
		objConfig.Fields.Item(smtpUrl)="relay-hosting.secureserver.net"
		objConfig.Fields.Update
		
		
		' Create and send the mail
		Set objMail=CreateObject("CDO.Message")
		' Use the config object created above
		Set objMail.Configuration=objConfig
		objMail.From="info@create2grow.com.au"
		objMail.ReplyTo=request.Form("From")
		objMail.To="info@create2grow.com.au"
		objMail.Subject= request.Form("Subject")
		objMail.TextBody= request.Form("Body")
		objMail.Send
		response.Write("<p>Thank you!</p>")
		response.Write("<p>Your message has been sent successfully.</p>")
		
		%>
        <a href="index.html"> Return Home </a>
	</div>
	
</body>
</html>

