<%@ language="VBScript" %>

<script language="VBScript">

bytecount = Request.TotalBytes
bytes = Request.BinaryRead(bytecount)

Set stream = Server.CreateObject("ADODB.Stream")
    stream.Type = 1 'adTypeBinary              
    stream.Open()                                   
        stream.Write(bytes)
        stream.Position = 0                             
        stream.Type = 2 'adTypeText                
        stream.Charset = "utf-8"                      
        s = stream.ReadText() 'here is your json as a string                
    stream.Close()
Set stream = nothing

Response.write(s)

</script>