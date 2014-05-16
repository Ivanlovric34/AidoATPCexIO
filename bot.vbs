    option explicit
    ' #
    ' #
    ' # HELP SECTION
    ' #
    ' #
    '
    '
    ' 1, Create your cex.io account from https://cex.io ( mandatory )
    ' 2, Tranfer Bitcoins to you account
    ' 3, Under your CEX account, create a public and secret key
    ' 4, Modify the 3 vars below ( G_USERNAME, G_APIKEY, G_APIKEY_SECRET )
    ' 5, If the decimal separator in your country is not ".", modify "G_SeparateurDecimalDuSyteme" ( eg "," for France )
    ' 6, Example below show how to display your balance and place an order
    ' 7, Send 0.01 BITCOINs to "1EfjFxXXxxz7PQBKFHe7UQ7LduR64Vczf5" and use this basic code exemple to do more and more
    '
    ' Usage ( after acount creation and key and secret key settings ) :
    '
    ' 1, Save this content to a file named c:\test.vbs
    ' 2, open a Windows command prompt ( start / run / type "cmd" )
    ' 3, type "cscript c:\test.vbs"
    '
    '
    ' If you win something with this code, help me and donate some BITCOINs to the following adress :
    '
    ' 1EfjFxXXxxz7PQBKFHe7UQ7LduR64Vczf5
    '
    '
    ' Creator Remy BILLIG - billig_remy@hotmail.com - December 2013 - ( France )
    ' Special thanks to Demon ( http://demon.tw )
    ' Debug and Modified by Ivan Lovric - January 2014 (France)
    '
    ' Thank you
    ' More info : https://cex.io/api
    '
    ' #
    ' #
    ' # END OF HELP SECTION
    ' #
    ' #
     
    dim G_USERNAME, G_APIKEY, G_APIKEY_SECRET
    G_USERNAME = "xxxxxx"
    G_APIKEY = "yyyyyyyyyyyyyyyy"
    G_APIKEY_SECRET = "zzzzzzzzzzzzzzzzz"
     
     
     
    dim G_SeparateurDecimalDuSyteme
    G_SeparateurDecimalDuSyteme = ","
     
    dim G_xmlhttp
    set G_xmlhttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    G_xmlhttp.setOption 2, 13056
    G_xmlhttp.open "GET", "https://cex.io/api//ticker/GHS/BTC"
    G_xmlhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
     
     
     
    main()
    Function main()
     
    if G_USERNAME = "rbillig" then
     
    wscript.echo vbnewline
    wscript.echo " ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !"
    wscript.echo vbnewline
    wscript.echo "1, Create your cex.io account from https://cex.io ( mandatory )"
    wscript.echo "2, Under your CEX account, create a public and secret key (API dedicated)"
    wscript.echo "3, Modify the 3 vars G_USERNAME, G_APIKEY, G_APIKEY_SECRET"
    wscript.echo " (in this c:\test.vbs file)"
    wscript.echo "4, Send 0.01 BITCOINs to 1EfjFxXXxxz7PQBKFHe7UQ7LduR64Vczf5 ( mandatory )"
    wscript.echo " and use this basic exemple to do more and more"
    wscript.echo vbnewline
    wscript.echo " Please read help section from this c:\test.vbs file !!!"
    wscript.echo vbnewline
    wscript.echo " ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! ! !"
    wscript.echo vbnewline
     
    else
     
     
    '
    '
    ' Get Balance
    '
    '
     
    dim ObjetMyBalance
    Set ObjetMyBalance = GetBalance()
    DisplayJasonObject(ObjetMyBalance)

    dim theTicker
    Set theTicker = GetTicker("GHS","BTC")
    'DisplayJasonObject(theTicker)

    dim theOrderBook
    'Set theOrderBook = GetOrderBook("GHS","BTC")
    'DisplayJasonObject(theOrderBook)
    
    dim Order
   
    dim i,j,k,j0,k0
    dim ta, tb
    for i=0 to 10000     
      on error resume next
      Set theTicker = GetTicker("GHS","BTC")
      j=CDbl(replace(left(theTicker.item("ask"),10),".",G_SeparateurDecimalDuSyteme ))
      k=CDbl(replace(left(theTicker.item("bid"),10),".",G_SeparateurDecimalDuSyteme ))
      if (abs(j-j0)>0.0001) then
         if (j>j0) then
            Wscript.echo "ASK UP UP UP" & chr(7)
         else
            Wscript.echo "ASK DOWN DOWN" & chr(7)
         end if
      end if
      if (abs(k-k0)>0.0001) then
         if (k>k0) then
            Wscript.echo "BID UP UP UP" & chr(7)
         else
            Wscript.echo "BID DOWN DOWN" & chr(7)
         end if
      end if
      if (j=j0) then
         ta="="
      else 
       if (j>=j0) then
        ta="/"
       else
        ta="\"
       end if
      end if
      if (k=k0 ) then
       tb="="
      else
       if (k>k0) then
        tb="/"
       else
        tb="\"
       end if
      end if
      Wscript.echo ta & " " &j & " :: " & tb & " " &k
       j0 = j
       k0 = k
      on error resume next
      on error goto 0

      Wscript.sleep 500
    next

    wscript.echo vbnewline
    wscript.echo vbnewline
     
    wscript.echo "#"
    wscript.echo "# Balance before Order"
    wscript.echo "#"
     
    wscript.echo ObjetMyBalance.item("BTC")("available") & " BTC"
    wscript.echo ObjetMyBalance.item("GHS")("available") & " GHS"
     
    ' Carrefull regarding the number of queries per minute
    ' => wait 5 second between queries
    'wscript.sleep 5000
     
    '
    ' Place an order to buy 1 GHS at the current Rate ( if you remove the comment ... )
    '
    ' NB, if decimal separator of th system is a "," and not a "." change G_SeparateurDecimalDuSyteme at the beginning
    '
    'wscript.echo "Place Order if you remove the comment"
    'Set Order = PlaceOrder("GHS", "BTC", "buy", "1" , left(ObjetMyBalance.item("asks")(0)(0) ,8) )
    'Set Order = PlaceOrder("GHS", "BTC", "buy", "0.001" , left(theTicker.item("ask") ,10 ) )

     
    ' Carrefull regarding the number of queries per minute
    ' => wait 5 second between queries again
    'wscript.sleep 5000
     
    Set ObjetMyBalance = GetBalance()
    DisplayJasonObject(ObjetMyBalance)
    wscript.echo "#"
    wscript.echo "# Balance after Order"
    wscript.echo "#"
    wscript.echo ObjetMyBalance.item("BTC")("available") & " BTC"
    wscript.echo ObjetMyBalance.item("GHS")("available") & " GHS"
     
    end if
     
    End Function
     
     
     
     
     
     
    '''
    ''' API DEDICATED FUNCTIONS
    '''
    '''
    '''
     
    function GetOrderBook(CUR_FROM, CUR_TO )
    dim nonce, StringToEncode, signature, JsonObject
    G_xmlhttp.open "GET", "https://cex.io/api/order_book/" & CUR_FROM &"/"& CUR_TO, false
    G_xmlhttp.send ""
    Set JsonObject= New VbsJson
    'Wscript.echo G_xmlhttp.responseText
    Set GetOrderBook = JsonObject.Decode(G_xmlhttp.responseText)
    end function
     
    function GetTicker(CUR_FROM, CUR_TO )
    dim nonce, StringToEncode, signature, JsonObject
    G_xmlhttp.open "GET", "https://cex.io/api/ticker/" & CUR_FROM &"/"& CUR_TO, false
    G_xmlhttp.send ""
    Set JsonObject= New VbsJson
    'Wscript.echo G_xmlhttp.responseText
    Set GetTicker = JsonObject.Decode(G_xmlhttp.responseText)
    end function
     
     
    function GetBalance()
    dim nonce, signature, JsonObject
    nonce = Get_Nonce()
    signature = hash_sha256( nonce & G_USERNAME & G_APIKEY, G_APIKEY_SECRET )
    G_xmlhttp.open "POST", "https://cex.io/api/balance/", false
    G_xmlhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    G_xmlhttp.send "key=" & G_APIKEY & "&signature=" & signature & "&nonce=" & nonce
    Set JsonObject= New VbsJson
    'Wscript.echo G_xmlhttp.responseText
    Set GetBalance = JsonObject.Decode(G_xmlhttp.responseText)
    end function
     
     
    function PlaceOrder(CUR_FROM, CUR_TO, buy_or_sell, amount, price )
    dim nonce, StringToEncode, signature, JsonObject
     
    amount=replace(amount ,G_SeparateurDecimalDuSyteme, "." )
    price =replace(price ,G_SeparateurDecimalDuSyteme, "." )
     
    nonce = Get_Nonce()
    signature = hash_sha256( nonce & G_USERNAME & G_APIKEY, G_APIKEY_SECRET )
    G_xmlhttp.open "POST", "https://cex.io/api/place_order/" & CUR_FROM &"/"& CUR_TO, false
    G_xmlhttp.setRequestHeader "Content-type","application/x-www-form-urlencoded"
    G_xmlhttp.send "key=" & G_APIKEY & "&signature=" & signature & "&nonce=" & nonce & "&type=" & buy_or_sell & "&amount=" & replace(amount,G_SeparateurDecimalDuSyteme,".") & "&price=" & replace(price,G_SeparateurDecimalDuSyteme,".")
    Set JsonObject= New VbsJson
    'Wscript.echo G_xmlhttp.responseText
    Set PlaceOrder = JsonObject.Decode(G_xmlhttp.responseText)
    end function
     
    function DisplayJasonObject(o)
    dim i, j, k
    For Each i In o.keys
    Wscript.echo i
    if i="timestamp" then
    Wscript.echo vbtab & i & "---" & o.item(i)
    end if
    if i="bids" then
    for j=lbound(o.item(i)) to 5 'ubound(o.item(i))
    wscript.echo vbtab & o.item(i)(j)(0) & "---" & o.item(i)(j)(1)
    next
    end if
    if i="asks" then
    for j=lbound(o.item(i)) to 5 'ubound(o.item(i))
    wscript.echo vbtab & o.item(i)(j)(0) & "---" & o.item(i)(j)(1)
    next
    end if
    if i="BTC" or i="NMC" or i="GHS" then
    wscript.echo vbtab & "available" & " --- " & o.item(i)("available")
    wscript.echo vbtab & "orders " & " --- " & o.item(i)("orders")
    end if
    if i="ask" or i="bid" then
    wscript.echo vbtab & i & " --- " & left(o.item(i),10)
    end if
    Next
    end function
     
     
    Function hash_sha256( StringToHash, KeyForHash )
    'On déclare la variable servant à crypter
    Dim sha256,ObjUTF8, hmac
    Set ObjUTF8 = createobject("System.Text.UTF8Encoding")
    set sha256 = createobject("system.security.cryptography.HMACSHA256")
    ' Hachage
    sha256.Key = ObjUTF8.GetBytes_4( KeyForHash )
    hmac = sha256.ComputeHash_2( ObjUTF8.GetBytes_4( StringToHash ) )
    'Libération des ressources
    sha256.Clear()
    ' Renvoi sous format string de type hexadecimale
    dim i, a, strText
    For i = 1 To LenB(hmac )
    a = a & Right("0" & Hex(AscB(MidB(hmac , i, 1))), 2)
    Next
    hash_sha256 = LCase(a)
    End Function
     
     
    Function Get_Nonce()
    Get_Nonce = int(10*24*3600*(date() + time()))
wscript.echo Get_Nonce
    end function
     
     
     
    '''
    ''' JSon dedicated FUNCTIONS - Special thanks to Demon ( http://demon.tw )
    '''
    '''
    '''
    Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: http://demon.tw
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t
    Private Sub Class_Initialize
    Whitespace = " " & vbTab & vbCr & vbLf
    b = ChrW(8)
    f = vbFormFeed
    r = vbCr
    n = vbLf
    t = vbTab
    Set NumberRegex = New RegExp
    NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
    NumberRegex.Global = False
    NumberRegex.MultiLine = True
    NumberRegex.IgnoreCase = True
    Set StringChunk = New RegExp
    StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
    StringChunk.Global = False
    StringChunk.MultiLine = True
    StringChunk.IgnoreCase = True
    End Sub
     
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript | JSON |
    '+===================+===============+
    '| Dictionary | object |
    '+-------------------+---------------+
    '| Array | array |
    '+-------------------+---------------+
    '| String | string |
    '+-------------------+---------------+
    '| Number | number |
    '+-------------------+---------------+
    '| True | true |
    '+-------------------+---------------+
    '| False | false |
    '+-------------------+---------------+
    '| Null | null |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
    Dim buf, i, c, g
    Set buf = CreateObject("Scripting.Dictionary")
    Select Case VarType(obj)
    Case vbNull
    buf.Add buf.Count, "null"
    Case vbBoolean
    If obj Then
    buf.Add buf.Count, "true"
    Else
    buf.Add buf.Count, "false"
    End If
    Case vbInteger, vbLong, vbSingle, vbDouble
    buf.Add buf.Count, obj
    Case vbString
    buf.Add buf.Count, """"
    For i = 1 To Len(obj)
    c = Mid(obj, i, 1)
    Select Case c
    Case """" buf.Add buf.Count, "\"""
    Case "\" buf.Add buf.Count, "\\"
    Case "/" buf.Add buf.Count, "/"
    Case b buf.Add buf.Count, "\b"
    Case f buf.Add buf.Count, "\f"
    Case r buf.Add buf.Count, "\r"
    Case n buf.Add buf.Count, "\n"
    Case t buf.Add buf.Count, "\t"
    Case Else
    If AscW(c)>= 0 And AscW(c) <= 31 Then
    c = Right("0" & Hex(AscW(c)), 2)
    buf.Add buf.Count, "\u00" & c
    Else
    buf.Add buf.Count, c
    End If
    End Select
    Next
    buf.Add buf.Count, """"
    Case vbArray + vbVariant
    g = True
    buf.Add buf.Count, "["
    For Each i In obj
    If g Then g = False Else buf.Add buf.Count, ","
    buf.Add buf.Count, Encode(i)
    Next
    buf.Add buf.Count, "]"
    Case vbObject
    If TypeName(obj) = "Dictionary" Then
    g = True
    buf.Add buf.Count, "{"
    For Each i In obj
    If g Then g = False Else buf.Add buf.Count, ","
    buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
    Next
    buf.Add buf.Count, "}"
    Else
    Err.Raise 8732,,"None dictionary object"
    End If
    Case Else
    buf.Add buf.Count, """" & CStr(obj) & """"
    End Select
    Encode = Join(buf.Items, "")
    End Function
    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON | VBScript |
    '+===============+===================+
    '| object | Dictionary |
    '+---------------+-------------------+
    '| array | Array |
    '+---------------+-------------------+
    '| string | String |
    '+---------------+-------------------+
    '| number | Double |
    '+---------------+-------------------+
    '| true | True |
    '+---------------+-------------------+
    '| false | False |
    '+---------------+-------------------+
    '| null | Null |
    '+---------------+-------------------+
    Public Function Decode(ByRef str)
    Dim idx
    idx = SkipWhitespace(str, 1)
    If Mid(str, idx, 1) = "{" Then
    Set Decode = ScanOnce(str, 1)
    Else
    Decode = ScanOnce(str, 1)
    End If
    End Function
     
    Private Function ScanOnce(ByRef str, ByRef idx)
    Dim c, ms
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "{" Then
    idx = idx + 1
    Set ScanOnce = ParseObject(str, idx)
    Exit Function
    ElseIf c = "[" Then
    idx = idx + 1
    ScanOnce = ParseArray(str, idx)
    Exit Function
    ElseIf c = """" Then
    idx = idx + 1
    ScanOnce = ParseString(str, idx)
    Exit Function
    ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
    idx = idx + 4
    ScanOnce = Null
    Exit Function
    ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
    idx = idx + 4
    ScanOnce = True
    Exit Function
    ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
    idx = idx + 5
    ScanOnce = False
    Exit Function
    End If
     
    Set ms = NumberRegex.Execute(Mid(str, idx))
    If ms.Count = 1 Then
    idx = idx + ms(0).Length
    ScanOnce = CDbl(replace( ms(0), ".",G_SeparateurDecimalDuSyteme ))
    Exit Function
    End If
     
    Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function
    Private Function ParseObject(ByRef str, ByRef idx)
    Dim c, key, value
    Set ParseObject = CreateObject("Scripting.Dictionary")
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
     
    If c = "}" Then
    Exit Function
    ElseIf c <> """" Then
    Err.Raise 8732,,"Expecting property name"
    End If
    idx = idx + 1
     
    Do
    key = ParseString(str, idx)
    idx = SkipWhitespace(str, idx)
    If Mid(str, idx, 1) <> ":" Then
    Err.Raise 8732,,"Expecting : delimiter"
    End If
    idx = SkipWhitespace(str, idx + 1)
    If Mid(str, idx, 1) = "{" Then
    Set value = ScanOnce(str, idx)
    Else
    value = ScanOnce(str, idx)
    End If
    ParseObject.Add key, value
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "}" Then
    Exit Do
    ElseIf c <> "," Then
    Err.Raise 8732,,"Expecting , delimiter"
    End If
    idx = SkipWhitespace(str, idx + 1)
    c = Mid(str, idx, 1)
    If c <> """" Then
    Err.Raise 8732,,"Expecting property name"
    End If
    idx = idx + 1
    Loop
    idx = idx + 1
    End Function
     
    Private Function ParseArray(ByRef str, ByRef idx)
    Dim c, values, value
    Set values = CreateObject("Scripting.Dictionary")
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "]" Then
    ParseArray = values.Items
    Exit Function
    End If
    Do
    idx = SkipWhitespace(str, idx)
    If Mid(str, idx, 1) = "{" Then
    Set value = ScanOnce(str, idx)
    Else
    value = ScanOnce(str, idx)
    End If
    values.Add values.Count, value
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "]" Then
    Exit Do
    ElseIf c <> "," Then
    Err.Raise 8732,,"Expecting , delimiter"
    End If
    idx = idx + 1
    Loop
    idx = idx + 1
    ParseArray = values.Items
    End Function
     
    Private Function ParseString(ByRef str, ByRef idx)
    Dim chunks, content, terminator, ms, esc, char
    Set chunks = CreateObject("Scripting.Dictionary")
    Do
    Set ms = StringChunk.Execute(Mid(str, idx))
    If ms.Count = 0 Then
    Err.Raise 8732,,"Unterminated string starting"
    End If
     
    content = ms(0).Submatches(0)
    terminator = ms(0).Submatches(1)
    If Len(content)> 0 Then
    chunks.Add chunks.Count, content
    End If
     
    idx = idx + ms(0).Length
     
    If terminator = """" Then
    Exit Do
    ElseIf terminator <> "\" Then
    Err.Raise 8732,,"Invalid control character"
    End If
     
    esc = Mid(str, idx, 1)
    If esc <> "u" Then
    Select Case esc
    Case """" char = """"
    Case "\" char = "\"
    Case "/" char = "/"
    Case "b" char = b
    Case "f" char = f
    Case "n" char = n
    Case "r" char = r
    Case "t" char = t
    Case Else Err.Raise 8732,,"Invalid escape"
    End Select
    idx = idx + 1
    Else
    char = ChrW("&H" & Mid(str, idx + 1, 4))
    idx = idx + 5
    End If
    chunks.Add chunks.Count, char
    Loop
    ParseString = Join(chunks.Items, "")
    End Function
    Private Function SkipWhitespace(ByRef str, ByVal idx)
    Do While idx <= Len(str) And _
    InStr(Whitespace, Mid(str, idx, 1))> 0
    idx = idx + 1
    Loop
    SkipWhitespace = idx
    End Function
    End Class