<div align="center">

## Get IP Location Info


</div>

### Description

With this simple function your able to return information about an ip address. :D
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kreshnik](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kreshnik.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kreshnik-get-ip-location-info__1-73132/archive/master.zip)





### Source Code

```
Attribute VB_Name = "mLocIP"
'---------------------------------------------------------------------------------------
' Module   : mLocIP
' DateTime  : 19/12/2009 08:55PM
' Author   : Kresha7
' Mail    : kresha7@hotmail.com
' Purpose   : Gets Information about the location of an IP address
'---------------------------------------------------------------------------------------
Public Function LocateIP(IPAddr As String) As String
  Dim HTTP  As Object
  Dim StrRes As String
  Dim IP   As String, Region As String, Country As String, City As String, Latitude As String, Longitude As String, TZone As String, ISP As String, ConT As String
  Const URL = "http://www.ip2location.com/"
  Set HTTP = CreateObject("Winhttp.Winhttprequest.5.1")
  With HTTP
    .Open "POST", URL & IPAddr
    .Send
    StrRes = .ResponseText
  End With
  IP = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblICountry")(1), "</span>")(0), 3)
  Region = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblIRegion")(1), "</span>")(0), 3)
  Country = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblICity")(1), "</span>")(0), 3)
  Latitude = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblILatitude")(1), "</span>")(0), 3)
  Longitude = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblILongitude")(1), "</span>")(0), 3)
  TZone = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblITimeZone")(1), "</span>")(0), 3)
  ConT = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblINetSpeed")(1), "</span>")(0), 3)
  ISP = Mid(Split(Split(StrRes, "dgLookup__ctl2_lblIISP")(1), "</span>")(0), 3)
LocateIP = IP & vbNewLine & Region & vbNewLine & Country & vbNewLine & Latitude & vbNewLine & Longitude & vbNewLine & TZone & vbNewLine & ConT & vbNewLine & ISP
End Function
```

