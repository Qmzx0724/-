set wshshell=wscript.createobject("wscript.shell")
wshshell.AppActivate"要发送的人的名字"
for i=1 to 1000
wscript.sleep 100
wshshell.sendKeys"^v"
wshshell.sendKeys"%s"
next