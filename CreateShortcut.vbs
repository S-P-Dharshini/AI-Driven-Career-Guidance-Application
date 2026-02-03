Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = oWS.SpecialFolders("Desktop") & "\Grow like a Pro.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = "C:\Career guidance application\index.html"
oLink.Description = "Launch Grow like a Pro! Career Guidance Application"
oLink.IconLocation = "C:\Career guidance application\favicon.ico, 0"
oLink.Save
MsgBox "Shortcut created on your desktop!", vbInformation, "Grow like a Pro!"
<!-- Google tag (gtag.js) -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-8HBSCW05KD"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());

  gtag('config', 'G-8HBSCW05KD');
</script>