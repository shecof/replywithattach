# replywithattach
simple VBA macro to create a missed basic button in MS Outlook. This button just reply to all recipients in selected mail item with before attached files

install:
1. file "VbaProject" move to %userprofile%\AppData\Roaming\Microsoft\Outlook\ (this is macros)
2. files "olkmailread" and "olkexplorer" move to %userprofile%\AppData\Local\Microsoft\Office\ (this is ui bars in outlook)
3. Create new self certificate by open this "C:\Program Files\Microsoft Office\Office16\SELFCERT.EXE"
4. Run certmgr (win+r) and find in personal your created certificate on 3 step
5. Open your certificate, export it and install to "root" and "trustedpublisher"
6. Open Outlook, open Visual Basic Editor and sign macro with your new certificate (Tools -> Digital Signature -> select your)


<img width="844" alt="image" src="https://user-images.githubusercontent.com/100531769/206876973-19a27600-94ff-4d9d-923d-7f5e56d36144.png">

