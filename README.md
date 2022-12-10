# replywithattach
simple VBA macro to create a missed basic button in MS Outlook. This button just reply to all recipients in selected mail item with before attached files

install:

0. Open Visual Basic Editor in Outlook
1. Create new module and paste code from module1.bas to editor
2. files "olkmailread" and "olkexplorer" move to %userprofile%\AppData\Local\Microsoft\Office\ (this is ui bars in outlook)
3. Create new self certificate by open this "C:\Program Files\Microsoft Office\Office16\SELFCERT.EXE"
4. Run certmgr (win+r) and find in personal your created certificate on 3 step
5. Open your certificate, export it and install to "root" and "trustedpublisher"
6. in Visual Basic Editor sign macro with your new certificate (Tools -> Digital Signature -> select your)
7. *IMPORTANT* create a new directory %USERPROFILE%\OutlookAttachments (macro need this directory to save attachments, don't worry macro delete attachments after reply message)

P.S.
Outlook can be installed in x64 and x86 versions. They doesn't have capatability with each other. You need to make 3-6 steps again with new certificate (for each version new certificate x64 or x86)

You can create 2 versions, then create a bat file that identifies the OS version and installs the correct version.


New button:

<img width="844" alt="image" src="https://user-images.githubusercontent.com/100531769/206876973-19a27600-94ff-4d9d-923d-7f5e56d36144.png">

If u see that error, you need to signate your macro with digital signature (3-6 steps):

<img width="566" alt="Снимок экрана 2022-12-11 в 03 11 20" src="https://user-images.githubusercontent.com/100531769/206877202-a85f9df1-3227-4260-a2c9-d082057331be.png">
