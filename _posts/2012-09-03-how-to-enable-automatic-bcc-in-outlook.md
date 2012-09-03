---
layout: post
title: "How to Enable Automatic BCC in Outlook"
category: 
tags: [ms office]
---
{% include JB/setup %}

Because I'm practicting GTD method in my work and life, getting my email organized is very important.


In David Allen's famous book *Getting Things Done*, he suggests readers to use **@Waiting** folder to track to delegated tasks. For instance, the emails I sent to other people and need their reply.


Normally, I would add myself into BCC lists. Thus I could directly drag the outgoing mail into **@Waiting** folder, without bothering to drag it from **Sent** box to **@Waiting** again.


Recently I am working on a project in which people are remotely collaborated , with jeglag. So almost each emails I sent need to be archived , or tracked. I strongly need outlook to BCC emails I sent.


However, Outlook2003 does not support the auto-BCC function itself, without installation of 3pp plugin.( I don't whether later Outlook has enabled to do this).And the reason why I'm not intersted in 3pp solution, is that it's company laptop and there are some security concerns.

Now , seems my only solution is to turn to VBA macro.


OK, here we come. 


**Note: I'm using Outlook2003, maybe your menu changed in later outlook.**

###1. Enable Macro in Outlook# 
* Open Outlook **Tools->Macro->Security**

* Select **Security Level** to *Medim*

* Insert snapshot 1 here.

###2. Add VBA code in ThisOutlookSession Module###

* Open VBA Editor by ...

* Insert snapshot 2 here.

* Paste below code into Editor ( replace someone@somewhere.com with your email)

>Private Sub Application_ItemSend(ByVal Item As Object, _Cancel As Boolean)
    Dim objRecip As Recipient
    Dim strMsg As String
    Dim res As Integer
    Dim strBcc As String
    On Error Resume Next

    ' #### USER OPTIONS ####
    ' address for Bcc -- must be SMTP address or resolvable
    ' to a name in the address book
    strBcc = "someone@somewhere.com"

    Set objRecip = Item.Recipients.Add(strBcc)
    objRecip.Type = olBCC
    If Not objRecip.Resolve Then
        strMsg = "Could not resolve the Bcc recipient. " & _
                 "Do you want still to send the message?"
        res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, _
                "Could Not Resolve Bcc Recipient")
        If res = vbNo Then
            Cancel = True
        End If
    End If

    Set objRecip = Nothing
    End Sub

* Save and exit VBA Editor.

###Here you get it###
###Send one test email Now!###
 
