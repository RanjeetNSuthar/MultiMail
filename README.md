# MultiMail

This Application is used to automate the process of generating multiple emails at the same time with greater ease.

The user can generate multiple emails at the same time by just listing down the email id's and attachment file paths for each recipient on the spreadsheet.
The user has the option of attaching or Creating the attachments if not exist while sending the email. 
Multimail App can also be used in email marketing in which the user want to send similar message to different people just by chaning some fields like name,id etc. which it does automatically thus saving time and human resource. 
Let’s say you run a small business that needs to send PDF invoices to your customers once a month. Your billing software generates the PDF invoices, one for each of your 100 customers. Now you need to send the right invoice to the right customer. Until now, you’d have to prepare multiple emails manually composing and sending 100 individual emails, and having to attach the right PDF to each email. But now, you can automate all of this using the MultiMail .


# Modules of the software:
<ul>
<li>•	Email id,password : A user has to provide his gmail id and password for login which will be encrypted by MIME. </li>
<li>•	Subject           : Here we have a Common subject for all the emails.</li>
<li>•	Body              : Here the users can put the body content for the email which is fetched by Notebpad.(variable fields should be written inside curly braces E.g {name} ).</li>
<li>•	Attachments       : In this the user can create new attachments in MS Word or add existing attachment just by providing the respective paths in the spread sheet.  
                      (In case of new attachments user only needs to modify one word file by mentioning variable fields inside curly braces in Italic font and the System will 
                      automatically generate multiple perosonalized attachments in Document folder which is included in the project folder.) </li>
<li>•	Send mail         : By clicking send button the email will be sent. </li>
</ul>

# Design Details :
The project is totally based on python and doesn’t require a database. In terms of storage it just requires a memory to store the attachments created during the process.

	Requirements: <br />
•	Python 
•	MS Word
•	MS Excel
•	Notepad

	The Application folder contains the following files which works as templates and helps in providing the right inputs so that the exceptions can be avoided.

<li>•	Template.docx              : This is a template in form of Word file where user 
                               Can make changes and thus create the layout of the 
                               Attachments which will be customized by the 
                               Application according to the recipients details.</li>

<li>•	Book1.xlsx                 : This is the spread sheet where user can fill the cells 
                               With the values which needs to be changed in the 
                               Template attachment. </li>

<li>•	body.txt                   : This is a txt file where user can type the body of the
                               Email. </li>

<li>•	email&attachment_list.xlsx : This is the spread sheet where user will
                               fill the recipients email id and the respective 
                               Attachment paths. </li>
  </ul>


The Gmail SMTP settings do have a sending limit, which is in place to prevent spamming. You can only send a total of 500 emails per day


![image](https://user-images.githubusercontent.com/76241195/121818006-3de33200-cca2-11eb-9d72-b22df0c21b81.png)


