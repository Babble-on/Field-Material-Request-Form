# Field-Material-Request-Form
The following is an application written for work to allow field employees to request construction parts from the field and have those parts ready when they reach a storeroom.

The GUI is written in WPF and was developed in visual studio 2015.
All of the coding is done in powershell.

On launch the GUI opens with 3 tabs.  "Find Parts to be Ordered" , "Preview Order", and "Send Order"


FIND PARTS TO BE ORDERED:
The "find parts to be ordered" tab contains 3 buttons and a drop down menu selector.   The drop down menu data source is mapped to the "Field.xml".  For the purposes od this upload I placed a dummy "field.xml" in the data directory as to not compromise my companies information.   Upon selection in the drop down menu the data grid just below is populated with all of the available parts of that data type, listing the Item number and Item Description.

By highlighting an item in the data grid and pressing the select button, the text box above "Selected Item Description" is populated with the selected items description and the text box abover "UOM" is populated with the selected items unit of measure.  The text box above QTY accepts 4 character int input.  The "Add to Order" button addes the selected item and your QTY input to the order array.   You can add a total of 10 different items to the array at a given time. 

The clear Selected button clears the drop down box, the data grid, the selected item description box, and UOM box. 


PREVIEW ORDER:
The "Preview order" tab displays any items you have added to the order array, up to 10 items.  If you attempt to add more then 10 items to the array a message displays informing you that you cannot do so.   The preview order page shows the "Stock ID#", "Description", unit of measure, and QTY you requested.

The "Clear Order" button clears the entire order array to start from scratch.


SEND ORDER:
 The "Send order" tab contains 5 input text boxes and a drop down box.  The "Enter your Project Name:" box accepts the name of a project which is later placed in a input field on the PDF that is generated.   The "WR#:" field accepts only integers as vaild input and will be used to later populate the PDF that is generated.  The "Select Service Territory:" drop down box selects which email address array to be used when sending the completed work request.  If "Reading" is selected then any email addresses in the Reading array will recieved the completed email.  The "Enter Your User Name:" field is used to accept the AD username of the person making the request.  This input is used in both the AD authentication prior to sending the email as well as the From address in the email field.  The "Enter Your Password:" is used to input the AD password for the associated username in the "Enter Your User Name:" field.  The password is stored as a secure string for AD authentication purposes.   The "Additional Comments" field is the only non-required field on this tab.  The additional comments input will be displayed in the final email once it is sent to the email array. 
 
 The "Complete Order" button first checks if all fields on this tab have been filled out as well as if any items have been added to the order.  It then runs an AD authentication with the username and password provided in the coresponding input boxes.  If either any of the required input fields are left blank or the AD authentication fails a message pops up informing the user of this. 
 
 If everything is entered correctly and the AD authentication passes then  a copy of the template.pdf is filled out with the information inputed by the users selections and text box input.  These fields include: Project name, work request number, parts ordered, the amount requested, the requestors name, and the date and time it was requested.  This PDF is then emailed to all members of the selected array email list with a subject line of "Field Material Request" and a from address of the email account associated with the AD authentication process.  The email also contains the date and time of the request, who else the email was sent to, and any additional comments the user may have added.  It then renames the created email on the local client to the date and moves it to a backup directory should the need arrise to review the request. 
 
 ########################################################################################################################
 
 There is alot that has been changed in the raw script in order to protect our data but i can assist with supplying the correct fields upon request. 


S
