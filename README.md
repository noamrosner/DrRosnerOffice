# Doctor's Office Automation
<br />
Created automations to enhance efficiency and improve patient services in a medical office setting. <br /><br />
 
- WhatsApp Appotiment Reminder

  •	Medical practice provides a wide range of services at a number of facilities to a large patient population
  •	Patient appointments are recorded in dedicated excel files (xlsx) depending on the service (e.g., consultation, procedure, etc) and its location/relevant medical facility (e.g., Private Clinics and various Medical Centers)
  •	The program is a simple minimal window with buttons representing each of the reminder options
  •	The user can choose the desired action, and press the corresponding button 
  •	A window will pop up and the user can choose/access the relevant excel file
  •	The program will open the chosen file (i.e., template of appointments for consultation or other services at relevant medical facilities)
  •	The program with then iterate through the rows and will send an appointment 

![Screen Shot 2022-10-18 at 12 52 55](https://user-images.githubusercontent.com/95490556/196398542-b3dac571-c77e-49c8-8def-c28d2e23cc48.png)

 After choosing the action and pressing the button accordingly a window will pop up, where the user can choose the relevant file.

![Screen Shot 2022-10-18 at 12 56 39](https://user-images.githubusercontent.com/95490556/196399228-76e6477a-c9a8-4537-abd9-c8e0f3de6bfd.png)

 The program will open the Excel file, for example a template of appotiments for procedures: <br />
 
![Screen Shot 2022-10-18 at 13 05 47](https://user-images.githubusercontent.com/95490556/196401398-e5373d82-040a-49f8-8130-e48ac3459d76.png)  
<br />
 The program will itirate through the rows and will send each patient a reminder.
     The program will open WhatsApp and send the messages, each message containg the relevant information as specified in the code, the patient's name, the date and time of the appointment and other relevant information.  
     
 The program can handle with issues that can be caused from the user's behavior and bad input.
 The program will notify at the end if some users haven't recieved the reminder and will alert what caused the specific issue.  
 
<br />
 The program is written in Python, using "openpyxl" for handling with XLSX file, "tkinter" for the UI and "pywhatkit" for sending WhatsApp messages.
   

<br /><br /><br /><br /><br />
- Breath Results Automation

   The results are seen in an Excel file (xlsx), the program can go over the file while printing the relevant information for the doctor to review.
   The program will add the Doctor's recommendation and the signature in the appropriate place.
