How to use
1.	Install environment support 
1.1	Install Python from https://www.python.org/downloads/ 
-	Choose Python3 version 3.7 up to stable
-	Check version : Open terminal  console and Enter “python --version”
1.2	Install pip from guide https://www.liquidweb.com/kb/install-pip-windows/
-	Check version : Open terminal console and Enter “pip --version”
1.3	Install Libs from pip command 
-	Open terminal console and Enter “pip install -r requirements.txt”

2.	Open filename “config.py” for edit data
-	USERNAME = Enter user to login N Smart
-	PASSWORD = Enter password to login N Smart
- DEPARTMENT_ID = Find id from "http://nsmart.nhealth-asia.com/MTDPDB01/reserve/reserve_back1.php" 
    view html code and get value  "<input value="00018" type="hidden" name="reserve_branchid">"
- BRANCH_ID = Find id from "http://nsmart.nhealth-asia.com/MTDPDB01/reserve/reserve_back1.php" 
    view html code and get value "<input value="0001801" type="hidden" name="reserve_dept">
- DEPT = It is list data of department and code, Find from "http://nsmart.nhealth-asia.com/MTDPDB01/reserve/reserve_back1.php" 
    view html code find "<select name="reserve_sub_dept">" and get "<OPTION VALUE="00018010002">BME-SSH</OPTION>"

3.	Start process
-	Open terminal console and enter “python bme_pool.py”
- "Enter Excel file name : "  input pool file Ex.  pool_mar_2020



