---Common Info---

PowerShell script to collect major hardware and some software information for local system. Designed for Windows OS family with Powershell version 5.1 or higher.

When executed, script collects system hardware data and network settings to save as file of your choice in script's root directory using local system name as filename. Output data formats are: JSON, CSV, XML, HTML(GUI) local files or SQL insert to your database. Multiple choiсe is possible. Script to create MySQL Database also provied in this repository. So, if you started script from D:\ drive, your system name is I123456-Home and you choose *.csv and *.json as output file format you end up with D:\I123456-Home.csv and D:\I123456-Home.json files

Incase your Windows OS Powershell execution policy is "Restricted" (Which is highly possible because this policy enabled by Default) you need to change execution policy either manually or start hwinfo.ps1 via hwinfo-start.bat (Run as Administrator) It will set *.ps1 files execution policy to "RemoteSigned" and return it to "Default" later.

---MySQL Info---

Incase you do not have MySQL Server, here is a brief tutorial (For ubuntu-server 22.04 LTS)

01. Install MySQL Server:

(It is supposed that you already have Putty and connect to your ubintu-server via Putty terminal)

>sudo apt update sudo apt install mysql-server

Wait a bit while Ubuntu finish up it console magic then make sure that mysql-server is up and running:

>sudo service mysql status

If everything ok, you must see mysql-server status like: active (running) highlighted by beautiful green color. MySQL Server installed. You're breathtaking!

02. Configure MySQL to accept remote connections: 

Install and run Midnight Commander

>sudo install mc 

>sudo mc

You must see amazing blue panels of Midnight Commander. Navigate to MySQL configuration file **/etc/mysql/mysql.conf.d/mysql.cnf** and press F4 to open file for editing. Find "bind-address" string and assign THE SAME IP-ADDRESS WHICH HAS YOUR UBUNTU-SERVER. For example, if your ubuntu-server IP address is 192.168.0.70 your bind-address string must looks like: bind-address = 192.168.0.70 

Press F2 to save changes, then press F10 to exit file. Type "exit" to exit amazing blue panels of Midnight Commander. Restart mysql-server for your efforts to take effect:

>sudo service mysql restart

MySQL Server configured. Good job, soldier!

03. Create MySQL Database and its tables.

Download mysql_database.sql script from this repository and copy its content. Open mysql-server console:

>sudo mysql

Paste contetnt of mysql_database.sql script in console by clicking RMB (Right Mouse Button) in Putty terminal. Wait while database and tables will be created. Press ENTER. Well done!

04. Create user for remote clients (You must still have mysql-console open in your terminal) 
Execute commands:

CREATE USER 'sqlwriter'@'192.168.0.%' IDENTIFIED BY 'your-secure-password'; 

GRANT INSERT, REFERENCES ON pcinfo.* TO 'sqlwriter'@'192.168.0.%'; 

FLUSH PRIVILEGES;

'sqlwriter' is username which remote clients will use to write data to your MySQL server's database tables. 

'192.168.0.%' is your network (JUST AN EXAMPLE, CHANGE IT TO YOUR CURRENT NETWORK RANGE). 

'your-secure-password' is, well, your secure password.

Now your MySQL server will accept connections from remote hosts of your network and allow them to INSERT and REFRENCES data in your databse. Yay! 
P.S. Remember, if (when) you would need to read or manage your data you'd better create user, say 'sqlreader', who can SELECT, UPDATE and do some other SQL stuff of your choice.

05. It is not over yet!

Download Oracles MySQL\Connector from their official site and make sure that current script can access it from script's root folder.
Just install connector, go to connector's folder and copy-paste MySQL.Data.dll to directory form where hwinfo.ps1 can download and use it.

You're good to go!
