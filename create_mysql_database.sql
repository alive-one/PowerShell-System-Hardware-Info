/*NOTE: It is strongly recommend to create DB as *NIX server user
and DO NOT give privileges to create or alter DB to Powershell Script from clients PCs
for they must only have PRIVILEGES to INSERT and REFERENCES data to DB Tables. 
*/

/* Create Your Database 
(You can choose any name you like, but do not forget to edit MySQL configuration part of hwinfo.ps1 then)
*/
CREATE DATABASE IF NOT EXISTS pcinfo;

-- Use your database 
USE pcinfo;
 
-- Baseboards 
CREATE TABLE baseboards (
    baseboard_id int auto_increment primary key,
    baseboard_manufacturer varchar(64) default 'unknown',
    baseboard_model varchar(128) default 'unknown',
    baseboard_ramslots tinyint unsigned
    );
 
 /* Central Processing Units 
 (cpus)M:1(baseboard)
 */
CREATE TABLE cpus (
    cpu_id int auto_increment primary key,
    cpu_socket smallint unsigned default 0,
    cpu_manufacturer varchar(32) default 'unknown',
    cpu_model varchar(128) default 'unknown',
    cpu_clock_mhz smallint unsigned default 0,
    cpu_cores tinyint unsigned default 0,
    cpu_threads tinyint unsigned default 0,
    baseboard_id int not null,
    FOREIGN KEY (baseboard_id) REFERENCES baseboards(baseboard_id)
    );

/* RAM Types 
(ramtype)1:M(rammodules) 
(ramtype)1:M(computers)
*/
create table ramtypes (
	ramtype_id int auto_increment primary key,
	ram_type varchar(24) default 'unknown'
    );

/* RAM Modules Table 
(ramtypes)1:M(rammodules)
(rammodules)M:1(baseboards)
*/
CREATE TABLE rammodules (
    module_id int auto_increment primary key,
    bank_label varchar(64) default 'unknown',
    module_size_mb smallint unsigned default 0,
    module_speed_mhz smallint unsigned default 0,
    baseboard_id int not null,
    ramtype_id int not null,
    FOREIGN KEY (baseboard_id) references baseboards(baseboard_id),
    FOREIGN KEY (ramtype_id) references ramtypes(ramtype_id)
    );

/* Computers 
(baseboard)1:M(computers)
(ramtypes)M:1(computer)
*/
CREATE TABLE computers (
    computer_id int auto_increment primary key,
    computer_name varchar(196) default 'unknown',
    server_date datetime default NOW(),
    total_ram_gb tinyint unsigned default 0,
    baseboard_id int not null,
    ramtype_id int not null,
    FOREIGN KEY (baseboard_id) REFERENCES baseboards(baseboard_id),
    FOREIGN KEY (ramtype_id) REFERENCES ramtypes(ramtype_id)
    );

/* Storage Devices
(storages)M:1(computer)
*/
CREATE TABLE storages (
    storage_id int auto_increment primary key,
    storage_manufacturer varchar(64) default 'unknown',
    storage_model varchar(128) default 'unknown',
    storage_bus varchar(16) default 'unknown',
    storage_type varchar(16) default 'unknown',
    storage_size_gb int unsigned default 0,
    computer_id int not null,
    FOREIGN KEY (computer_id) REFERENCES computers(computer_id)
    );
/* Graphical Adapters
(gpus)M:1(computer)
*/   
CREATE TABLE gpus (
    gpu_id int auto_increment primary key,
    gpu_manufacturer varchar(24) default 'unknown',
    gpu_model varchar(128) default 'unknown',
    gpu_memory_mb smallint unsigned,
    computer_id int not null,
    FOREIGN KEY (computer_id) REFERENCES computers(computer_id)
    );
  
  /*
  Network Adapters
  (netadapters)M:1(computer)
  */
CREATE TABLE netadapters (
    netadapter_id int auto_increment primary key,
    netadapter_manufacturer varchar(32) default 'unknown',
    netadapter_model varchar(128) default 'unknown',
    netadapter_mac varchar(17) default 'unknown',
    netadapter_type varchar(32) default 'unknown',
    netadapter_speed_mbit smallint unsigned, 
    computer_id int not null,   
    FOREIGN KEY (computer_id) REFERENCES computers(computer_id)
    );
  
  /*
  Network Connections
  (netconnections)M:1(computer)
  */
CREATE TABLE netconnections (
    connection_id int auto_increment primary key,
    сonnection_name VARCHAR(128) default 'unknown',
    connection_mac varchar(24) default 'unknown',
    connection_ip bigint unsigned not null,
    connection_netmask bigint unsigned not null,
    connection_gateway bigint unsigned not null,
    computer_id int not null,   
    FOREIGN KEY (computer_id) REFERENCES computers(computer_id)
    );
