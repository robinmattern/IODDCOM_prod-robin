
## Instructions for building a VM with MySQL on AWS
See Video: https://www.youtube.com/watch?v=FanoTGjkxhQ&t=1135s<br>
Blog: https://jasonwatmore.com/post/2019/11/18/react-nodejs-on-aws-how-to-deploy-a-mern-stack-app-to-amazon-ec2

 1. Create local folder for VM
    ```
    mkdir ./VMs/sc216
    ```

 2. Download AWS KeyPair (e.g Putty format: `sc216_FormR-KeyPair.pem`) and connect to server via SSH
    ```
    $ chmod 400 sc216_FormR-KeyPair.pem

    $ ssh -i .ssh/sc216_FormR-KeyPair.pem ubuntu@ec2-13-57-57-151.us-west-1.compute.amazonaws.com

        The authenticity of host 'ec2-13-57-57-151.us-west-1.compute.amazonaws.com (13.57.57.151)' can't be established.
        ECDSA key fingerprint is SHA256:g4x9xtvYfLVzBufgcog1lGNlfmK60rivsARnnovV+L4.
        Are you sure you want to continue connecting (yes/no)? yes

        Warning: Permanently added 'ec2-13-57-57-151.us-west-1.compute.amazonaws.com,13.57.57.151' (ECDSA) to the list of known hosts.
        Welcome to Ubuntu 18.04.5 LTS (GNU/Linux 5.4.0-1029-aws x86_64)

         * Documentation:  https://help.ubuntu.com
         * Management:     https://landscape.canonical.com
         * Support:        https://ubuntu.com/advantage

          System information as of Tue Dec 29 00:28:15 UTC 2020

          System load:  0.0               Processes:           92
          Usage of /:   14.6% of 7.69GB   Users logged in:     0
          Memory usage: 17%               IP address for eth0: 172.31.30.151
          Swap usage:   0%

        0 packages can be updated.
        0 updates are security updates.

        The programs included with the Ubuntu system are free software;
        the exact distribution terms for each program are described in the
        individual files in /usr/share/doc/*/copyright.

        Ubuntu comes with ABSOLUTELY NO WARRANTY, to the extent permitted by
        applicable law.

        To run a command as administrator (user "root"), use "sudo <command>".
        See "man sudo_root" for details.
     ```

  3. Download and configure Bitvise SSH / SFTP browser

  4. Create bin folder for bash scripts
     ```
     $ sudo mkdir       /home/_
     $ sudo mkdir       /home/._0/bin
     $ echo #!/bin/sh  >/home/._0/bin/hello.sh
     $ echo #!/bin/sh >>/home/._0/bin/hello.sh
     $ chmod 777        /home/._0/bin/hello.sh
     $ hello.sh
     ```

  5. Install NodeJS, NGinx and MongoDB
     Run script from ... and/or run a modified version of it without MongoDB installation
     ```
     $ curl https://gist.githubusercontent.com/cornflourblue/f0abd30f47d96d6ff127fe8a9e5bbd9f/raw/e3047c9dc3ce8b796e7354c92d2c47ce61981d2f/setup-nodejs-mongodb-production-server-on-ubuntu-1804.sh | sudo bash
     
     $ cp setup-nodejs-server-on-ubuntu-1804.sh /home/_/bin/setup-nodejs-server-on-ubuntu-1804.sh
     setup-nodejs-server-on-ubuntu-1804.sh >setup-nodejs-server-on-ubuntu_v01228.log
     ```
  6. Install MySQL
     ```
     $ sudo apt-get update

        Hit:1 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic InRelease
        Get:2 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates InRelease [88.7 kB]
        Get:3 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-backports InRelease [74.6 kB]
        Hit:4 https://deb.nodesource.com/node_10.x bionic InRelease
        Get:5 http://security.ubuntu.com/ubuntu bionic-security InRelease [88.7 kB]
        Fetched 252 kB in 0s (575 kB/s)
        Reading package lists... Done

     $ sudo apt-get install mysql-server

        Reading package lists... Done
        Building dependency tree
        Reading state information... Done
        The following additional packages will be installed:
          libaio1 libcgi-fast-perl libcgi-pm-perl libencode-locale-perl libevent-core-2.1-6 libfcgi-perl
          libhtml-parser-perl libhtml-tagset-perl libhtml-template-perl libhttp-date-perl
          libhttp-message-perl libio-html-perl liblwp-mediatypes-perl libtimedate-perl liburi-perl
          mysql-client-5.7 mysql-client-core-5.7 mysql-common mysql-server-5.7 mysql-server-core-5.7
        Suggested packages:
          libdata-dump-perl libipc-sharedcache-perl libwww-perl mailx tinyca
        The following NEW packages will be installed:
          libaio1 libcgi-fast-perl libcgi-pm-perl libencode-locale-perl libevent-core-2.1-6 libfcgi-perl
          libhtml-parser-perl libhtml-tagset-perl libhtml-template-perl libhttp-date-perl
          libhttp-message-perl libio-html-perl liblwp-mediatypes-perl libtimedate-perl liburi-perl
          mysql-client-5.7 mysql-client-core-5.7 mysql-common mysql-server mysql-server-5.7
          mysql-server-core-5.7
        0 upgraded, 21 newly installed, 0 to remove and 50 not upgraded.
        Need to get 19.7 MB of archives.
        After this operation, 157 MB of additional disk space will be used.
        Do you want to continue? [Y/n] y
        Get:1 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 mysql-common all 5.8+1.0.4 [7308 B]
        Get:2 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates/main amd64 libaio1 amd64 0.3.110-5ubuntu0.1 [6476 B]
        Get:3 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates/main amd64 mysql-client-core-5.7 amd64 5.7.32-0ubuntu0.18.04.1 [6660 kB]
        Get:4 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates/main amd64 mysql-client-5.7 amd64 5.7.32-0ubuntu0.18.04.1 [1943 kB]
        Get:5 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates/main amd64 mysql-server-core-5.7 amd64 5.7.32-0ubuntu0.18.04.1 [7455 kB]
        Get:6 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libevent-core-2.1-6 amd64 2.1.8-stable-4build1 [85.9 kB]
        Get:7 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates/main amd64 mysql-server-5.7 amd64 5.7.32-0ubuntu0.18.04.1 [2935 kB]
        Get:8 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libhtml-tagset-perl all 3.20-3 [12.1 kB]
        Get:9 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 liburi-perl all 1.73-1 [77.2 kB]
        Get:10 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libhtml-parser-perl amd64 3.72-3build1 [85.9 kB]
        Get:11 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libcgi-pm-perl all 4.38-1 [185 kB]
        Get:12 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libfcgi-perl amd64 0.78-2build1 [32.8 kB]
        Get:13 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libcgi-fast-perl all 1:2.13-1 [9940 B]
        Get:14 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libencode-locale-perl all 1.05-1 [12.3 kB]
        Get:15 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libhtml-template-perl all 2.97-1 [59.0 kB]
        Get:16 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libtimedate-perl all 2.3000-2 [37.5 kB]
        Get:17 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libhttp-date-perl all 6.02-1 [10.4 kB]
        Get:18 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libio-html-perl all 1.001-1 [14.9 kB]
        Get:19 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 liblwp-mediatypes-perl all 6.02-1 [21.7 kB]
        Get:20 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic/main amd64 libhttp-message-perl all 6.14-1 [72.1 kB]
        Get:21 http://us-west-1.ec2.archive.ubuntu.com/ubuntu bionic-updates/main amd64 mysql-server all 5.7.32-0ubuntu0.18.04.1 [9944 B]
        Fetched 19.7 MB in 1s (38.9 MB/s)
        Preconfiguring packages ...
        Selecting previously unselected package mysql-common.
        (Reading database ... 62832 files and directories currently installed.)
        Preparing to unpack .../0-mysql-common_5.8+1.0.4_all.deb ...
        Unpacking mysql-common (5.8+1.0.4) ...
        Selecting previously unselected package libaio1:amd64.
        Preparing to unpack .../1-libaio1_0.3.110-5ubuntu0.1_amd64.deb ...
        Unpacking libaio1:amd64 (0.3.110-5ubuntu0.1) ...
        Selecting previously unselected package mysql-client-core-5.7.
        Preparing to unpack .../2-mysql-client-core-5.7_5.7.32-0ubuntu0.18.04.1_amd64.deb ...
        Unpacking mysql-client-core-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Selecting previously unselected package mysql-client-5.7.
        Preparing to unpack .../3-mysql-client-5.7_5.7.32-0ubuntu0.18.04.1_amd64.deb ...
        Unpacking mysql-client-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Selecting previously unselected package mysql-server-core-5.7.
        Preparing to unpack .../4-mysql-server-core-5.7_5.7.32-0ubuntu0.18.04.1_amd64.deb ...
        Unpacking mysql-server-core-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Selecting previously unselected package libevent-core-2.1-6:amd64.
        Preparing to unpack .../5-libevent-core-2.1-6_2.1.8-stable-4build1_amd64.deb ...
        Unpacking libevent-core-2.1-6:amd64 (2.1.8-stable-4build1) ...
        Setting up mysql-common (5.8+1.0.4) ...
        update-alternatives: using /etc/mysql/my.cnf.fallback to provide /etc/mysql/my.cnf (my.cnf) in auto mode
        Selecting previously unselected package mysql-server-5.7.
        (Reading database ... 63000 files and directories currently installed.)
        Preparing to unpack .../00-mysql-server-5.7_5.7.32-0ubuntu0.18.04.1_amd64.deb ...
        Unpacking mysql-server-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Selecting previously unselected package libhtml-tagset-perl.
        Preparing to unpack .../01-libhtml-tagset-perl_3.20-3_all.deb ...
        Unpacking libhtml-tagset-perl (3.20-3) ...
        Selecting previously unselected package liburi-perl.
        Preparing to unpack .../02-liburi-perl_1.73-1_all.deb ...
        Unpacking liburi-perl (1.73-1) ...
        Selecting previously unselected package libhtml-parser-perl.
        Preparing to unpack .../03-libhtml-parser-perl_3.72-3build1_amd64.deb ...
        Unpacking libhtml-parser-perl (3.72-3build1) ...
        Selecting previously unselected package libcgi-pm-perl.
        Preparing to unpack .../04-libcgi-pm-perl_4.38-1_all.deb ...
        Unpacking libcgi-pm-perl (4.38-1) ...
        Selecting previously unselected package libfcgi-perl.
        Preparing to unpack .../05-libfcgi-perl_0.78-2build1_amd64.deb ...
        Unpacking libfcgi-perl (0.78-2build1) ...
        Selecting previously unselected package libcgi-fast-perl.
        Preparing to unpack .../06-libcgi-fast-perl_1%3a2.13-1_all.deb ...
        Unpacking libcgi-fast-perl (1:2.13-1) ...
        Selecting previously unselected package libencode-locale-perl.
        Preparing to unpack .../07-libencode-locale-perl_1.05-1_all.deb ...
        Unpacking libencode-locale-perl (1.05-1) ...
        Selecting previously unselected package libhtml-template-perl.
        Preparing to unpack .../08-libhtml-template-perl_2.97-1_all.deb ...
        Unpacking libhtml-template-perl (2.97-1) ...
        Selecting previously unselected package libtimedate-perl.
        Preparing to unpack .../09-libtimedate-perl_2.3000-2_all.deb ...
        Unpacking libtimedate-perl (2.3000-2) ...
        Selecting previously unselected package libhttp-date-perl.
        Preparing to unpack .../10-libhttp-date-perl_6.02-1_all.deb ...
        Unpacking libhttp-date-perl (6.02-1) ...
        Selecting previously unselected package libio-html-perl.
        Preparing to unpack .../11-libio-html-perl_1.001-1_all.deb ...
        Unpacking libio-html-perl (1.001-1) ...
        Selecting previously unselected package liblwp-mediatypes-perl.
        Preparing to unpack .../12-liblwp-mediatypes-perl_6.02-1_all.deb ...
        Unpacking liblwp-mediatypes-perl (6.02-1) ...
        Selecting previously unselected package libhttp-message-perl.
        Preparing to unpack .../13-libhttp-message-perl_6.14-1_all.deb ...
        Unpacking libhttp-message-perl (6.14-1) ...
        Selecting previously unselected package mysql-server.
        Preparing to unpack .../14-mysql-server_5.7.32-0ubuntu0.18.04.1_all.deb ...
        Unpacking mysql-server (5.7.32-0ubuntu0.18.04.1) ...
        Setting up libhtml-tagset-perl (3.20-3) ...
        Setting up libevent-core-2.1-6:amd64 (2.1.8-stable-4build1) ...
        Setting up libencode-locale-perl (1.05-1) ...
        Setting up libtimedate-perl (2.3000-2) ...
        Setting up libio-html-perl (1.001-1) ...
        Setting up liblwp-mediatypes-perl (6.02-1) ...
        Setting up libaio1:amd64 (0.3.110-5ubuntu0.1) ...
        Setting up liburi-perl (1.73-1) ...
        Setting up libhtml-parser-perl (3.72-3build1) ...
        Setting up libcgi-pm-perl (4.38-1) ...
        Setting up mysql-client-core-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Setting up libfcgi-perl (0.78-2build1) ...
        Setting up libhttp-date-perl (6.02-1) ...
        Setting up libhtml-template-perl (2.97-1) ...
        Setting up mysql-server-core-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Setting up libcgi-fast-perl (1:2.13-1) ...
        Setting up libhttp-message-perl (6.14-1) ...
        Setting up mysql-client-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        Setting up mysql-server-5.7 (5.7.32-0ubuntu0.18.04.1) ...
        update-alternatives: using /etc/mysql/mysql.cnf to provide /etc/mysql/my.cnf (my.cnf) in auto mode
        Renaming removed key_buffer and myisam-recover options (if present)
        Created symlink /etc/systemd/system/multi-user.target.wants/mysql.service ? /lib/systemd/system/mysql.service.
        Setting up mysql-server (5.7.32-0ubuntu0.18.04.1) ...
        Processing triggers for libc-bin (2.27-3ubuntu1.2) ...
        Processing triggers for systemd (237-3ubuntu10.42) ...
        Processing triggers for man-db (2.8.3-2ubuntu0.1) ...
        Processing triggers for ureadahead (0.100.0-21) ...

      $ mysql --version

        mysql  Ver 14.14 Distrib 5.7.32, for Linux (x86_64) using  EditLine wrapper
      ```

  7. Lock down MySQL
      ```
      $ sudo mysql_secure_installation

        Securing the MySQL server deployment.

        Connecting to MySQL using a blank password.

        VALIDATE PASSWORD PLUGIN can be used to test passwords
        and improve security. It checks the strength of password
        and allows the users to set only those passwords which are
        secure enough. Would you like to setup VALIDATE PASSWORD plugin?

        Press y|Y for Yes, any other key for No: y

        There are three levels of password validation policy:

        LOW    Length >= 8
        MEDIUM Length >= 8, numeric, mixed case, and special characters
        STRONG Length >= 8, numeric, mixed case, special characters and dictionary                  file

        Please enter 0 = LOW, 1 = MEDIUM and 2 = STRONG: 1
        Please set the password for root here.

        New password: Scn.adm216

        Re-enter new password: Scn.adm216

        Estimated strength of the password: 100

        Do you wish to continue with the password provided?(Press y|Y for Yes, any other key for No) : y
        By default, a MySQL installation has an anonymous user,
        allowing anyone to log into MySQL without having to have
        a user account created for them. This is intended only for
        testing, and to make the installation go a bit smoother.
        You should remove them before moving into a production
        environment.

        Remove anonymous users? (Press y|Y for Yes, any other key for No) : y
        Success.

        Normally, root should only be allowed to connect from
        'localhost'. This ensures that someone cannot guess at
        the root password from the network.

        Disallow root login remotely? (Press y|Y for Yes, any other key for No) : y
        Success.

        By default, MySQL comes with a database named 'test' that
        anyone can access. This is also intended only for testing,
        and should be removed before moving into a production
        environment.

        Remove test database and access to it? (Press y|Y for Yes, any other key for No) : n

         ... skipping.
        Reloading the privilege tables will ensure that all changes
        made so far will take effect immediately.

        Reload privilege tables now? (Press y|Y for Yes, any other key for No) : y
        Success.

        All done!
     ```

  8. Enable root to use MySQL shell.  First login doesn't require a password.
     ```
     $ sudo mysql

     mysql> SELECT user,authentication_string,plugin,host FROM mysql.user;
         +------------------+-------------------------------------------+-----------------------+-----------+
         | user             | authentication_string                     | plugin                | host      |
         +------------------+-------------------------------------------+-----------------------+-----------+
         | root             |                                           | auth_socket           | localhost |
         | mysql.session    | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost |
         | mysql.sys        | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost |
         | debian-sys-maint | *3A18876D8D1027DC37AAE3FA14D6C9A22FBA3ED3 | mysql_native_password | localhost |
         +------------------+-------------------------------------------+-----------------------+-----------+
         4 rows in set (0.00 sec)

     mysql> ALTER USER 'root'@'localhost' IDENTIFIED WITH mysql_native_password BY 'Scn.adm216';
     mysql> FLUSH PRIVILEGES;
     mysql> SELECT user,authentication_string,plugin,host FROM mysql.user;
        +------------------+-------------------------------------------+-----------------------+-----------+
        | user             | authentication_string                     | plugin                | host      |
        +------------------+-------------------------------------------+-----------------------+-----------+
        | root             | *13B429B7780A66DCEF0B336DF32C025CF0D97496 | mysql_native_password | localhost |
        | mysql.session    | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost |
        | mysql.sys        | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost |
        | debian-sys-maint | *3A18876D8D1027DC37AAE3FA14D6C9A22FBA3ED3 | mysql_native_password | localhost |
        +------------------+-------------------------------------------+-----------------------+-----------+
        4 rows in set (0.00 sec)

     mysql> EXIT;

        Bye

 9. Stop, Start and check status of MySQL
      ```
      $ sudo systemctl stop mysql
      $ sudo systemctl start mysql
      $ sudo systemctl status mysql.service

         mysql.service - MySQL Community Server
            Loaded: loaded (/lib/systemd/system/mysql.service; enabled; vendor preset: enabled)
            Active: active (running) since Tue 2020-12-29 03:10:26 UTC; 24s ago
           Process: 16541 ExecStart=/usr/sbin/mysqld --daemonize --pid-file=/run/mysqld/mysqld.pid (code=exited, status=
           Process: 16518 ExecStartPre=/usr/share/mysql/mysql-systemd-start pre (code=exited, status=0/SUCCESS)
          Main PID: 16543 (mysqld)
             Tasks: 27 (limit: 1140)
            CGroup: /system.slice/mysql.service
                    +-16543 /usr/sbin/mysqld --daemonize --pid-file=/run/mysqld/mysqld.pid

         Dec 29 03:10:26 ip-172-31-30-151 systemd[1]: Starting MySQL Community Server...
         Dec 29 03:10:26 ip-172-31-30-151 systemd[1]: Started MySQL Community Server.
      ```

10. Login to mysqladmin 
      $ sudo mysqladmin -p -u root version

        mysqladmin  Ver 8.42 Distrib 5.7.32, for Linux on x86_64
        Copyright (c) 2000, 2020, Oracle and/or its affiliates. All rights reserved.

        Oracle is a registered trademark of Oracle Corporation and/or its
        affiliates. Other names may be trademarks of their respective
        owners.

        Server version          5.7.32-0ubuntu0.18.04.1
        Protocol version        10
        Connection              Localhost via UNIX socket
        UNIX socket             /var/run/mysqld/mysqld.sock
        Uptime:                 4 min 7 sec

        Threads: 1  Questions: 2  Slow queries: 0  Opens: 105  Flush tables: 1  Open tables: 98  Queries per second avg: 0.008
    ```

11. Allow remote access to MySQL
    ```
    $ nano /etc/mysql/mysql.conf.d/mysqld.cnf

      Change line:
        bind-address            = 127.0.0.1
      to:
        bind-address            = 0.0.0.0

    $ sudo systemctl restart mysql.service

    $ sudo netstat -tulnp | grep mysql

        tcp  0   0 0.0.0.0:3306    0.0.0.0:*   LISTEN      16802/mysqld

    $ sudo mysql -p
        password: Scn.adm216

    mysql> CREATE USER 'sco'@'%' IDENTIFIED WITH mysql_native_password BY 'Scn.dbo216';
    mysql> CREATE USER 'robin'@'69.251.71.138' IDENTIFIED WITH mysql_native_password BY 'Scn.rnr216';

    mysql> SELECT user,authentication_string,plugin,host FROM mysql.user;

        +------------------+-------------------------------------------+-----------------------+---------------+
        | user             | authentication_string                     | plugin                | host          |
        +------------------+-------------------------------------------+-----------------------+---------------+
        | root             | *13B429B7780A66DCEF0B336DF32C025CF0D97496 | mysql_native_password | localhost     |
        | mysql.session    | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost     |
        | mysql.sys        | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost     |
        | debian-sys-maint | *3A18876D8D1027DC37AAE3FA14D6C9A22FBA3ED3 | mysql_native_password | localhost     |
        | sco              | *91DE2D914E5541862E3C7A2FAD3AADB4E60A81E6 | mysql_native_password | %             |
        | robin            | *F641B7862AF80A43A038893D3F08D8A014EC2387 | mysql_native_password | 69.251.71.138 |
        +------------------+-------------------------------------------+-----------------------+---------------+
        6 rows in set (0.00 sec)

    mysql> EXIT
    ```

12. Grant ALL privileges to Users, robin and bruce   
    Note: after granting ALL privileges to robin, I was able to remotely create users,
          but I could only grant ALL privileges to bruce when logged in as root via the server command line.
    ```
    $ sudo mysql -p
    mysql> CREATE USER 'io'@'%'    IDENTIFIED WITH mysql_native_password BY 'IO.usr216';
    mysql> CREATE USER 'bruce'@'%' IDENTIFIED WITH mysql_native_password BY 'IO.btr216';
    mysql> GRANT ALL PRIVILEGES ON *.* TO 'robin'@'69.251.71.138';
    mysql> GRANT ALL PRIVILEGES ON *.* TO 'bruce'@'%';

    mysql> SELECT user,authentication_string,plugin,host FROM mysql.user;

        +------------------+-------------------------------------------+-----------------------+---------------+
        | user             | authentication_string                     | plugin                | host          |
        +------------------+-------------------------------------------+-----------------------+---------------+
        | root             | *13B429B7780A66DCEF0B336DF32C025CF0D97496 | mysql_native_password | localhost     |
        | mysql.session    | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost     |
        | mysql.sys        | *THISISNOTAVALIDPASSWORDTHATCANBEUSEDHERE | mysql_native_password | localhost     |
        | debian-sys-maint | *3A18876D8D1027DC37AAE3FA14D6C9A22FBA3ED3 | mysql_native_password | localhost     |
        | sco              | *91DE2D914E5541862E3C7A2FAD3AADB4E60A81E6 | mysql_native_password | %             |
        | robin            | *F641B7862AF80A43A038893D3F08D8A014EC2387 | mysql_native_password | 69.251.71.138 |
        | io               | *AA30383F7F91857E5E250F35681C4C3CF51D3F90 | mysql_native_password | %             |
        | bruce            | *B52E6D6CC6CA47DF4A1C8C5F26BC56554B8F2A63 | mysql_native_password | %             |
        +------------------+-------------------------------------------+-----------------------+---------------+
        8 rows in set (0.00 sec)

    ```
13. Open firewall rule for port 3306
    ```
    $ sudo ufw allow 3306/tcp
        Rule added
        Rule added (v6)
    ```

14. Check remote access to MySQL
    ```
    $ telnet 192.168.1.10 3306
    ```

15. Login remotely with MySQL Shell. You can save the password and not need it on subsequent logins.
    ```
    $ mysqlsh --uri robin@13.57.57.151:3306

        Please provide the password for 'robin@13.57.57.151:3306': **********
        Save password for 'robin@13.57.57.151:3306'? [Y]es/[N]o/Ne[v]er (default No): yes

        MySQL Shell 8.0.22

        Copyright (c) 2016, 2020, Oracle and/or its affiliates.
        Oracle is a registered trademark of Oracle Corporation and/or its affiliates.
        Other names may be trademarks of their respective owners.

        Type '\help' or '\?' for help; '\quit' to exit.
        Creating a session to 'robin@13.57.57.151:3306'
        Fetching schema names for autocompletion... Press ^C to stop.
        Your MySQL connection id is 74
        Server version: 5.7.32-0ubuntu0.18.04.1 (Ubuntu)
        No default schema selected; type \use <schema> to set one.

        MySQL  13.57.57.151:3306 ssl  JS > \quit
        Bye!

    $ mysqlsh --uri robin@13.57.57.151:3306

        MySQL Shell 8.0.22

        Copyright (c) 2016, 2020, Oracle and/or its affiliates.
        Oracle is a registered trademark of Oracle Corporation and/or its affiliates.
        Other names may be trademarks of their respective owners.

        Type '\help' or '\?' for help; '\quit' to exit.
        Creating a session to 'robin@13.57.57.151:3306'
        Fetching schema names for autocompletion... Press ^C to stop.
        Your MySQL connection id is 77
        Server version: 5.7.32-0ubuntu0.18.04.1 (Ubuntu)
        No default schema selected; type \use <schema> to set one.

        MySQL  13.57.57.151:3306 ssl  JS > \sql SELECT user,authentication_string,plugin,host FROM mysql.user;

        MySQL  13.57.57.151:3306 ssl  JS > \sql select * from mysql.user;

            ERROR: 1142 (42000): SELECT command denied to user 'robin'@'c-69-251-71-138.hsd1.va.comcast.net' for table 'user'
    ```

16. Check disk usage and that MySQL is running
    ```
    $ df
        Filesystem     1K-blocks    Used Available Use% Mounted on
        udev              486744       0    486744   0% /dev
        tmpfs             100212     772     99440   1% /run
        /dev/xvda1       8065444 1868284   6180776  24% /
        tmpfs             501060       0    501060   0% /dev/shm
        tmpfs               5120       0      5120   0% /run/lock
        tmpfs             501060       0    501060   0% /sys/fs/cgroup
        /dev/loop0        100096  100096         0 100% /snap/core/10185
        /dev/loop1         28800   28800         0 100% /snap/amazon-ssm-agent/2012
        tmpfs             100212       0    100212   0% /run/user/1000

    $ ps -aux | awk /mysqld/

        ubuntu    2718  0.0  0.3   26060   3080 pts/0  S+   18:46   0:00 awk /mysqld/
        mysql    16915  0.0 19.4 1177276 194940 ?      Sl   04:36   0:21 /usr/sbin/mysqld --daemonize --pid-file=/run/mysqld/mysqld.pid
    ```

17. Create new root password 
    
    Stop the MySQL service, then start it in background without loading grant tables. 
    Login without a password and reset root's password. The shutdown and restart the 
    MySQL server and login with the new password 
   ```
    $ sudo systemctl stop mysql
    $ sudo mysqld_safe --skip-grant-tables &  

    $ mysql -u root
    mysql> ALTER USER 'root'@'localhost' IDENTIFIED BY 'MY_NEW_PASSWORD';
    mysql> FLUSH PRIVILEGES;
      or 
    mysql> UPDATE mysql.user SET authentication_string = PASSWORD('MY_NEW_PASSWORD')
    mysql> WHERE User = 'root' AND Host = 'localhost';
    mysql> FLUSH PRIVILEGES;
      Query OK, 0 rows affected (0.00 sec)

    $ mysqladmin -u root -p shutdown
      password: MY_NEW_PASSWORD  

    $ sudo systemctl start mysql

    $ mysql -u root -p
      password: MY_NEW_PASSWORD  

    ```



