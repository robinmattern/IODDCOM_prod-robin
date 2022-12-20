
<html>
 <body>
  <style><!--
    details > ul > li            { margin-top:-10px !important; margin-bottom:20px !important; }
    details > ul > li > p        { color: #810d0d; padding-left: 20px; margin-top:-17px !important; text-indent: -20px; line-height: 22px !important; }
    details > ul > li:last-child { display: none; }                    					            /* .(21218.02.1 RAM Don't display MT last child) */
    code                         { color: black; font-size: 12px; margin: 0px 0px 0px 16px !important; padding-bottom: 0px; }   /* .(21218.02.3 RAM) */
  --></style>

  <div style="margin-left:25px;">

#
# <u>IODD.COM</u>
  <h2 style="font-size:24px; margin: -18px 0px 15px 12px;">(IODD_dev01-robin)</h2>

  <details><summary><b style="font-size:24px;">Client1 Home Apps</b></summary>

  - ### [1c1. IODD Home App](home/index.html)
    Brilliant work by Rick 

  -

</details>
  <details><summary><b style="font-size:24px;">Server1 getJSON Tool</b></summary>

  - ### [1s1. JSON Data API](server1/1s1_iodd-json-js/db.json.js)
    A Node script to extract data from MySQL DB.  See .env for parameters.

              1. Install Server1 node_modules    
                `$ cd server1     `  
                `$ npm install     `  

              2. Configure MySQL extraction parameters    
                `$ cd 1s1*   `  
                `$ nano .env   `  
                 - DBNAME, DBHOST, DBUSER, DBPASS = database connection parameters  
                 - DBSQL1-n=SELECT * FROM ... statements will be joined together into one db.json file  
                 - RENAME_EM=false will add a timestamp to the saved db.json and db.json.js  
                 - RENAME_EM=true will save the current file with a timestamp  

              3. Run the extraction program   
                `$ npm run getJSON   `  

              4. Copy the extracted db.json files to their data locaton  
                 - Remove the timestamp _vYMMDD.HHMM if necessary  
                 - Copy the file, db.json.js, into ./client1/home/assets/json for use by the Client1 App  
              

  -

</details>
  <details><summary><b style="font-size:24px;">Server2 JSON Server APIs</b></summary>

  - ### [2s1. JSON Server API](server2/2s1_iodd-json-api/testAPIs.html)
    A simple API that returns data using json-server with a db.json file.

              1. Install Server2 node_modules    
                `$ cd server2     `  
                `$ npm install     `  

              2. Extract db.json from MySQL   
                 - Use the Server1 getJSON Tool described above if necessary  
                 - Copy the file, db.json, into ./server2/api/models for use by the JSON Server API  

              3. Run and test the JSON Server APIs  
                `$ cd server2   `  
                `$ cd 1s2*   `  
                `$ npm start   `  
                 - Open the file, testAPIs.html, in Live Server  
              

  -

</details>
  <details><summary><b style="font-size:24px;">Server3 MySQL Server APIs</b></summary>

  - ### [3s1. MySQL Data API](server2/2s1-json-server-api/testAPIs.html)
    Returns data requested viawith an API call.

  -

</details>

# <u>formR's React Data Apps</u>
  <h2 style="font-size:24px; margin: -18px 0px 15px 12px;">(FormR_prod-master)</h2>

  <details><summary><b style="font-size:24px;">Client React Admin Apps</b></summary>

  - ### [1c. World](client/1c-world/public/index.html)
    A React client app to view data in the World database.

  - ### [3c. RAuth](client/3c-rauth/public/index.html)
    A React client app to authenticate users in the RAuth table.

  - ### [4c. Admin](client/4c-admin/public/index.html)
    A React client app to view, add, edit or delete users in the RAuth table

  - ### [5c. formR](client/5c-formr/public/index.html)
    A React client app to view, add, edit or delete data in the World database

  -

</details>
  <details><summary><b style="font-size:24px;">Server MySQL APIs</b></summary>

  - ### [1s. World](server/1s-world/index.html)
    A React server api to view data in the World database.

  - ### [3s. RAuth](server/3s-rauth/index.html)
    A React server api to authenticate users in the RAuth table.

  - ### [4s. Admin](server/4s-admin/index.html)
    A React server api to view, add, edit or delete users in the RAuth table

  - ### [5s. formR](server/5s-formr/index.html)
    A React server api to view, add, edit or delete data in the World database

  -

</details>

  </div>
 </body>
</html>
