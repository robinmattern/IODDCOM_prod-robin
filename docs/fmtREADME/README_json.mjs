//console.log( require( 'util').inspect( README_JSON(), { depth: 9 } ) )

  export default README_JSON

function README_JSON() {

return [
, { proj: "IODD.COM"
  , url:  'C:/WEBs/8020/VMs/et218t/webs/nodeapps/FRApps_/prod-robin_v21017/'
  , name: "IODD_dev01-robin"
  , stages: [

  , { stage: "Client1 Home Apps"
    , apps: [

    , { app: "1c1. IODD Home App"
      , txt: "Brilliant work by Rick "
      , url: 'home/index.html' }
    ] } // eoo stage apps

  , { stage: "Server1 getJSON Tool", apps: [

    , { app: "1s1. JSON Data API"
      , txt: `A Node script to extract data from MySQL DB.  See .env for parameters. 
        
              1. Install Server1 modules   
                 $ cd server1   
                 $ npm install   
                      
              2. Configure MySQL extraction parameters   
                 $ cd 1s1*   
                 $ nano .env   
                 - DBSQL1-n=SELECT * FROM ... statements will be joined together into one db.json file   
                 - RENAME_EM=false will add a timestamp to the saved db.json and db.json.js   
                 - RENAME_EM=true will save the current file with a timestamp      
                       
              3. Copy the db.json.js or db.json to its data locaton      
                 - Remove the _vYMMDD.HHMM if necessary   
                 - Copy the db.json file for use by json-server    
              `         
      , url: 'server1/1s1-iodd-json.js/testAPIs.html' }
    ] } // eoo stage apps

  , { stage: "Server2 JSON Server APIs", apps: [

    , { app: "2s1. JSON Server API"
      , txt: "A simple API that returns data using json-server and a db.json file."
      , url: 'server2/2s1-json-server-api/testAPIs.html' }
    ] } // eoo stage apps

  , { stage: "Server3 MySQL Server APIs", apps: [

    , { app: "3s1. MySQL Data API"
      , txt: "A simple API that returns data using an express app connected to MySQL DB Server."
      , txt: "Returns data requested viawith an API call."
      , url: 'server2/2s1-json-server-api/testAPIs.html' }

    ] } // eoo stage apps
  ] } // eoo proj stages

, { proj: "formR's React Data Apps"
  , url:  'C:/WEBs/8020/VMs/et218t/webs/nodeapps/FormR_/prod-master/'
  , name: "FormR_prod-master", stages: [

  , { stage: "Client React Admin Apps", apps: [

    , { app: "1c. World"
      , txt: "A React client app to view data in the World database."
      , url: 'client/1c-world/public/index.html' }

    , { app: "3c. RAuth"
      , txt: "A React client app to authenticate users in the RAuth table."
      , url: 'client/3c-rauth/public/index.html' }

    , { app: "4c. Admin"
      , txt: "A React client app to view, add, edit or delete users in the RAuth table"
      , url: 'client/4c-admin/public/index.html' }

    , { app: "5c. formR"
      , txt: "A React client app to view, add, edit or delete data in the World database"
      , url: 'client/5c-formr/public/index.html' }
    ] }

  , { stage: "Server MySQL APIs", apps: [

    , { app: "1s. World"
      , txt: "A React server api to view data in the World database."
      , url: 'server/1s-world/index.html' }

    , { app: "3s. RAuth"
      , txt: "A React server api to authenticate users in the RAuth table."
      , url: 'server/3s-rauth/index.html' }

    , { app: "4s. Admin"
      , txt: "A React server api to view, add, edit or delete users in the RAuth table"
      , url: 'server/4s-admin/index.html' }

    , { app: "5s. formR"
      , txt: "A React server api to view, add, edit or delete data in the World database"
      , url: 'server/5s-formr/index.html' }

    ] } // eoo stage apps
  ] } // eoo proj stages
] // eoo projs

} // eof README_JSON()
