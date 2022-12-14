/*\
##=========+====================+================================================+
##RD         IODD-members       | Format IODD Members table
##RFILE    +====================+=======+===============+======+=================+
##FD   IODD.Members.js          |   4757| 12/08/22  9:30|   101| p1.00-21208-0930
##FD   IODD.Members.js          |   5772| 12/09/22 13:15|   105| p1.01-21209-1315
##FD   IODD-members.js          |   6007| 12/10/22 16:40|   119| p1.02-21210.1640
##DESC     .--------------------+-------+---------------+------+-----------------+
#            Read Members Data and format it
#
##LIC      .--------------------+----------------------------------------------+
#            Copyright (c) 2022 8020Data-formR * Released under
#            MIT License: http://www.opensource.org/licenses/mit-license.php

##FNCS     .--------------------+----------------------------------------------+
#            fmtMembers         | Format array of Member records
#              fmtMember        | Format each Member record
#            renderJSON         | Get, format and assign HTM from JSON                  // .(21214.02.1)
#            parseJSON          | Convert .json.js into pJSON object                    // .(21213.06.1) 
#            getJSON            | Just get JSON                                         // .(21214.02.2)
#              readFile         | Read a file (no error handling)                       // .(21210.03.1)
#            sortitems          | Sort Members by LasName, FirstName                    // .(21209.04.1)
#
##CHGS     .--------------------+----------------------------------------------+
# .(21208.01 12/08/21 RAM  9:30a| Created
# .(21209.04 12/09/21 RJS  1:15p| Add sortitems
# .(21210.03 12/10/21 RAM  4:40p| Add readFile
# .(21213.06 12/13/21 RAM 11:20p| Read db.json.js, not IODD-members_u2a.json.js
# .(21214.01 12/14/21 RAM  7:30a| Read pJSON from json-server 
# .(21214.02 12/14/21 RAM  8:30a| Use getJSON or renderJSON functions

##PRGM     +====================+===============================================+
##ID 69.600. Main               |
##SRCE     +====================+===============================================+
\*/
//========================================================================================================= #  ===============================  #

//---------------------------------------------------------------------------------------------------

       aTests='live in Browser'
       aTests='test1 in NodeJS'

       API_URL       = 'http://localhost:3000/members'

   if (aTests.match( /test1/ )) {
       var  fetch               =  require( 'node-fetch' )                              // .(21214.01.1 RAM Install NPM Module node-fetch)
       var  inspect_            =  require( 'util' ).inspect
  function  inspect( pObj, nLv ) { return inspect_( pObj, { depth: nLv ? nLv : 99 } ) }

   ( async  ( ) => {   // function( ) { ... }                                           // .(21214.01.2 RAM Enable async / await )

//     var  pJSON               =  parseJSON( '../json/db.json.js' )                    //#.(21213.06.2 RJS Was: IODD-members_u2a.json.js).(21214.02.3) 
//     var  aHTML               =  await renderJSON( fmtMembers, API_URL, 'debug' )     //#.(21214.02.3 RAM Use new renderJSON function)
//     var  aHTML               =  await renderJSON( fmtMembers, 'debug' )              //#.(21214.02.4)
           
//     var  pJSON               =  await getJSON( API_URL )                             //#.(21214.02.5 RAM Use new getJSON function)
       var  pJSON               =  await getMembersJSON( )                              // .(21214.02.13)
//     var  pJSON               =  await getJSON( )                                     //#.(21214.02.6)
       var  aHTML               =  fmtMembers( pJSON ) 

            console.log( `Rendered HTML:\n ${ aHTML }` )
            } )( );                                                                     // .(21214.01.3)

       } // eif test1 
//---------------------------------------------------------------------------------------------------

//              renderJSON( 'http://localhost:3000/members', fmtMembers, container, 'debug' ) 
   
async  function renderJSON( fmtRecords, aURL, pDIV, aBug ) {                            // .(21214.02.7 RAM Beg Write renderJSON) 
       var  bDebug              = `${aBug}`.match( 'debug' ) ?  1   : 0
            bDebug              = `${pDIV}`.match( 'debug' ) ?  1   : bDebug  
            bDebug              = `${aURL}`.match( 'debug' ) ?  1   : bDebug  
            bDIV                = `${pDIV}`.match( 'debug' ) ? aBug : pDIV 
            aAPI_URL            = `${aURL}`.match( 'http'  ) ? aURL : API_URL 

//	   var  pResponse 			=  await fetch( aAPI_URL )
//     var  pJSON               =  await pResponse.json( )
       var  pJSON               =  await getJSON( aAPI_URL )
//
//   		pJSON 				={ members: pJSON }
//    		console.log( pJSON );

 	   var  aHTML  				=  fmtRecords( pJSON )

       if (!bDebug) {
      		pDIV.innerHTML = aHTML
      		} 
       
        if (bDebug) {
//          console.log( `Response Body length: ${ pResponse.body.bytesWritten }` )
//          console.log( `Response JSON:\n ${ inspect( pJSON, 9 ) }` )
            }  
      		return aHTML 

  		    } // eof renderJSON                                                         // .(21214.02.7 RAM End) 
//     ---  -------  =  -----------------------------------

//              getJSON( 'http://localhost:3000/members' ) 

async  function getJSON( aURL ) {                                                       // .(21214.02.8 RAM Beg Write getJSON) 
       var  aAPI_URL            = `${aURL}`.match( 'http' ) ? aURL : API_URL 
	   var  pResponse 			=  await fetch( aAPI_URL );
       var  pJSON               =  await pResponse.json( );
//     var  aText               =  await pResponse.text( )  // body aslready used for ...
    return  pJSON 
             
  		    } // eof getJSON                                                            // .(21214.02.8 RAM End) 
//--------  -------  =  -------------------------------------------------------

async  function getMembersJSON( aURL ) {                                                // .(21214.02.10 RAM Beg Write getMembersJSON) 
       var  aAPI_URL            = `${aURL}`.match( 'http' ) ? aURL : API_URL 
       var  pJSON               =  await getJSON( aAPI_URL )
    return  pJSON
            }
//--------  -------  =  -------------------------------------------------------

  function  fmtMembers( pJSON ) {
//     var  mMembers =  pJSON.results.items                                             //#.(21213.06.3)  
//     var  mMembers =  pJSON.members                                                   // .(21213.06.3 RJS Was pJSON.results.items).(21214.02.9)  
       var  mMembers =  pJSON.members ? pJSON.members : pJSON                           // .(21214.02.9 RAM Was pJSON.members)  

//     var  aHTML    =  mMembers.map(  fmtMember ).join( "\n" )
       var  aHTML    =  mMembers.sort( sortitem).map( fmtMember ).join( "\n" )
//     var  mHTMLs=[];  mMembers.forEach( ( pMember, i ) => { fmtMember( pMember, i ) } ); aHTML = mHTMLs.join( "\n" )
    return  aHTML

//     ---  -------  =  -----------------------------------

  function  fmtMember( pMember, i ) {

       var  aClass   =            i % 2 == 1 ? "row-even" : "row-odd"
//     var  aClass   = "row-" + ( i % 2      ?     "even" :     "odd" )
//     var  aClass   =   (  `class="row-even"` )

       var  aMI      =     pMember.Middlename;  aMI = ( aMI  > "" ) ?   ` ${ aMI.substr(0,1) }. ` : ""      // .(21213.06.4 RJS Beg Change case of column names)
       var  aName    = `${ pMember.FirstName }${aMI} ${ pMember.LastName }`
       var  aPhone   =     pMember.Phone1 + ( pMember.Phone2 > ""   ? `, ${ pMember.Phone2  }` : "" )
            aPhone   =     aPhone == "null" ? "" : aPhone                                                   // .(21213.06.5 RJS Get rid of "null")
       var  aEmail   =     pMember.Email                                                                    // .(21213.06.4 RJS End) 

       var  aRow     = `  <tr Class="${ aClass }" id="R${ `${ i + 1 }`.padStart( 3, "0" ) }">\n`
                     + `  <td class="name"><strong><a href="syschangepassword.js?username=${ aName }">${ aName }</a></strong></td>\n`
                     + `  <td class="email"><small ><a href="mailto:${ aEmail }">Email Address</a></small></td>\n`
                     + `  <td class="phone"><small >${ aPhone }&nbsp;&nbsp;&nbsp;</small></td>\n`
                     + `</tr>\n`

//          mHTMLs.push( aRow )
//          aData    = aHeadRow + aRow
            aData    = aRow
    return  aData
            }   // eof  fmtMember
//     ---  -------  =  ----------------------------------
            }   // eof  fmtMembers
//--------  -------  =  -------------------------------------------------------

   function readFile( aFile ) {                                                         // .(21210.03.1 RAM Beg Add readFile)
       var  pFS      =  require( 'fs' )
       var  aDir     =  __dirname
       var  aText    =  pFS.readFileSync( `${aDir}/${aFile}`, "ASCII" )
    return  aText
            }                                                                           // .(21210.03.1 RAM End)
//--------  -------  =  -------------------------------------------------------

  function  parseJSON(  aFile ) {
       var  aText    =  readFile( aFile )                                               // .(21210.03.2).
            aText    =  aText.replace( //g, "'" )
            aText    =  aText.replace( //g, "'" )

        if (aFile.match( /\.json$/)) {
       var  pJSON    =  JSON.parse( aText )
            }
        if (aFile.match( /\.js$/)) {
            eval( aText )
            }

    return  pJSON
            }   // eof  parseJSON
//--------  -------  =  -------------------------------------------------------

  function  sortitem( a, b ) {                                                          // .(21209.04.1 RJS Beg Add SortItems)
    return (a.LastName + a.FirstName) > (b.LastName + b.FirstName) ? 1 : -1             // .(21213.06.6)
            }                                                                           // .(21209.04.1 RJS End)
//---------------------------------------------------------------------------------------------------

//========================================================================================================= #  ===============================  #
/*\
##SRCE     +====================+===============================================+
##RFILE    +====================+=======+===================+======+=============+
\*/
