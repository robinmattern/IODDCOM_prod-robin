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
#            parseJSON          | Convert .json text into pJSON object
#              readFile         | Read a file (no error handling)                       // .(21210.03.1)
#            sortitems          | Sort Members by LasName, FirstName                    // .(21209.04.1)
#
##CHGS     .--------------------+----------------------------------------------+
# .(21208.01 12/08/21 RAM  9:30a| Created
# .(21209.04 12/09/21 RJS  1:15p| Add sortitems
# .(21210.03 12/10/21 RAM  4:40p| Add readFile
# .(21213.06 12/13/21 RAM 11:20p| Read db.json.js, not IODD-members_u2a.json.js

##PRGM     +====================+===============================================+
##ID 69.600. Main               |
##SRCE     +====================+===============================================+
\*/
//========================================================================================================= #  ===============================  #

//---------------------------------------------------------------------------------------------------

      aTests='live in Browser'
//    aTests='test1 in NodeJS'

// var aHeadRow = `<tr class="head-row"><td>Name</td><td>Email</td><td>Phone / Mobile</td></tr>`

  if ( aTests.match( /test1/ ) ) {

       var  pJSON    =  parseJSON( '../json/db.json.js' )                               // .(21213.06.1 RJS Was: IODD-members_u2a.json.js) 
       var  aHTML    =  fmtMembers(  pJSON )

            console.log( aHTML )
            }

//---------------------------------------------------------------------------------------------------

  function  fmtMembers( pJSON ) {
       var  mMembers =  pJSON.members                                                   // .(21213.06.2 RJS Was pJSON.results.items)  

//     var  aHTML    =  mMembers.map( fmtMember ).join( "\n" )
       var  aHTML    =  mMembers.sort(sortitem).map( fmtMember ).join( "\n" )
//     var  mHTMLs=[];  mMembers.forEach( ( pMember, i ) => { fmtMember( pMember, i ) } ); aHTML = mHTMLs.join( "\n" )
    return  aHTML

//     ---  -------  =  -----------------------------------

  function  fmtMember( pMember, i ) {

       var  aClass   =            i % 2 == 1 ? "row-even" : "row-odd"
//     var  aClass   = "row-" + ( i % 2      ?     "even" :     "odd" )
//     var  aClass   =   (  `class="row-even"` )

       var  aMI      =     pMember.Middlename;  aMI = ( aMI  > "" ) ?   ` ${ aMI.substr(0,1) }. ` : ""      // .(21213.06.3 RJS Beg Change case of column names)
       var  aName    = `${ pMember.FirstName }${aMI} ${ pMember.LastName }`
       var  aPhone   =     pMember.Phone1 + ( pMember.Phone2 > ""   ? `, ${ pMember.Phone2  }` : "" )
            aPhone   =     aPhone == "null" ? "" : aPhone                                                   // .(21213.06.4 RJS Get rid of "null")
       var  aEmail   =     pMember.Email                                                                    // .(21213.06.3 RJS End) 

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
    return (a.LastName + a.FirstName) > (b.LastName + b.FirstName) ? 1 : -1             // .(21213.06.5)
            }                                                                           // .(21209.04.1 RJS End)
//---------------------------------------------------------------------------------------------------

//========================================================================================================= #  ===============================  #
/*\
##SRCE     +====================+===============================================+
##RFILE    +====================+=======+===================+======+=============+
\*/
