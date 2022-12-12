/*\
##=========+====================+================================================+
##RD         IODD-members       | Format IODD Members table
##RFILE    +====================+=======+===============+======+=================+
##FD   IODD.Members.js          |   4757| 12/08/22  9:30|   101| p1.00-21208-0930
##FD   IODD.Members.js          |   5772| 12/09/22 13:15|   105| p1.01-21209-1315
##FD   IODD-members.js          |   6007| 12/10/22 16:40|   119| p1.02-21210.1640
##FD   IODD.Members.mjs         |   8219| 12/11/22 12:45|   145| p1.02-21211-1245       // .(21211.01.1 RAM Change extension from .js to .mjs))
##FD   IODD-members.mjs         |  10160| 12/11/22 15:43|   167| p1-05_21211.1543
##DESC     .--------------------+-------+---------------+------+-----------------+
#            Read Members Data and format it
#
##LIC      .--------------------+----------------------------------------------+
#            Copyright (c) 2022 8020Data-formR * Released under
#            MIT License: http://www.opensource.org/licenses/mit-license.php
#
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
# .(21211.01 12/11/21 RAM 12:45p| Convert into Node Module script, .mjs
# .(21211.02 12/11/21 RAM  3:43p| Add module export

##PRGM     +====================+===============================================+
##ID 69.600. Main               |
##SRCE     +====================+===============================================+
\*/
//========================================================================================================= #  ===============================  #

//import  pModule from "module";                                                            // .(21211.01.2 RAM Beg Import required modules)
//import  pPath   from 'path';                                                              //#.(21211.01.8 RAM Beg I'd rather not use all this)
//import  pURL    from 'url';
//import  pFS     from 'fs';

//     ---  -------  =  -----------------------------------

// const    require  =  pModule.createRequire( import.meta.url );
// const  __filename =  pURL.fileURLToPath(    import.meta.url );
// const  __dirname  =  pPath.dirname( __filename );                                        //#.(21211.01.8 RAM End
     var    mPath    =  import.meta.url.split( "/" )                                        // .(21211.01.8 RAM Beg)
   const  __filename =  mPath[ mPath.length - 1 ]
   const  __dirname  =  mPath.splice( 3, mPath.length - 4 ).join( "/" )                     // .(21211.01.8 RAM End)

//----------------------------------------------------------------------------------------- // ----------- #

     var    aTests='live in Browser'                                                        // .(21211.01.6 RAM var is required in .mjs)
//   var    aTests='test1 in NodeJS'                                                        // .(21211.01.6)

//-----------------------------------------------------------------------------------------

//                  var Members = { format: fmtMembers, parseJSON, readFile, sortitem };
// export           var Members                                                             //#.(21211.02.0 RAM Export using ES6 Module syntax)
// module.exports =     Members                                                             //#.(21211.02.0 RAM Export using old CommonJS syntax)

   export           { fmtMembers, parseJSON, readFile, sortitem };                          // .(21211.02.1 RAM Export seperately named functions or vars)
// import           { fmtMembers, parseJSON, readFile, sortitem } from IODD-Members.mjs;    // .(21211.02.1 RAM Import seperately named functions or vars)


// export           var Members = { format: fmtMembers, parseJSON, readFile, sortitem };    // .(21211.02.2 RAM Export seperately named function or var)
// import               Members                                   from IODD-Members.mjs;    // .(21211.02.2 RAM Import seperately named function or var)

                    var Members = { format: fmtMembers, parseJSON, readFile, sortitem };
// export   default     Members                                                             // .(21211.02.3 RAM Export one anonymous function or var)
// import { default as  Members }                                 from IODD-Members.mjs;    // .(21211.02.3 RAM Import one anonymous function or var)

//---------------------------------------------------------------------------------------------------

// var aHeadRow = `<tr class="head-row"><td>Name</td><td>Email</td><td>Phone / Mobile</td></tr>`

  if ( aTests.match( /test1/ ) ) {

     (async function() {                                                                    // .(21211.01.3 RAM Define async anonymous function)

//     var  pJSON    =  await parseJSON( '../json/IODD-members_u2a.json.js' )               // .(21211.01.4 RAM Add await).(21211.02.1)
//     var  aHTML    =  fmtMembers( pJSON )                                                 // .(21211.02.1)

       var  pJSON    =  await Members.parseJSON( '../json/IODD-members_u2a.json.js' )       // .(21211.01.4 RAM Add await).(21211.02.2)
       var  aHTML    =  Members.format( pJSON )                                             // .(21211.02.2)

            console.log( aHTML )
//          } )                                                                             //#.(21211.01.5 RAM Finish definition)
            } )( )                                                                          // .(21211.01.5 RAM Execute it)
       }

//---------------------------------------------------------------------------------------------------

  function  fmtMembers( pJSON ) {
       var  mMembers =  pJSON.results.items

//     var  aHTML    =  mMembers.map( fmtMember ).join( "\n" )
       var  aHTML    =  mMembers.sort(sortitem).map( fmtMember ).join( "\n" )
//     var  mHTMLs=[];  mMembers.forEach( ( pMember, i ) => { fmtMember( pMember, i ) } ); aHTML = mHTMLs.join( "\n" )
    return  aHTML

//     ---  -------  =  -----------------------------------

  function  fmtMember( pMember, i ) {

       var  aClass   =            i % 2 == 1 ? "row-even" : "row-odd"
//     var  aClass   = "row-" + ( i % 2      ?     "even" :     "odd" )
//     var  aClass   =   (  `class="row-even"` )

       var  aMI      =     pMember.middlename;  aMI = ( aMI  > "" ) ?   ` ${ aMI.substr(0,1) }. ` : ""
       var  aName    = `${ pMember.firstname }${aMI} ${ pMember.lastname }`
       var  aPhone   =     pMember.phone1 + ( pMember.phone2 > ""   ? `, ${ pMember.phone2  }` : "" )
       var  aEmail   =     pMember.email

       var  aRow     = `  <tr Class="${ aClass }" id="R${ `${ i + 1 }`.padStart( 3, "0" ) }">\n`
                     + `  <td class="name"><strong><a href="syschangepassword.js?username=${ aName }">${ aName }</a></strong></td>\n`
                     + `  <td class="email"><small ><a href="mailto:${ aEmail }">Email Address</a></small></td>\n`
                     + `  <td class="phone"><small >${ aPhone }&nbsp;&nbsp;&nbsp;</small></td>\n`
                     + `</tr>\n`

//          mHTMLs.push( aRow )
//          aData    = aHeadRow + aRow
       var  aData    = aRow                                                             // .(21211.01.6)
    return  aData
            }   // eof  fmtMember
//     ---  -------  =  ----------------------------------
            }   // eof  fmtMembers
//--------  -------  =  -------------------------------------------------------

 async function readFile( aFile ) {                                                     // .(21210.03.2 RAM Beg Add readFile).(21211.01.7 RAM Make it async)
//     var  pFS      =  require( 'fs' )
       var  pFS      =  await import( 'node:fs' )
       var  aDir     =  __dirname
       var  aText    =  pFS.readFileSync( `${aDir}/${aFile}`, "ASCII" )
    return  aText
            }                                                                           // .(21210.03.2 RAM End)
//--------  -------  =  -------------------------------------------------------

 async function  parseJSON(  aFile ) {                                                  // .(21211.01.8 RAM Make function async)

        if (typeof(pJSON) != 'object' ) {

       var  aText    =  await readFile( aFile )                                         // .(21210.03.3).(21211.01.9)
            aText    =  aText.replace( //g, "'" )
            aText    =  aText.replace( //g, "'" )

        if (aFile.match( /\.json$/)) {
       var  pJSON    =  JSON.parse( aText )
            }
        if (aFile.match( /\.js$/)) {
            eval( aText )
            }
        }
    return  pJSON
            }   // eof  parseJSON
//--------  -------  =  -------------------------------------------------------

  function  sortitem( a, b ) {                                                          // .(21209.04.1 RJS Beg Add SortItems)
    return (a.lastname + a.firstname) > (b.lastname + b.firstname) ? 1 : -1
            }                                                                           // .(21209.04.1 RJS End)
//---------------------------------------------------------------------------------------------------

//========================================================================================================= #  ===============================  #
/*\
##SRCE     +====================+===============================================+
##RFILE    +====================+=======+===================+======+=============+
\*/
