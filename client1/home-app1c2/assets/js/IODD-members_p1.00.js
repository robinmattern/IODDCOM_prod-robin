/*\
##=========+====================+================================================+
##RD         IODD-members       | Format IODD Members table
##RFILE    +====================+=======+===============+======+=================+
##FD   IODD.Members.js          |   4757| 12/08/22  9:30|   101| p1.00-21208-0930
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
#
##CHGS     .--------------------+----------------------------------------------+
# .(21208.01 12/08/21 RAM  9:30a| Created

##PRGM     +====================+===============================================+
##ID 69.600. Main               |
##SRCE     +====================+===============================================+
\*/
//========================================================================================================= #  ===============================  #

//---------------------------------------------------------------------------------------------------

       aTests='live in Browser'
//     aTests='test1 in NodeJS'

// var aHeadRow = `<tr class="head-row"><td>Name</td><td>Email</td><td>Phone / Mobile</td></tr>`

  if ( aTests.match( /test1/ ) ) {

       var  pJSON    =  parseJSON( '../json/IODD-members_u2a.json.js' )
       var  aHTML    =  fmtMembers(  pJSON )

            console.log( aHTML )
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
            aData    = aRow
    return  aData
            }   // eof  fmtMember
//     ---  -------  =  ----------------------------------
            }   // eof  fmtMembers
//--------  -------  =  -------------------------------------------------------

  function  parseJSON(  aFile ) {
       var  pFS      =  require( 'fs' )
       var  aDir     =  __dirname
       var  aText    =  pFS.readFileSync( `${aDir}/${aFile}`, "ASCII" )
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

//========================================================================================================= #  ===============================  #
/*\
##SRCE     +====================+===============================================+
##RFILE    +====================+=======+===================+======+=============+
\*/
