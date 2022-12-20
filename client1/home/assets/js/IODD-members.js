//---------------------------------------------------------------------------------------------------
 
       aTests='live in Browser'
//     aTests='test1 in NodeJS'

  var aHeadRow = `<tr class="head-row"><td>Name</td><td>Email</td><td>Phone / Mobile</td></tr>`

  if ( aTests.match( /test1/ ) ) { 

       var  pJSON    =  parseJSON( '../json/IODD-members_u2a.json.js' )
       var  aHTML    =  fmtMembers(  pJSON )

            console.log( aHTML )
            }

//---------------------------------------------------------------------------------------------------

function  fmtMembers( pJSON ) {     
       var  mMembers =  pJSON.members

//       var  aHTML    =  mMembers.map( fmtMember ).join( "\n" )
         var  aHTML    =  mMembers.sort(sortitem).map( fmtMember ).join( "\n" )
//     var  mHTMLs=[];  mMembers.forEach( ( pMember, i ) => { fmtMember( pMember, i ) } ); aHTML = mHTMLs.join( "\n" ) 
    return  aHTML
            
//     ---  -------  =  -----------------------------------

  function  fmtMember( pMember, i ) {

       var  aClass = i % 2 == 1 ? "row-even" : "row-odd"
//     var  aClass = "class=row-" + ( i % 2 ? "even" : "odd" )
//     var  aClass   =   (  `class="row-even"` )

       var  aMI      =     pMember.Middlename;  aMI = ( aMI  > "" ) ?   ` ${ aMI.substr(0,1) }. ` : ""
       var  aName    = `${ pMember.FirstName }${aMI} ${ pMember.LastName }`
//     var  aPhone   =     pMember.Phone1 + ( pMember.Phone2 > ""   ? `, ${ pMember.Phone2  }` : "" )
//          aPhone   =     aPhone != "null" ? aPhone : ""
       var  aPhone1  =     pMember.Phone1
            aPhone1  =     aPhone1 != null ? aPhone1 : ""
       var  aPhone2  =     pMember.Phone2
            aPhone2  =     aPhone2 != null ? aPhone2 : ""
       var  aEmail   =     pMember.Email

       var  aRow     = `  <tr Class="${ aClass }" id="R${ `${ i + 1 }`.padStart( 3, "0" ) }">\n`
               + `  <td class="name"><a href="syschangepassword.js?username=${ aName }">${ aName }</a></td>\n`
               + `  <td class="email"><a href="mailto:${ aEmail }"><img class="email-image" src="../assets/images/Email.gif"></a></td>\n`
               + `  <td class="phone1">${ aPhone1 }</td>\n`
               + `  <td class="phone2">${ aPhone2}</td>\n`
               + `</tr>\n`

//          mHTMLs.push( aRow )                  
     //aData = aHeadRow + aRow
     aData = aRow
     return aData
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

function sortitem(a,b) {
	return (a.LastName + a.FirstName) > (b.LastName + b.FirstName) ? 1 : -1
	}
//---------------------------------------------------------------------------------------------------

