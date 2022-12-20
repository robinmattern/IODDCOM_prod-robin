//---------------------------------------------------------------------------------------------------
 
aTests='live in Browser'
//     aTests='test1 in NodeJS'

  if ( aTests.match( /test1/ ) ) { 

       var  pJSON    =  parseJSON( '../json/db_v2.json.js' )
       var  aHTML    =  fmtMembersProjects(  pJSON )

            console.log( aHTML )
            }

//---------------------------------------------------------------------------------------------------

function  fmtMembersProjectsLaptop( pJSON ) {     
       var  mMembers_projects =  pJSON.members_projects

//       var  aHTML    =  mMembers_projects.map( fmtMember ).join( "\n" )
//       var  aHTML    =  mMembers_projects.sort(sortitem).map( fmtMemberProject ).join( "\n" )
         var  aHTML    =  mMembers_projects.map( fmtMemberProjectLaptop ).join( "\n" )
//       var  mHTMLs=[];  mMembers_projects.forEach( ( pMember, i ) => { fmtMember( pMember, i ) } ); aHTML = mHTMLs.join( "\n" ) 
    return  aHTML
            
//     ---  -------  =  -----------------------------------

  function  fmtMemberProjectLaptop( pMember, i ) {

       var  aClass = i % 2 == 1 ? "row-even" : "row-odd"
//     var  aClass = "class=row-" + ( i % 2 ? "even" : "odd" )
//     var  aClass   =   (  `class="row-even"` )

       var  aMI      =     pMember.Middlename;  aMI = ( aMI  > "" ) ?   ` ${ aMI.substr(0,1) }. ` : ""
       var  aNameBR    = `${ pMember.FirstName }${aMI} <br> ${ pMember.LastName }`
       var  aName    =   `${ pMember.FirstName }${aMI} ${ pMember.LastName }`
//     var  aPhone   =     pMember.Phone1 + ( pMember.Phone2 > ""   ? `, ${ pMember.Phone2  }` : "" )
//          aPhone   =     aPhone != "null" ? aPhone : ""
       var  aPhone1  =     pMember.Phone1
            aPhone1  =     aPhone1 != null ? aPhone1 : ""
       var  aPhone2  =     pMember.Phone2
            aPhone2  =     aPhone2 != null ? aPhone2 : ""
       var  aEmail   =     pMember.Email
       var  aAddress =     pMember.Address1
       var  aCity    =     pMember.City
       var  aState   =     pMember.State
       var  aZip     =     pMember.Zip
       var  aProjName =    pMember.ProjectName
       var  aClient  =     pMember.Client
       var  aClientWeb =   pMember.ClientWeb
       var  aLocation =    pMember.Location
       var  aProjType =    pMember.ProjectType
       var  aStyle    =    pMember.Style
       var  aRole    =     pMember.Role
       var  aDates   =     pMember.Dates
       var  aDuration =    pMember.Duration
       var  aBio     =     pMember.Bio
       var  aDescription = pMember.Description
       var  aSort    =     pMember.Sort
       var  aRow     = 

       `  <tr Class="${ aClass }" id="R${ `${ i + 1 }`.padStart( 3, "0" ) }">\n`

+ `       <div class="members-projects">\n`
+ `       <table class="main-table">\n`
+ `       <tbody><tr>\n`
+ `          <td colspan="3" valign="top" align="center">\n`
+ `           <table class="${ aStyle}" cellspacing="0" cellpadding="10" align="left">\n`
+ `             <tbody>\n`
+ `            <tr><td colspan="4" class="member-line"><hr></td></tr>\n`
+ `             <tr>\n`
+ `               <td class="members-name">${ aNameBR }</td>\n`
+ `               <td class="members-resume" rowspan="2" valign="top"><span class="member-bio-tag">Bio:</span><br><br>${ aBio }</td>\n`
+ `             </tr>\n`
+ `             <tr>\n`
+ `               <td>\n`
+ `                 <table class="members-address" width="100%" cellspacing="0" cellpadding="0" align="center">\n`
+ `                   <tbody>\n`
+ `                   <tr><td>${ aAddress }</td></tr>\n`
+ `                   <tr><td>${ aCity }, ${ aState } ${ aZip }</td></tr>\n`
+ `                   <tr><td>${ aPhone1}</td></tr>\n`
+ `                   <tr><td><a href="mailto:${ aEmail }" title="Email ${ aName}">Email</a> <a href="mailto:${ aEmail }"><img class="email-image" src="../assets/images/Email.gif"></a></td></tr>\n`
+ `                 </tbody></table>\n`
+ `               </td>\n`
+ `             </tr>\n`
+ `             <tr><td class="project-title" colspan="4">Project(s):&nbsp;<hr></td></tr>\n`
+ `           </tbody>\n`
+ `           </table>\n`
+ `         <tr>\n`
+ `            <td> &nbsp; </td>\n`
+ `         </tr>\n`          
+ `       <tr>\n`
+ `       <td width="45%" valign="top">\n`
+ `         <table class="project-summary">\n`
+ `           <tbody>\n`
+ `                <tr><td valign="top">Project: ${ aProjName }</td></tr>\n`
+ `             <tr><td valign="top">Owner: ${ aClient }</td></tr>\n`                     
+ `             <tr><td valign="top">Owner Web: <a href="${ aClientWeb }">${ aClientWeb }</a></td></tr>\n`
+ `             <tr><td valign="top">Dates: ${ aDates }</td></tr>\n`        
+ `             <tr><td valign="top">Location: ${ aLocation }</td></tr>\n`
+ `             <tr><td valign="top">Role: ${ aRole }</td></tr>\n`
+ `             <tr><td valign="top">Duration: ${ aDuration }</td></tr>\n`
+ `             <tr><td valign="top">Type: ${ aProjType }</td></tr>\n`
+ `           </tbody>\n`
+ `         </table>\n`
+ `       </td>\n`		   
+ `       <td class="project-details" >${ aDescription }</td>\n`
+ `       </tr>\n`
+ `             <tr><td colspan='4' class="project-line"><hr></td></tr>\n`
+ `       </tbody>\n`
+ `       </div>\n`

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
     var aSort = String(51 - a.Sort).padStart(2,"0")
	return (a.LastName + a.FirstName + aSort) > (b.LastName + b.FirstName + aSort) ? 1 : -1
	}
//---------------------------------------------------------------------------------------------------

