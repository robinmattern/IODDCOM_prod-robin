
   import   util            from 'util'; var inspect = function(p) { console.log( util.inspect(p, { depth: 9 } ) ) }
   import { writeFileSync } from 'fs';
      var   aURI        =   import.meta.url
      var __filename    =   aURI.replace( /^.+\//, "" )
      var __dirname     =   aURI.replace( `/${__filename}`, "" ).replace( "file:///", "" )

// ------   ------------------- =  --------------------------------------------
   import   README_JSON      from './README_json.mjs';

      var  mProjs =  README_JSON()

      var  nPrt   =  2    // 1) to stdout, 2) to file, 3) both
      var  nCmd   =  1    // 1) README.md, 2) Index.html, 3) Apps, 4) Links

      var  aCmd   =  [ , 'ReadMe', 'Index', 'Apps', 'Links' ][ nCmd ]

       if (aCmd  == 'ReadMe') { savReadMe( mProjs, nPrt, 'README.md'  ) }
       if (aCmd  == 'Index' ) { savIndex(  mProjs, nPrt, 'index.html' ) }
       if (aCmd  == 'Apps'  ) { shoApps(   mProjs, 1 ) }
       if (aCmd  == 'Links' ) { shoLinks(  mProjs, 1 ) }

// ------  -------------------- =  --------------------------------------------

 function  getTopBot_4README( ) {

      var  aTop = `
<html>
 <body>
  <style><!--
    details > ul > li            { margin-top:-10px !important; margin-bottom:20px !important; }
    details > ul > li > p        { color: #810d0d; padding-left: 20px; margin-top:-17px !important; text-indent: -20px; line-height: 22px !important; }
    details > ul > li:last-child { display: none; }                    					            /* .(21218.02.1 RAM Don't display MT last child) */
    code                         { color: black; font-size: 12px; margin: 0px 0px 0px 16px !important; padding-bottom: 0px; }   /* .(21218.02.3 RAM) */
  --></style>

  <div style="margin-left:25px;">

#`
      var  aBot = `
  </div>
 </body>
</html>\n`

   return [ aTop, aBot ]

            }  // eof getTopBot_4README
//  -----   ------------------- =  ----------------------------------

 function  savReadMe( mProjs, nPrt, aFile ) {

      var [ aTop, aBot ] = getTopBot_4README()

//    var   mProjs       =  mProjs.map( fmtProj );              // inspect( mProjs )
      var   aProjs       =  mProjs.map( fmtProj ).join( '\n' )

      var   aMarkdown    =  aTop + aProjs + aBot;           // console.log( aHTML )

            prtOut( aFile, aMarkdown, nPrt )

//  -----   ------------------- =  ----------------------------------

 function  fmtProj( pProj ) {
      var  aHTML = `# <u>${pProj.proj}</u>\n  `
        +  `<h2 style="font-size:24px; margin: -18px 0px 15px 12px;">(${pProj.name})</h2>\n`
        +  (pProj.stages ? pProj.stages.map( fmtStage ).join( '\n  ' ) : '' ) + '\n'
   return   aHTML
            }
//  -----   ------------------- =  ----------------------------------

 function  fmtStage( pStage ) {
   return  `<details><summary><b style="font-size:24px;">${pStage.stage}</b></summary>\n`
        +   pStage.apps.map(  fmtApp ).join( '\n  ' ) + '\n  '
        +  '-\n\n'                                                                     // #.(21218.02.2 RAM Create an MT last child)
        +  `</details>`
            }
//  -----   ------------------- =  ----------------------------------

 function  fmtApp( pApp ) {
//         pApp.txt = pApp.txt.replace( /  - /g, "  &nbsp;&nbsp;&nbsp; &bull;&nbsp; " )
//         pApp.txt = pApp.txt.replace( /\n +\n/g,  "\n                 <blankline>  \n"  )
           pApp.txt = pApp.txt.replace( /  \$ (.+)\n/g,  " `$ $1 `  \n"       )     

   return  `- ### [${pApp.app}](${pApp.url})\n    `      
        +  `${pApp.txt }\n`
            }
//  -----   ------------------- =  ----------------------------------
        }  // eof savREADME
// ------  -------------------- =  --------------------------------------------

 function  getTopBot_4Index( ) {

      var  aTop =
`<html>
  <head>
    <style>
     :root { --color: #084074; }  /* #084074; #DEE6ED; */
      main { margin-left: 20px; padding: 0px 0px 40px 40px; border: 2px solid var(--color); width: 600px;}
      h1   {                    font-size: 42px; color: var(--color); margin-bottom: 5px; }
      h2   { margin-top: -10px; font-size: 36px; color: var(--color); margin-bottom: 5px; margin-left: 20px; }
      details > summary {       font-size: 28px; color: var(--color); font-weight:  bold; }
      h3   { margin-top:  10px; margin-block-start: -10px; margin-block-end: -1px; }
/*    desc { margin-top: -15px; margin-block-start:   0px; margin-bottom:    20px; }  */
      li   { margin-bottom: 15px; }
    </style>
  </head>
  <body>
    <main>\n`

      var  aBot =
`
    </main>
  </body>
</html>\n`

   return [ aTop, aBot ]

            }  // eof getTopBot_4Index
//  -----   ------------------- =  ----------------------------------

 function   savIndex( mProjs, nPrt, aFile ) {

      var [ aTop, aBot ] = getTopBot_4Index()

//    var   mProjs = mProjs.map( fmtProj )               inspect( mProjs )
      var   aProjs = mProjs.map( fmtProj ).join( '\n' )

      var   aHTML  =  aTop + aProjs + aBot;           // console.log( aHTML )

            prtOut( aFile, aHTML, nPrt )

//  -----   ------------------- =  ----------------------------------

 function   fmtProj( pProj ) {
   return  `<h1><u>${pProj.proj}</u></h1>\n`
        +  `<h2>(${pProj.name})</h2>\n`
        +   pProj.stages.map( fmtStage ).join( '\n  ' ) + '\n'
            }
//  -----   ------------------- =  ----------------------------------

 function   fmtStage( pStage ) {
   return  `<details><summary>${pStage.stage}</summary><ul>\n`
        +   pStage.apps.map(  fmtApp ).join( '\n    ' ) + '\n  '
        +  `</ul></details>`
            }
//  -----   ------------------- =  ----------------------------------

 function   fmtApp( pApp ) {
   return  `<li><h3><a href="${pApp.url}" target="_blank">${pApp.app}</a></h3>\n    `
        +  `<desc>${pApp.txt}</desc></li>\n`
            }
//  -----   ------------------- =  ----------------------------------
        }  //  eof savIndex
// ------  -------------------- =  --------------------------------------------

 function   shoApps( mProjs, nPrt, aFile ) {

      var   aList = mProjs.map( fmtProj  ).join( '\n' )

            prtOut( aFile, aList, nPrt )

//  -----   ------------------- =  ----------------------------------

 function   fmtProj( pProj ) {
   return   pProj.stages.map(   fmtStage ).filter( notMT ).join( '\n' )
            }
//  -----   ------------------- =  ----------------------------------

 function   fmtStage( pStage ) {
   return   pStage.apps.map(    fmtApp   ).filter( notMT ).join( '\n' )
            }
//  -----   ------------------- =  ----------------------------------

 function   fmtApp( pApp ) {
   return  `${pApp.url}`
            }
//  -----   ------------------- =  ----------------------------------
        }  //  eof shoApps
// ------  -------------------- =  --------------------------------------------

 function   shoLinks( mProjs ) {

      var   aHTML = `
                <!-- ------- --------------------------------------------------  ------------------ --------------- -->`

         + ` ${ mProjs.map( pProj => `
                <div class="title"><u><b>${ pProj.proj }</b></u></div>`

            + ` ${ pProj.stages.map( pStage =>

   `         ` + ` ${ pStage.apps.map( pApp =>
   `                  <li><a href="${ pProj.url + pApp.url }">${ pApp.app }</a></li>`
                      ).join( '\n' )
                      } `
                   ).join( '\n' )
                   } `
               ).join( '\n' )
               } `

            prtOut( '', aHTML, 1 )

            }  //  eof shoLinks
// ------  -------------------- =  --------------------------------------------

 function   shoLinks2( mProjs ) {

      var   aHTML = `
                <!-- ------- --------------------------------------------------  ------------------ --------------- -->`

         + ` ${ mProjs.map( pProj => `
                <div class="title"><u><b>${ pProj.proj }</b></u></div>`

            + ` ${ pProj.stages.map( pStage => `
                  <div class="subtitle">${ pStage.stage }</div>
 `
               + ` ${ pStage.apps.map( pApp =>
 `                    <li><a href="${ pProj.url + pApp.url }">${ pApp.app }</a></li>`
                      ).join( '\n' )
                      } `
                   ).join( '\n' )
                   } `
               ).join( '\n' )
               } `

            prtOut( '', aHTML, 1 )

            }  //  eof shoLinks2
// ------  -------------------- =  --------------------------------------------

 function   shoLinks3( mProjs ) {

      var   aHTML = `
                <!-- ------- --------------------------------------------------  ------------------ --------------- -->`

         + ` ${ mProjs.map( pProj => `
                <div class="title"><u><b>${ pProj.proj }</b></u></div>`

            + ` ${ pProj.stages.map( pStage =>
`                  <div class="subtitle">${ pStage.stage }</div>`

               + ` ${ pStage.apps.map( pApp =>
 `                    <li><a href="${ pProj.url + pApp.url }">${ pApp.app }</a></li>`
                      ).join( '\n' )
                      } `
                   ).join( '\n' )
                   } `
               ).join( '\n' )
               } `

            prtOut( '', aHTML, 1 )

            }  //  eof showLinks3
// ------  -------------------- =  --------------------------------------------

 function   notMT( aRow ) {
   return   aRow != ""
            }

 function   prtOut( aFile, aText, nPrt ) {

        if (nPrt == 1 || nPrt == 3) {  // to stdout 
            console.log( aText )
            }  

        if (nPrt == 2 || nPrt == 3) {  // to file
            aFile = aFile.match( /[\/]/ ) ? aFile : `${__dirname}/${aFile}` 
            writeFileSync( aFile, aText )
            console.log( `* Saved ${aFile}` )
            }
        }   
// ------  -------------------- =  --------------------------------------------



