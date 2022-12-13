#!/bin/sh

  mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql  --execute "SELECT * FROM members" | awk 'NR >= 4' >members-mysql.json

  echo " pFS   = require( 'fs' )"                                    >members-mysql.njs
  echo " aText = pFS.readFileSync( 'members-mysql.json', 'ASCII' )" >>members-mysql.njs
  echo " aText = aText.replace( //g, '\'' )"                       >>members-mysql.njs
  echo " aText = aText.replace( //g, '\'' )"                       >>members-mysql.njs
  echo " pJSON = JSON.parse( aText )"                               >>members-mysql.njs
  echo " console.dir( pJSON, { depth: 9 } )"                        >>members-mysql.njs

  node members-mysql.njs




