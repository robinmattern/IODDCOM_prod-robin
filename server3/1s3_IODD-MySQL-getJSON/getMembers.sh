#!/bin/sh

# mysqlsh -u root -pFormR\!1234 --json=raw                  --sql -f getMembers.sql
# mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql -f getMembers.sql
# mysqlsh -u root -pFormR\!1234 --json=raw    --database=io --js  -f getMembers.sjs
  mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql  --execute "SELECT * FROM members"


