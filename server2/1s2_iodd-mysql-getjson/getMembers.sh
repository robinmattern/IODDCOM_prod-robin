#!/bin/sh

# mysqlsh -u root -pFormR\!1234 --json=raw                  --sql -f getMembers.sql
# mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql -f getMembers.sql
# mysqlsh -u root -pFormR\!1234 --json=raw    --database=io --js  -f getMembers.sjs
  mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql  --execute "SELECT * FROM members"                # .(21213.02.1 | Create db.json using mysqlsh)
  mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql  --execute "SELECT * FROM projects"               # .(21213.02.2)
  mysqlsh -u root -pFormR\!1234 --json=pretty --database=io --sql  --execute "SELECT * FROM members_projects_view"  # .(21213.02.3)


