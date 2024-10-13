dbPath=$1
dbLockPath=$(echo $dbPath | sed 's/mdb/ldb/')

if [ -e "$dbLockPath" ]; then
    echo "DB is locked, try again later"
    exit 1
fi

mdb-export "$dbPath" MEMBERS > members.csv