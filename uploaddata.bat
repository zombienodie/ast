cd %1

echo ---------------%time%---------------

git init
git add .
git commit -m "first commit"
git remote add origin %2
git push -u origin master

echo ---------------%time%---------------