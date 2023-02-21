@echo off
set /p input="url D/F> " 
set /p input1="url GITHUB> " 

cd %input%
echo ---------------%time%---------------

git init
git add .
git commit -m "first commit"
git remote add origin %input1%
git push -u origin master

git fetch origin master:tmp

git push -u origin master

echo ---------------%time%---------------
pause