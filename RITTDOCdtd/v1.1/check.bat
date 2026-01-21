REM checker for dtd
c:
cd C:\projects\Rittenhouse\dtds\dtd
nsgmls.exe -sv -wxml -D"." -D".\ent" -c catalog.xml -E1000 -ftmp.txt  -c "C:\Program Files\jclark\nsgmls\pubtext\xml.soc" -wall demo2.xml
fc tmp.txt tmp1.txt > diff.txt
pause
cp -f tmp1.txt tmp2.txt
cp -f tmp.txt tmp1.txt
