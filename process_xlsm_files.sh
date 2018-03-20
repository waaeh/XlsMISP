# Script requires the following packages installed
# sudo apt-get install pip
# sudo -H pip install -U oletools
# sudo apt-get install unzip
# sudo apt-get install libxml2-utils

#cd [YOUR LOCAL PATH]

for f in XlsMISP*.xlsm; do
  rm -rf "./$f.src/"
  unzip -u "$f" -d "./$f.src/"
 
  for x in $(find "./$f.src/" -name '*.xml'); do
 	xmllint --format "$x" --output "$x.xmllint"
	mv "$x.xmllint" "$x"
  done

  
  olevba "$f" -c > "./$f.src/vbaProject.txt"
done
