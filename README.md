# change_header
is a program that change the headers of the file(txt,xml,xlsx,xls) to new headers
the program works with capital and small letters, and you can choose if you want the new data in the same worksheet or in a new worksheet

the headers map of old column names to new columns names exict in config file... 
<config>
  <headers>
    <header existingHeader ="ColA"  newHeader="ColAA" />
    <header existingHeader = "ColC"  newHeader="ColCC" />
    <header existingHeader = "ColB"  newHeader="ColBB" />
  </headers>
</config>
