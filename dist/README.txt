copy this folder dist to C: drive
copy the excel contains Parcel Number list to this folder and rename it to "input.xlsx"
this will be the input file
make sure the first row #1 contains the column headings
run the command taxlien
if command completed without errors, then check the output file "taxlien.xlsx"

changes 20230604
- add error handling for NoneType attribute error (ie if webpage return empty page)
- add random delay time to emulate variable wait time between process each parcel