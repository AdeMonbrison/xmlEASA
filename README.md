# xmlEASA
A small VBA code to transcript EASA pdf into excel sheets


How to use it //


1- Use Foxit Reader to transfer a EASA PDF into excel file
2- Use the "correction" sub to clean up the excel sheet. This sub has been created following several test and should cover any issues
a) if there is any issues, add code at the BOTTOM of the correction sub
b) aim of this program is to merge separated content : please keep it that way and always try to merge content in the same cell for other sub to work correctly
3) launch the getlistofcsr (create a sheet named "CS")
4) launch the getlistofreg (create a sheet named "Reg")
5) launch the getlistofguid (create a sheet named "GM")

Optional : add new rules to make it nicer ::
![image](https://user-images.githubusercontent.com/114153756/191734454-b71010d0-dfac-41aa-bd06-54b737e56411.png)
