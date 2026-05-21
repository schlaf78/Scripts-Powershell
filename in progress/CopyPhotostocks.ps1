#Script is suitable to copy all files across  HDD\Servers with NTFS permissions saving.


robocopy "D:\photostock-brands\FOLDER NAME" "H:\FOLDER NAME" *.* /TEE /E /NP /V /DCOPY:T /COPY:DATSOU /PURGE /MIR /R:1 /W:1 /ZB /mt:16 /unilog:H:\SourceFolder-20250227-RobocopyResults.log


#Used keys:
#/TEE
#/E
#/NP
#/V
#/DCOPY:T
#/COPY:DATSOU
#/PURGE
#/MIR
#/R:1
#/W:1
#/ZB
#/mt:16