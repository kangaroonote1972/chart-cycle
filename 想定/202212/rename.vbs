dim fso
set fso = CreateObject("Scripting.FileSystemObject")

dim folder
set folder = fso.GetFolder(".")

' ƒtƒ@ƒCƒ‹ˆê——
dim f, fn
for each f in folder.files
    if Right(f.name, 4) = ".png" then
    	set fn = fso.GetFile(f.name)
    	fn.name = Mid(f.name, 1, 6) & ".png"
    end if
next 

