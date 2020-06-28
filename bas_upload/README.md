Open Excel. Go to Options Trust center - Trust center Settings - ActiveX settings: Enable all controls; Macro Settings:Enable all macros and trust access to VBA. 
This is usually done only once after a fresh Office installation. 

Record Macro
Alt+F11
Click on Modules-your module(macro) export file as .bas
*This .bas file should be generated automatically by Julia, along with the .xlsx you want to process with the macro.

Edit path to .xlsx and .bas in import_bas.vbs.
Double click import_bas.vbs.
* This can be automatically called from Julia with:
# call vbs
mycommand = `WScript.exe "import_bas.vbs"`
wait(run(mycommand))
