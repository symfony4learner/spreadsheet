A simple symfony 4 project to demonstrate how to use PHPSpreadsheet.

I used the bundle :
https://github.com/yectep/phpspreadsheet-bundle
to do the stuff. 

This is symfony 4, don't forget to add this last in '/vendor/yectep/phpspreadsheet-bundle/src/Resources/config/services.yml':

services:
    phpoffice.spreadsheet:
        class: '%phpofficebundle_spreadsheet_class%'
        public: true

