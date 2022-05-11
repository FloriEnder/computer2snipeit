Die Software liest "wichtige" Informationen aus Windows Computer aus und trägt diese, über die API, in Snipe-IT ein.

Die Config.xml muss im Root-Verzeichniss der .EXE liegen.
In der Configdatei:
- logpath -> letzter \ muss gesetzt werden. (Bsp c:\test\ oder .\)
- APIPath -> Pfad zur lokalen Snipe-IT Seite. (Bsp https://snipe-it.test.de/api/v1/) Wichtig ist das nur bis zu "api/v1/" der Pfad angegeben wird.
- APIKey -> Kann in der Web-Oberfläche erzeugt werden. (https://snipe-it.readme.io/reference/generating-api-tokens)
- APIlocation_id -> ID des Standortes zudem die Geräte hinzugefügt werden sollen.
- APIcompany_id -> ID der Firma zudem die Geräte hinzugefügt werden sollen.
- APIcategorieid -> IDs der Kategorien die vorher angelegt werden müssen. Die Reinfolge ist wichtig! (CPU,GPU,HDD/SSD,RAM,Mainboard,Computer/Notebook) BSP eingabe(1,2,3,4,5,6). 1 = CPU, 2 = GPU, etc
- APIfieldsetid -> ID des Feldsatzes für Computer. Der Feldsatz und die dazugehörigen benutzerdefienierten Felder müssen vorher angelegt werden. Anzulegende Felder sind Virenschutz,Computername,Letzter Benutzer,Letzter Durchlauf,BIOS Version und Mac Adresse Computer. Der "DB_Field" für die einzelnen Benutzerdefinierten Felder muss in der Programm.vb in den Zeilen 778 und 797 ersetzt werden.
