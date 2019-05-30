# IPI
Although most solutions i offer are hosted at http://www.ipi-international.co.uk, until I move that site to https I'll make good use of sharing here!

IPI Database Copy and Search.hta is a hypertext application written mostly in VbScript with a little JavaScript and PHP thrown in for good measure. It utilises ADODB connections and recordsets to Source data from Microsoft Excel; Access; SQL Server CE; SQL Server LocalDB; and SQL Server. Although PHP Object Oriented MYSQLI is used to source and alter MySQL, and PHP PDO for SQLite databases, a Wscript Shell with a mix of ADODB Streams and Recordsets are used to structure the data for sourcing functions. The SQLite database functions also can utilise a separate wrapper as described in the accompanying readme pdf.
