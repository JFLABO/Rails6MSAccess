require 'dbi'
  # dsn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=c://db/Sample.accdb;)
  dsn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=//127.0.0.1/db/Sample.accdb;)
  DBI.connect(dsn)

    sql = to_sjis("SELECT Piyo From Hoge WHERE Fuga = ?")
    bind = to_sjis(params[:hoge].to_s)

    dbh = create_database_handle
    results = dbh.execute(sql, bind)

    @fuga = ""
    results.each { |r| @fuga += to_utf8(r[:Piyo]) }
