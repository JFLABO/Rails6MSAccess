require 'dbi'
require "nkf"
def to_sjis(str)
  str ? NKF.nkf('-s', str) : ""
end


def to_utf8(str)
  str ? NKF.nkf('-w', str) : ""
end


def create_database_handle
  # プログラムでODBCの接続文字列を指定する場合
  # 上：ローカルドライブ、下：ネットワーク越しの書き方
  # dsn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=c://db/Sample.accdb;)
  # dsn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=//127.0.0.1/db/Sample.accdb;)
    dsn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=j://sample.accdb;)

  DBI.connect(dsn)

  # ODBCデータソースアドミニストレーターで設定してある場合
  # DBI.connect('DBI:ODBC:Sample')
end
  #dsn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=j:\sample.accdb;)
  #sn = %q(DBI:ODBC:Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=//127.0.0.1/db/Sample.accdb;)
  #DBI.connect(dsn)
  
    sql = "SELECT * From lins"
    bind = ""

    dbh = create_database_handle
    #results = dbh.execute(sql)
	#sth = dbh.prepare("SELECT 'Hello, DBI World' AS Message")
	sth = dbh.prepare("SELECT * from links")
	sth.execute()

    #@fuga = ""
    #results.each { |r| @fuga += to_utf8(r[:title]) }
	
while row = sth.fetch do
    print "Message", "\n";
    print "-------------------\n";
    print to_utf8(row["title"])
end
sth.finish
dbh.disconnect
#Microsoft Access Driver (*.mdb, *.accdb) is the name of Microsoft's ODBC driver for Microsoft Access. It is only available for Windows. 

test
