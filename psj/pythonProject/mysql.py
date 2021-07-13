import pymysql

def selectMember():
    conn = pymysql.connect(host='192.168.0.104',user='root',password='1234',db='hwp',charset='utf8')
    try:
        sql = "select * from member where bn =%s"
        cursor = conn.cursor()
        cursor.execute(sql, ('2'))
        result=cursor.fetchall()
        # print(result)
        for row in result:
            print(row[0],row[1],row[2])
    except:
        pass
    finally:
        conn.close()


selectMember()