import pymysql
import openpyxl
# 连接数据库
db = pymysql.connect(user='root',  #用户名
                     password = '000000', #密码
                     database = 'excel', #数据库
                     charset='utf8')

# 获取游标
cur = db.cursor()

# 打开excel文件，获取工作簿对象
wb = openpyxl.load_workbook('表名.xlsx')#相对路径/绝对路径
# 当前活跃的表单
ws = wb.active  # ws = <Worksheet "工作表表名">

# 获取每一行的单元格
# ws.rows:是一个 generator object
row_range = tuple(ws.rows) #((A1单元格,B1单元格,C1单元格...),(A2,B2,C2...),(A3,B3,C3...),...)

# 获取每一列的单元格
# col_range = tuple(ws.columns) #((A1单元格,A2单元格,...An单元格),(B1,B2,...Bn),...)

#第一行为字段名，已经在mysql数据库中创建好了，所以只需要row_range[1:]
for row in row_range[1:]:
    id = row[0].value
    name = row[1].value
    score = row[2].value
    # 执行语句
    sql = "insert into 表名 values(%s,%s,%s);"
    cur.execute(sql,[id,name,score])

try:
    # 同步数据库
    db.commit()
except:
    db.rollback()

# 关闭游标
cur.close()

# 关闭数据库
db.close()