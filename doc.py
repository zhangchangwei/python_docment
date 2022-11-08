import os, datetime
from docxtpl import DocxTemplate


system_type = '系统'
templat_path = os.path.abspath(".")

# 定义doc表格数据
table_data = [{
    'province': '山东',
    'cost': 1000,
    'avg': 500,
    'comment': ''
}, {
    'province': '江苏',
    'cost': 2000,
    'avg': 1600,
    'comment': ''
}, {
    'province': '北京',
    'cost': 1200,
    'avg': 700,
    'comment': ''
}]

# 指定doc模板
doc_rb = DocxTemplate(templat_path+'\/templat\日报.docx')

# 向doc文档传参写入数据
context = {'SystemType': system_type, 'Cost': 20000, 'table_data': table_data}

# 获取当前日期
now_time = datetime.datetime.now().strftime('%Y-%m-%d')

# 同步数据至doc
doc_rb.render(context)

filename='{}\daily\{}-{}日报.docx'.format(templat_path, now_time, system_type)

# 保存数据
doc_rb.save(filename)
