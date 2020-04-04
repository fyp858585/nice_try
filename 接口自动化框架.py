import xlrd
import requests
from xlutils.copy import copy
from urllib import parse


excel_path = r'D:\Soft\VScode\Python\接口自动化用例.xls'
excelfile = xlrd.open_workbook(excel_path,formatting_info=True)
sheet_1 = excelfile.sheets()[0]
# cols 是列
#cols = sheet_1.col_values(0)
# rows 是行
#rows = sheet_1.row_values(0)

# cell 单元格查找，先列再行，7C也就是6,2
# cell_7C=sheet_1.cell(6,2).value

# 全局变量
answer = []
pass_or_false = []
body = {}
headers = {"User-Agent":
        "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/45.0.2427.7 Safari/537.36"
}

# 得到接口信息类
class Api_get_result():
    # 初始化属性
    # url,phone,key,parameter_phone,parameter_key
    def __init__(self,url,parameter_keys,parameter_values):
        self.url = url
        for i in range(len(parameter_keys)):
            body[parameter_keys[i]] = parameter_values[i]

    # 判断是get请求还是post请求
    def way_to_requests(self,way):
        if way == 'get':
            return True
        elif way == 'post':
            return False
        else:
            return None
    
    # 发送 post请求
    def api_post(self):
        get_url = self.url + '?' + parse.urlencode(body)
        res = requests.post(get_url,headers=headers)
        # res = requests.post(self.url,data = body,headers=headers)
        return res
    
    # 发送 get请求
    def api_get(self):
        get_url = self.url + '?' + parse.urlencode(body)
        res = requests.get(get_url,headers=headers)
        return res

# 处理Excel类
class Dispose_excel:
    # 得到case多少列函数
    def get_case_count(self):
        # 获取一共有多少列
        rows_count = sheet_1.nrows
        # 去掉表头 一共有多少个case
        real_count = rows_count - 6
        return real_count
    
    # 得到接口URL地址
    def get_url(self):
        url = rows[2]
        return url

    # 得到请求方式
    def get_way(self):
        way = rows[3]
        return way


    # 得到请求参数
    def requests_parameter(self):
        requests_parameter_cell = rows[4]
        return requests_parameter_cell
    
    # 得到正确状态码
    def should_status_code(self):
        status_code = rows[5]
        return status_code
    
    # 得到预期结果
    def expected_result(self):
        result = rows[6]
        return result
    
    # 写入 实际结果 和pass false
    def write_result(self):
        answer_excel = copy(excelfile)
        workSheet = answer_excel.get_sheet(0)
        start = 6
        for i in range(len(answer)):
            workSheet.write(start,7,answer[i])
            workSheet.write(start,8,pass_or_false[i])
            start += 1
        answer_excel.save('接口测试结果.xls')
        

   
if __name__ == "__main__":
    # 处理excel初始化函数
    document_data = Dispose_excel()
    # 得到一共有多少case
    count = document_data.get_case_count()
    # 遍历
    for i in range(6,6+count):
        # 得到各个参数
        rows = sheet_1.row_values(i)
        url = document_data.get_url()
        way = document_data.get_way()
        res_parameter = document_data.requests_parameter()
        status = document_data.should_status_code()
        expect_result = document_data.expected_result()
        # print(url,way,res_parameter,status,expect_result)
        # 处理请求参数，变为字典
        parameter_dic = eval(res_parameter)
        num = len(parameter_dic)
        # 得到具体输入 api_get类里的参数
        parameter_dic_keys = []
        parameter_dic_values = []
        for i in range(num):
            parameter_dic_keys.append(list(parameter_dic.keys())[i])
            parameter_dic_values.append(list(parameter_dic.values())[i])

        # 调用api_get类
        api = Api_get_result(url,parameter_dic_keys,parameter_dic_values)
        # 判断请求方式，发起请求 且得到结果
        if way == 'get':
            result = api.api_get().text
        elif way == 'post':
            result = api.api_post().text
        # 把每次结果都添加至列表
        answer.append(result)
        # 判断是否与预期结果相符
        if result == expect_result:
            pass_or_false.append('Pass')
        else:
            pass_or_false.append('False')
    
    # 把结果写入新的表格
    document_data.write_result()
    # 结束

