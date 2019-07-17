import os
import os.path

def createTask(pyName, funName, xml):
    try:
        import time
        dir_path = os.path.dirname(os.path.abspath(__file__)) + '\\Task\\'
        now = time.strftime('%Y%m%d_%H%M%S', time.localtime())  #将指定格式的当前时间以字符串输出
        suffix = ".py"
        newfile= dir_path + now + suffix  #新建任务的绝对路径
        if not os.path.exists(dir_path):
            os.mkdir(dir_path)
        if not os.path.exists(newfile):
            Task = open(newfile, 'w+')
            Task.write('import FileTranslator.' + pyName + ' as PyPath \n')
            Task.write('xml = r"' + xml + '" \n')
            Task.write('dateId = r"' + now + '" \n')
            Task.write('PyPath.' + funName + '(xml, dateId) \n')
            Task.close()
            print(newfile + " created.")
        else:
            print(newfile + " already existed.")
            Task = open(newfile, 'w+')
            Task.write('import FileTranslator.' + pyName + ' as PyPath \n')
            Task.write('xml = r"' + xml + '" \n')
            Task.write('dateId = r"' + now + '" \n')
            Task.write('PyPath.' + funName + '(xml, dateId) \n')
            Task.close()
        with open(newfile, encoding="UTF-8") as f:
            exec(f.read())
        return newfile
    except ValueError as e:
        print(e.args)
        return e.args