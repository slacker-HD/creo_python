# 系统配置：

1. 安装vbapi并配置好。

2. 安装python3。

3. 安装pywin32：

```bash
pip install pywin32
```

4. 配置pywin32.  

```bash
cd ~\AppData\Local\Programs\Python\Python37\Lib\site-packages\win32com\client
python makepy.py
```

选择vbapi组件生成。提示 Generating to C:\Users\hudi\AppData\Local\Temp\gen_py\3.7\176453F2-6934-4304-8C9D-126D98C1700Ex0x1x0.py

在上面的py文件里面找enum对应的constants类，记录了所有的enum等。  
类对应的uuid也在上面有，client.Dispatch(uuid)即可生成对象。  
类的继承很有意思，父-子-另一个父的方法直接就可以用，不需要转换。  

## RelationOperation.py

- 批量添加关系
- 批量清空关系

## Familyexport.py

- 导出文件族表到单个实例文件
