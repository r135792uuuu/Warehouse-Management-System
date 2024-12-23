# Warehouse-Management-System
仓库管理系统，开发工具python，开发平台windows

2024年12月18日 19:30

0.2.0

完成主要功能的开发，初步实现增删查改，借还状态，查找财产

**ADD：**
 - 界面UI史诗级更新

**TODO：**
- 查询功能：实现具体的查询某一次借用物品状态，然后看后续流程
- 数据库：接入用户和管理员之后需要实时读取在线数据库
- 按照飞书知识库继续修改

**FIX：**
 - 增加物品的功能空值bug
 - 存放位置的下拉菜单读取问题
 - 增加更多下拉菜单
 - - 名字也可以整个好看的菜单
 - - 存放位置也需要整一下。最好是输入物品之后，自动查找所有存放的位置

**SUPPLY：**
 - 增加物品的功能说明：目前增加物品的合并逻辑是判断大类名称，小类名称，备注三者相同就合并。
 名称识别都没问题，但是备注有一点需要注意。
 - - 若是备注加入仓库的时候就是有字（e.g. 坏的），那么每次新打开软件或者已经打开软件跑过新增之后都会成功加入。
 - - 若是备注是没有字的，那么每次新打开软件，然后我加入一个已有的物品就无法正常合并。但是加入这一次之后，软件开着
 我再加入一次这个物品，那么就可以正常识别。很奇怪，我的空值检测不知道是不是因为中文的问题。暂时再说，bug后面想办法。